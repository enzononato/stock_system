"""Testes do cliente do RevSign (interface, falso e HTTP) -- sem rede."""
import json
from datetime import datetime, timedelta, timezone

import httpx
import pytest

from app.services.revsign_client import (
    API_PREFIXO,
    PDF_FAKE,
    DocumentoCriado,
    EventoWebhook,
    FakeRevSignClient,
    HttpRevSignClient,
    NovoDocumento,
    RevSignError,
    WebhookInvalido,
    WebhookMalformado,
)

pytestmark = pytest.mark.unit

BASE = "https://rs.teste"
TOKEN = "tok-123"


def _novo_doc(**extra) -> NovoDocumento:
    dados = dict(
        titulo="Termo de entrega — Notebook Dell Latitude (patrimônio #5)",
        arquivo=b"PK\x03\x04conteudo",
        nome_arquivo="termo.docx",
        signatario_nome="Maria Souza",
        signatario_email="maria.souza@gmail.com",
        referencia_externa="patrimonio-termo-5",
    )
    dados.update(extra)
    return NovoDocumento(**dados)


def _cliente(handler, requisicoes: list | None = None) -> HttpRevSignClient:
    def _h(request: httpx.Request) -> httpx.Response:
        if requisicoes is not None:
            requisicoes.append(request)
        return handler(request)

    return HttpRevSignClient(BASE, TOKEN, "segredo", transport=httpx.MockTransport(_h))


def test_criar_documento_envia_multipart_e_distribui():
    reqs: list[httpx.Request] = []

    def handler(request):
        if request.url.path.endswith("/document/create"):
            return httpx.Response(200, json={"envelopeId": "env_1", "id": 42})
        return httpx.Response(200, json={})

    cliente = _cliente(handler, reqs)
    assert cliente.criar_documento(_novo_doc()) == DocumentoCriado("42")

    criar, distribuir = reqs[0], reqs[1]
    assert criar.method == "POST"
    assert str(criar.url) == f"{BASE}{API_PREFIXO}/document/create"
    assert criar.headers["Authorization"] == f"Bearer {TOKEN}"
    corpo = criar.content
    assert b'name="payload"' in corpo
    assert b'name="file"; filename="termo.docx"' in corpo
    assert '"externalId": "patrimonio-termo-5"'.encode() in corpo
    assert b'"role": "SIGNER"' in corpo
    assert b'"type": "SIGNATURE"' in corpo
    assert '"email": "maria.souza@gmail.com"'.encode() in corpo
    assert b'"language": "pt-BR"' in corpo

    assert distribuir.method == "POST"
    assert str(distribuir.url) == f"{BASE}{API_PREFIXO}/document/distribute"
    assert json.loads(distribuir.content) == {"documentId": 42}


def test_distribuir_falha_apaga_rascunho_e_levanta():
    reqs: list[httpx.Request] = []

    def handler(request):
        if request.url.path.endswith("/document/create"):
            return httpx.Response(200, json={"envelopeId": "env_1", "id": 42})
        if request.url.path.endswith("/document/distribute"):
            return httpx.Response(500, json={"message": "erro"})
        return httpx.Response(200, json={})

    cliente = _cliente(handler, reqs)
    with pytest.raises(RevSignError):
        cliente.criar_documento(_novo_doc())
    assert len(reqs) == 3
    assert str(reqs[2].url) == f"{BASE}{API_PREFIXO}/document/delete"
    assert json.loads(reqs[2].content) == {"documentId": 42}


@pytest.mark.parametrize(
    "resposta, situacao, motivo",
    [
        ({"status": "COMPLETED", "completedAt": "2026-11-12T13:00:00.000Z"}, "assinado", None),
        (
            {"status": "REJECTED", "recipients": [{"id": 1, "rejectionReason": "Dados errados"}]},
            "recusado",
            "Dados errados",
        ),
        ({"status": "CANCELLED"}, "cancelado", None),
        ({"status": "PENDING", "recipients": [{"id": 1, "signingStatus": "NOT_SIGNED"}]}, "pendente", None),
        (
            {
                "status": "PENDING",
                "recipients": [
                    {
                        "id": 1,
                        "signingStatus": "NOT_SIGNED",
                        "expiresAt": (datetime.now(timezone.utc) - timedelta(days=1)).isoformat(),
                    }
                ],
            },
            "expirado",
            None,
        ),
    ],
)
def test_consultar_mapeia_status(resposta, situacao, motivo):
    reqs: list[httpx.Request] = []
    cliente = _cliente(lambda r: httpx.Response(200, json=resposta), reqs)
    status = cliente.consultar("42")
    assert status.situacao == situacao
    assert status.motivo == motivo
    assert reqs[0].method == "GET"
    assert str(reqs[0].url) == f"{BASE}{API_PREFIXO}/document/42"
    if situacao == "assinado":
        assert status.assinado_em is not None


def test_reenviar_redistribui_so_quem_nao_assinou():
    reqs: list[httpx.Request] = []

    def handler(request):
        if request.method == "GET":
            return httpx.Response(
                200,
                json={
                    "status": "PENDING",
                    "recipients": [
                        {"id": 7, "signingStatus": "NOT_SIGNED"},
                        {"id": 8, "signingStatus": "SIGNED"},
                    ],
                },
            )
        return httpx.Response(200, json={})

    _cliente(handler, reqs).reenviar("42")
    post = reqs[-1]
    assert post.method == "POST"
    assert str(post.url) == f"{BASE}{API_PREFIXO}/document/redistribute"
    assert json.loads(post.content) == {"documentId": 42, "recipients": [7]}


def test_cancelar_trata_404_como_sucesso():
    cliente = _cliente(lambda r: httpx.Response(404, json={"message": "nao existe"}))
    cliente.cancelar("42")  # não levanta


def test_baixar_assinado_exige_pdf():
    reqs: list[httpx.Request] = []
    cliente = _cliente(lambda r: httpx.Response(200, content=b"%PDF-1.7 conteudo"), reqs)
    assert cliente.baixar_assinado("42") == b"%PDF-1.7 conteudo"
    assert str(reqs[0].url) == f"{BASE}{API_PREFIXO}/document/42/download"

    cliente = _cliente(lambda r: httpx.Response(200, content=b"<html>"))
    with pytest.raises(RevSignError):
        cliente.baixar_assinado("42")


def test_baixar_assinado_recusa_pdf_grande(monkeypatch):
    import app.services.revsign_client as mod

    monkeypatch.setattr(mod.HttpRevSignClient, "MAX_PDF_BYTES", 10)
    cliente = _cliente(lambda r: httpx.Response(200, content=b"%PDF-1.7 " + b"x" * 50))
    with pytest.raises(RevSignError):
        cliente.baixar_assinado("42")


def test_erro_de_rede_vira_revsign_error_sem_token():
    def handler(request):
        raise httpx.ConnectError("x")

    with pytest.raises(RevSignError) as exc:
        _cliente(handler).consultar("42")
    assert TOKEN not in str(exc.value)

    with pytest.raises(RevSignError) as exc:
        _cliente(lambda r: httpx.Response(401, json={"message": "Bearer tok-123 invalido"})).consultar("42")
    assert TOKEN not in str(exc.value)


def test_sem_configuracao_levanta():
    with pytest.raises(RevSignError) as exc:
        HttpRevSignClient("", "", "s").criar_documento(_novo_doc())
    assert str(exc.value) == "RevSign não configurado (REVSIGN_BASE_URL/REVSIGN_API_TOKEN)."


CORPO_OK = json.dumps(
    {"event": "DOCUMENT_COMPLETED", "payload": {"id": 42, "externalId": "patrimonio-termo-5"}}
).encode()


@pytest.mark.parametrize(
    "segredo_servidor, cabecalhos, corpo, esperado",
    [
        ("s3", {"X-Documenso-Secret": "s3"}, CORPO_OK, EventoWebhook("DOCUMENT_COMPLETED", "42", "patrimonio-termo-5")),
        ("s3", {"x-documenso-secret": "s3"}, CORPO_OK, EventoWebhook("DOCUMENT_COMPLETED", "42", "patrimonio-termo-5")),
        ("s3", {"X-Documenso-Secret": "errado"}, CORPO_OK, WebhookInvalido),
        ("s3", {}, CORPO_OK, WebhookInvalido),
        ("", {"X-Documenso-Secret": ""}, CORPO_OK, WebhookInvalido),
        ("s3", {"X-Documenso-Secret": "s3"}, b"nao-json", WebhookMalformado),
        ("s3", {"X-Documenso-Secret": "s3"}, b"[]", WebhookMalformado),
        ("s3", {"X-Documenso-Secret": "s3"}, b'{"payload": {}}', WebhookMalformado),
    ],
)
def test_validar_webhook(segredo_servidor, cabecalhos, corpo, esperado):
    for cliente in (FakeRevSignClient(webhook_secret=segredo_servidor), HttpRevSignClient(BASE, TOKEN, segredo_servidor)):
        if isinstance(esperado, EventoWebhook):
            assert cliente.validar_webhook(cabecalhos, corpo) == esperado
        else:
            with pytest.raises(esperado):
                cliente.validar_webhook(cabecalhos, corpo)


def test_webhook_malformado_nao_vaza_segredo_no_tipo():
    # WebhookInvalido não é WebhookMalformado: o router distingue 401 de 400.
    assert not issubclass(WebhookInvalido, WebhookMalformado)
    assert issubclass(WebhookInvalido, RevSignError)


def test_fake_ciclo():
    fake = FakeRevSignClient()
    criado = fake.criar_documento(_novo_doc())
    assert criado.documento_id == "1001"
    assert fake.consultar("1001").situacao == "pendente"
    with pytest.raises(RevSignError, match="Documento ainda não assinado."):
        fake.baixar_assinado("1001")
    fake.marcar_assinado("1001")
    assert fake.consultar("1001").situacao == "assinado"
    assert fake.baixar_assinado("1001") == PDF_FAKE
    with pytest.raises(RevSignError, match="Documento não encontrado no RevSign."):
        fake.consultar("9999")
    fake.cancelar("1001")
    fake.reenviar("1001")
    assert fake.cancelados == ["1001"] and fake.reenviados == ["1001"]
    assert fake.criar_documento(_novo_doc()).documento_id == "1002"

    fake.falhar_em = {"criar_documento"}
    with pytest.raises(RevSignError, match="Falha simulada em criar_documento."):
        fake.criar_documento(_novo_doc())
