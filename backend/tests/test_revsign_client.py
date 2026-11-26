"""Testes do cliente do RevSign (interface, falso e HTTP) -- sem rede."""
import hashlib
import hmac
import json
from datetime import datetime, timezone

import httpx
import pytest

from app.services.revsign_client import (
    PDF_FAKE,
    DocumentoCriado,
    EventoWebhook,
    FakeRevSignClient,
    HttpRevSignClient,
    NovoDocumento,
    RevSignError,
    WebhookInvalido,
    WebhookMalformado,
    assinar_webhook,
)

pytestmark = pytest.mark.unit

BASE = "https://rs.teste/api/revalle/patrimonio/v1"
TOKEN = "tok-123"


def _novo_doc(**extra) -> NovoDocumento:
    dados = dict(
        titulo="Termo de entrega — Notebook Dell Latitude (patrimônio #5)",
        arquivo=b"%PDF-1.7 conteudo",
        nome_arquivo="termo.pdf",
        signatario_nome="Maria Souza",
        signatario_email="maria.souza@gmail.com",
        referencia_externa="patrimonio-termo-5",
        tipo="ENTREGA",
        signatario_cpf="52998224725",
    )
    dados.update(extra)
    return NovoDocumento(**dados)


def _cliente(handler, requisicoes: list | None = None) -> HttpRevSignClient:
    def _h(request: httpx.Request) -> httpx.Response:
        if requisicoes is not None:
            requisicoes.append(request)
        return handler(request)

    return HttpRevSignClient(BASE, TOKEN, "segredo", transport=httpx.MockTransport(_h))


def _campos_multipart(request: httpx.Request) -> dict[str, str]:
    """Campos de texto do multipart (nome -> valor); o arquivo fica de fora."""
    boundary = request.headers["Content-Type"].split("boundary=")[1].encode()
    campos = {}
    for parte in request.content.split(b"--" + boundary):
        if b'name="' not in parte or b'filename="' in parte:
            continue
        cabecalho, _, valor = parte.partition(b"\r\n\r\n")
        nome = cabecalho.split(b'name="')[1].split(b'"')[0].decode()
        campos[nome] = valor.rstrip(b"\r\n").decode()
    return campos


# ── Criar ─────────────────────────────────────────────────────────────────────


def test_criar_envia_multipart_do_contrato():
    reqs: list[httpx.Request] = []
    cliente = _cliente(lambda r: httpx.Response(201, json={"termoId": "envelope_x", "status": "PENDING"}), reqs)
    assert cliente.criar_documento(_novo_doc()) == DocumentoCriado("envelope_x")

    assert len(reqs) == 1
    req = reqs[0]
    assert req.method == "POST"
    assert str(req.url) == f"{BASE}/termos"
    assert req.headers["Authorization"] == f"Bearer {TOKEN}"
    assert req.headers["Content-Type"].startswith("multipart/form-data")
    assert _campos_multipart(req) == {
        "externalId": "patrimonio-termo-5",
        "kind": "ENTREGA",
        "title": "Termo de entrega — Notebook Dell Latitude (patrimônio #5)",
        "signerName": "Maria Souza",
        "signerEmail": "maria.souza@gmail.com",
        "signerCpf": "52998224725",
        "message": "Revalle · Patrimônio: assine o termo do equipamento sob sua responsabilidade.",
    }
    assert b'name="file"; filename="termo.pdf"' in req.content
    assert b"Content-Type: application/pdf" in req.content
    assert b"%PDF-1.7 conteudo" in req.content


def test_criar_devolucao_envia_kind_devolucao():
    reqs: list[httpx.Request] = []
    cliente = _cliente(lambda r: httpx.Response(201, json={"termoId": "envelope_x"}), reqs)
    cliente.criar_documento(_novo_doc(tipo="DEVOLUCAO"))
    assert _campos_multipart(reqs[0])["kind"] == "DEVOLUCAO"


def test_criar_duplicado_reaproveita_termo():
    """Falha de rede na criação + repetição: o 409 com o termoId existente vale como sucesso."""
    reqs: list[httpx.Request] = []
    resposta = {"error": {"code": "DUPLICATE_EXTERNAL_ID", "message": "já existe", "termoId": "envelope_y"}}
    cliente = _cliente(lambda r: httpx.Response(409, json=resposta), reqs)
    assert cliente.criar_documento(_novo_doc()) == DocumentoCriado("envelope_y")
    assert len(reqs) == 1  # não cria outro


def test_criar_depois_de_falha_de_rede_repete_e_reaproveita():
    """Timeout depois de o RevSign criar o termo: a repetição automática recebe o 409 com o
    termoId existente (contrato §2), e o termo não fica órfão."""
    chamadas = []

    def handler(request):
        chamadas.append(request)
        if len(chamadas) == 1:
            raise httpx.ReadTimeout("x")
        return httpx.Response(409, json={"error": {"code": "DUPLICATE_EXTERNAL_ID", "termoId": "envelope_z"}})

    cliente = _cliente(handler)
    assert cliente.criar_documento(_novo_doc()) == DocumentoCriado("envelope_z")
    assert len(chamadas) == 2


def test_criar_repete_uma_vez_apos_falha_de_transporte():
    chamadas = []

    def handler(request):
        chamadas.append(request)
        if len(chamadas) == 1:
            raise httpx.ConnectError("x")
        return httpx.Response(201, json={"termoId": "envelope_w", "status": "PENDING"})

    assert _cliente(handler).criar_documento(_novo_doc()) == DocumentoCriado("envelope_w")
    assert len(chamadas) == 2
    assert _campos_multipart(chamadas[0]) == _campos_multipart(chamadas[1])  # o mesmo pedido


def test_criar_falha_de_transporte_nas_duas_tentativas_levanta():
    chamadas = []

    def handler(request):
        chamadas.append(request)
        raise httpx.ReadTimeout("x")

    with pytest.raises(RevSignError) as exc:
        _cliente(handler).criar_documento(_novo_doc())
    assert str(exc.value) == "RevSign indisponível."
    assert exc.value.codigo is None
    assert len(chamadas) == 2  # só uma repetição


def test_criar_erro_com_codigo_nao_repete():
    chamadas = []

    def handler(request):
        chamadas.append(request)
        return httpx.Response(422, json={"error": {"code": "INVALID_CPF"}})

    with pytest.raises(RevSignError):
        _cliente(handler).criar_documento(_novo_doc())
    assert len(chamadas) == 1


def test_criar_duplicado_sem_termo_id_levanta():
    cliente = _cliente(lambda r: httpx.Response(409, json={"error": {"code": "DUPLICATE_EXTERNAL_ID"}}))
    with pytest.raises(RevSignError) as exc:
        cliente.criar_documento(_novo_doc())
    assert exc.value.codigo == "DUPLICATE_EXTERNAL_ID"


@pytest.mark.parametrize(
    "status, codigo, mensagem",
    [
        (401, "UNAUTHORIZED", "Chave da integração com o RevSign recusada."),
        (404, "NOT_CONFIGURED", "Integração desligada no RevSign."),
        (422, "INVALID_CPF", "CPF do colaborador recusado pelo RevSign."),
        (422, "COLABORADOR_NOT_FOUND", "Colaborador não está na base do RH do RevSign."),
        (503, "HR_BASE_UNAVAILABLE", "Base do RH do RevSign fora do ar."),
        (422, "MARKER_NOT_FOUND", "O termo não tem o marcador de assinatura."),
        (422, "INVALID_FILE", "O RevSign recusou o PDF do termo."),
        (413, "FILE_TOO_LARGE", "PDF do termo maior que 10 MB."),
        (429, "RESEND_TOO_SOON", "Aguarde 10 minutos entre reenvios."),
        (409, "NOT_PENDING", "O termo já não está pendente no RevSign."),
        (429, "RATE_LIMITED", "RevSign ocupado; tente de novo em instantes."),
        (400, "INVALID_REQUEST", "O RevSign recusou os dados do termo."),
    ],
)
def test_criar_erro_vira_mensagem(status, codigo, mensagem):
    # a mensagem do servidor (inclusive um eventual token) nunca vai para o usuário
    corpo = {"error": {"code": codigo, "message": f"detalhe com {TOKEN}"}}
    cliente = _cliente(lambda r: httpx.Response(status, json=corpo))
    with pytest.raises(RevSignError) as exc:
        cliente.criar_documento(_novo_doc())
    assert str(exc.value) == mensagem
    assert exc.value.codigo == codigo
    assert TOKEN not in str(exc.value)


@pytest.mark.parametrize(
    "resposta",
    [
        httpx.Response(500, json={"error": {"code": "INTERNAL_ERROR", "message": "x"}}),
        httpx.Response(502, content=b"<html>bad gateway</html>"),
        httpx.Response(500, json=["lista"]),
    ],
)
def test_erro_desconhecido_ou_sem_corpo_usa_o_status(resposta):
    cliente = _cliente(lambda r: resposta)
    with pytest.raises(RevSignError) as exc:
        cliente.consultar("envelope_x")
    assert str(exc.value) == f"RevSign respondeu {resposta.status_code}."


def test_revsign_error_codigo_e_opcional():
    assert RevSignError("msg").codigo is None
    assert RevSignError("msg", "X").codigo == "X"
    assert str(RevSignError("msg", "X")) == "msg"


# ── Consultar, cancelar, reenviar, baixar ────────────────────────────────────


@pytest.mark.parametrize(
    "termo, situacao, motivo, assinado",
    [
        (
            {"status": "SIGNED", "completedAt": "2026-11-12T13:05:10.000Z",
             "signer": {"signedAt": "2026-11-12T13:05:00.000Z"}},
            "assinado", None, datetime(2026, 11, 12, 13, 5, 0, tzinfo=timezone.utc),
        ),
        (
            {"status": "SIGNED", "completedAt": "2026-11-12T13:05:10.000Z", "signer": {"signedAt": None}},
            "assinado", None, datetime(2026, 11, 12, 13, 5, 10, tzinfo=timezone.utc),
        ),
        (
            {"status": "REJECTED", "signer": {"rejectionReason": "Dados errados"}},
            "recusado", "Dados errados", None,
        ),
        ({"status": "EXPIRED", "signer": {"expiresAt": "2026-11-01T00:00:00.000Z"}}, "expirado", None, None),
        ({"status": "CANCELLED", "cancelReason": "trocou"}, "cancelado", None, None),
        ({"status": "PENDING", "signer": {"signedAt": None}}, "pendente", None, None),
    ],
)
def test_consultar_mapeia_status(termo, situacao, motivo, assinado):
    reqs: list[httpx.Request] = []
    cliente = _cliente(lambda r: httpx.Response(200, json={"termoId": "envelope_x", **termo}), reqs)
    status = cliente.consultar("envelope_x")
    assert status.situacao == situacao
    assert status.motivo == motivo
    assert status.assinado_em == assinado
    assert reqs[0].method == "GET"
    assert str(reqs[0].url) == f"{BASE}/termos/envelope_x"
    assert reqs[0].headers["Authorization"] == f"Bearer {TOKEN}"


def test_consultar_status_desconhecido_levanta():
    cliente = _cliente(lambda r: httpx.Response(200, json={"status": "ALGO_NOVO"}))
    with pytest.raises(RevSignError, match="Resposta inválida do RevSign."):
        cliente.consultar("envelope_x")


def test_cancelar_envia_motivo_e_aceita_404():
    reqs: list[httpx.Request] = []
    cliente = _cliente(lambda r: httpx.Response(200, json={"status": "CANCELLED"}), reqs)
    cliente.cancelar("envelope_x", "Item trocado antes da assinatura")
    assert reqs[0].method == "POST"
    assert str(reqs[0].url) == f"{BASE}/termos/envelope_x/cancelar"
    assert json.loads(reqs[0].content) == {"reason": "Item trocado antes da assinatura"}

    nao_existe = {"error": {"code": "NOT_FOUND", "message": "x"}}
    _cliente(lambda r: httpx.Response(404, json=nao_existe)).cancelar("envelope_x")  # não levanta

    # 404 de integração desligada NÃO é "já não existe"
    desligada = {"error": {"code": "NOT_CONFIGURED", "message": "x"}}
    with pytest.raises(RevSignError) as exc:
        _cliente(lambda r: httpx.Response(404, json=desligada)).cancelar("envelope_x")
    assert exc.value.codigo == "NOT_CONFIGURED"


def test_cancelar_sem_motivo_nao_manda_corpo():
    reqs: list[httpx.Request] = []
    _cliente(lambda r: httpx.Response(200, json={}), reqs).cancelar("envelope_x")
    assert reqs[0].content == b""


def test_cancelar_nao_pendente_levanta():
    corpo = {"error": {"code": "NOT_PENDING", "message": "x", "status": "SIGNED"}}
    with pytest.raises(RevSignError) as exc:
        _cliente(lambda r: httpx.Response(409, json=corpo)).cancelar("envelope_x")
    assert str(exc.value) == "O termo já não está pendente no RevSign."
    assert exc.value.codigo == "NOT_PENDING"


def test_reenviar_ok():
    reqs: list[httpx.Request] = []
    _cliente(lambda r: httpx.Response(200, json={"status": "PENDING"}), reqs).reenviar("envelope_x")
    assert reqs[0].method == "POST"
    assert str(reqs[0].url) == f"{BASE}/termos/envelope_x/reenviar"


def test_reenviar_429_mensagem():
    corpo = {"error": {"code": "RESEND_TOO_SOON", "message": "x"}}
    cliente = _cliente(lambda r: httpx.Response(429, json=corpo, headers={"Retry-After": "300"}))
    with pytest.raises(RevSignError) as exc:
        cliente.reenviar("envelope_x")
    assert str(exc.value) == "Aguarde 10 minutos entre reenvios."
    assert exc.value.codigo == "RESEND_TOO_SOON"


def test_baixar_pdf():
    reqs: list[httpx.Request] = []
    cliente = _cliente(lambda r: httpx.Response(200, content=b"%PDF-1.7 conteudo"), reqs)
    assert cliente.baixar_assinado("envelope_x") == b"%PDF-1.7 conteudo"
    assert reqs[0].method == "GET"
    assert str(reqs[0].url) == f"{BASE}/termos/envelope_x/pdf"

    cliente = _cliente(lambda r: httpx.Response(200, content=b"<html>"))
    with pytest.raises(RevSignError):
        cliente.baixar_assinado("envelope_x")


@pytest.mark.parametrize("transforma", [str.lower, str.upper])
def test_baixar_confere_o_sha256_anunciado(transforma, caplog):
    import logging

    conteudo = b"%PDF-1.7 conteudo selado"
    hash_ok = transforma(hashlib.sha256(conteudo).hexdigest())
    cliente = _cliente(lambda r: httpx.Response(200, content=conteudo, headers={"X-RevSign-SHA256": hash_ok}))
    with caplog.at_level(logging.INFO, logger="app.services.revsign_client"):
        assert cliente.baixar_assinado("envelope_x") == conteudo
    assert any("envelope_x" in r.getMessage() and hash_ok.lower() in r.getMessage() for r in caplog.records)


def test_baixar_com_sha256_diferente_levanta():
    conteudo = b"%PDF-1.7 conteudo selado"
    outro = hashlib.sha256(b"outro").hexdigest()
    cliente = _cliente(lambda r: httpx.Response(200, content=conteudo, headers={"X-RevSign-SHA256": outro}))
    with pytest.raises(RevSignError) as exc:
        cliente.baixar_assinado("envelope_x")
    assert str(exc.value) == "O PDF assinado chegou corrompido do RevSign."


def test_baixar_sem_cabecalho_sha256_aceita():
    cliente = _cliente(lambda r: httpx.Response(200, content=b"%PDF-1.7 x"))
    assert cliente.baixar_assinado("envelope_x") == b"%PDF-1.7 x"


def test_baixar_nao_assinado_levanta_com_codigo():
    corpo = {"error": {"code": "NOT_SIGNED", "message": "x"}}
    with pytest.raises(RevSignError) as exc:
        _cliente(lambda r: httpx.Response(409, json=corpo)).baixar_assinado("envelope_x")
    assert exc.value.codigo == "NOT_SIGNED"
    assert str(exc.value) == "RevSign respondeu 409."


def test_baixar_assinado_recusa_pdf_grande(monkeypatch):
    import app.services.revsign_client as mod

    monkeypatch.setattr(mod.HttpRevSignClient, "MAX_PDF_BYTES", 10)
    cliente = _cliente(lambda r: httpx.Response(200, content=b"%PDF-1.7 " + b"x" * 50))
    with pytest.raises(RevSignError):
        cliente.baixar_assinado("envelope_x")


def test_erro_de_rede_vira_revsign_error_sem_token():
    def handler(request):
        raise httpx.ConnectError("x")

    with pytest.raises(RevSignError) as exc:
        _cliente(handler).consultar("envelope_x")
    assert TOKEN not in str(exc.value)

    with pytest.raises(RevSignError) as exc:
        _cliente(lambda r: httpx.Response(401, json={"message": "Bearer tok-123 invalido"})).consultar("envelope_x")
    assert TOKEN not in str(exc.value)


def test_sem_configuracao_levanta():
    with pytest.raises(RevSignError) as exc:
        HttpRevSignClient("", "", "s").criar_documento(_novo_doc())
    assert str(exc.value) == "RevSign não configurado (REVSIGN_BASE_URL/REVSIGN_API_TOKEN)."


# ── Webhook ───────────────────────────────────────────────────────────────────

AGORA = 1_800_000_000
CORPO_OK = json.dumps(
    {
        "eventId": "evt_1",
        "event": "termo.signed",
        "termoId": "envelope_x",
        "externalId": "patrimonio-termo-5",
        "status": "SIGNED",
        "occurredAt": "2026-11-12T13:05:10.123Z",
        "signer": {"name": "Maria Souza", "email": "maria.souza@gmail.com"},
    }
).encode()
EVENTO_OK = EventoWebhook("termo.signed", "envelope_x", "patrimonio-termo-5")


@pytest.fixture
def relogio_fixo(monkeypatch):
    import app.services.revsign_client as mod

    monkeypatch.setattr(mod.time, "time", lambda: float(AGORA))


def _assinado(segredo="s3", corpo=CORPO_OK, ts=AGORA):
    return {**assinar_webhook(segredo, corpo, ts), "Content-Type": "application/json"}


def _assinado_bruto(timestamp: str, segredo="s3", corpo=CORPO_OK):
    """Assina com o timestamp exatamente como vai no cabeçalho (ex.: com zeros à esquerda)."""
    mac = hmac.new(segredo.encode(), timestamp.encode() + b"." + corpo, hashlib.sha256).hexdigest()
    return {"X-RevSign-Timestamp": timestamp, "X-RevSign-Signature": f"sha256={mac}"}


def _sem(cabecalhos, nome):
    return {k: v for k, v in cabecalhos.items() if k != nome}


def test_assinar_webhook_formato_do_contrato():
    cab = assinar_webhook("s3", b'{"a":1}', 1234)
    esperado = hmac.new(b"s3", b'1234.{"a":1}', hashlib.sha256).hexdigest()
    assert cab == {"X-RevSign-Timestamp": "1234", "X-RevSign-Signature": f"sha256={esperado}"}


@pytest.mark.parametrize(
    "segredo_servidor, cabecalhos, corpo, esperado",
    [
        ("s3", _assinado(), CORPO_OK, EVENTO_OK),
        ("s3", {k.lower(): v for k, v in _assinado().items()}, CORPO_OK, EVENTO_OK),
        ("s3", _assinado(ts=AGORA - 300), CORPO_OK, EVENTO_OK),  # limite da janela
        ("s3", _assinado(ts=AGORA + 300), CORPO_OK, EVENTO_OK),
        ("s3", _assinado(segredo="errado"), CORPO_OK, WebhookInvalido),
        ("s3", {}, CORPO_OK, WebhookInvalido),
        ("s3", _sem(_assinado(), "X-RevSign-Signature"), CORPO_OK, WebhookInvalido),
        ("s3", _sem(_assinado(), "X-RevSign-Timestamp"), CORPO_OK, WebhookInvalido),
        (
            "s3",
            {**_assinado(), "X-RevSign-Signature": _assinado()["X-RevSign-Signature"].removeprefix("sha256=")},
            CORPO_OK,
            WebhookInvalido,
        ),
        ("s3", _assinado(ts=AGORA - 301), CORPO_OK, WebhookInvalido),
        ("s3", _assinado(ts=AGORA + 301), CORPO_OK, WebhookInvalido),
        ("s3", {**_assinado(), "X-RevSign-Timestamp": "abc"}, CORPO_OK, WebhookInvalido),
        ("s3", {**_assinado(), "X-RevSign-Timestamp": "9" * 13}, CORPO_OK, WebhookInvalido),
        ("s3", {**_assinado(), "X-RevSign-Timestamp": "9" * 400}, CORPO_OK, WebhookInvalido),
        ("s3", {**_assinado(), "X-RevSign-Timestamp": "9" * 5000}, CORPO_OK, WebhookInvalido),
        ("s3", _assinado_bruto(f"0{AGORA}"), CORPO_OK, EVENTO_OK),  # HMAC sobre o texto do cabeçalho
        ("s3", _assinado(), CORPO_OK + b" ", WebhookInvalido),  # corpo adulterado
        ("", _assinado(segredo=""), CORPO_OK, WebhookInvalido),
        ("s3", _assinado(corpo=b"nao-json"), b"nao-json", WebhookMalformado),
        ("s3", _assinado(corpo=b"[]"), b"[]", WebhookMalformado),
        ("s3", _assinado(corpo=b'{"termoId": "x"}'), b'{"termoId": "x"}', WebhookMalformado),
    ],
)
def test_validar_webhook(relogio_fixo, segredo_servidor, cabecalhos, corpo, esperado):
    for cliente in (
        FakeRevSignClient(webhook_secret=segredo_servidor),
        HttpRevSignClient(BASE, TOKEN, segredo_servidor),
    ):
        if isinstance(esperado, EventoWebhook):
            assert cliente.validar_webhook(cabecalhos, corpo) == esperado
        else:
            with pytest.raises(esperado):
                cliente.validar_webhook(cabecalhos, corpo)


def test_validar_webhook_sem_termo_nem_referencia(relogio_fixo):
    corpo = b'{"event": "termo.sent"}'
    ev = HttpRevSignClient(BASE, TOKEN, "s3").validar_webhook(_assinado(corpo=corpo), corpo)
    assert ev == EventoWebhook("termo.sent", None, None)


def test_webhook_malformado_nao_vaza_segredo_no_tipo():
    # WebhookInvalido não é WebhookMalformado: o router distingue 401 de 400.
    assert not issubclass(WebhookInvalido, WebhookMalformado)
    assert issubclass(WebhookInvalido, RevSignError)


# ── Falso ─────────────────────────────────────────────────────────────────────


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
