"""Retorno do RevSign: webhook autenticado e sincronização idempotente das assinaturas."""
from datetime import datetime

import pytest

from app.services.revsign_client import PDF_FAKE
from tests.fixtures_termos import (  # noqa: F401
    SEGREDO_WEBHOOK,
    colaborador,
    corpo_emprestimo_colaborador,
    criar_colaborador,
    evento_webhook,
    linhas_audit,
    postar_webhook,
    revsign_fake,
    unidade_juazeiro,
)

pytestmark = pytest.mark.integration


def _emprestar(client_gestor, item_id, colaborador_id):
    resp = client_gestor.post("/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador_id))
    assert resp.status_code == 200, resp.text
    return resp.json()["termo"]


def _termo(termo_id):
    from app.dependencies import get_termo_db

    return get_termo_db().obter(termo_id)


def _historico(item_id, operacao):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT * FROM history WHERE item_id=%s AND operation=%s ORDER BY id", (item_id, operacao))
        return list(cur.fetchall())


def _item(client, item_id):
    return client.get(f"/api/items/{item_id}").json()


def test_webhook_assinado_conclui_emprestimo_e_grava_pdf(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    from app.core.storage import storage

    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    revsign_fake.marcar_assinado("1001")

    resp = postar_webhook(client, evento_webhook("termo.signed", "1001", termo["id"]))  # SEM login
    assert resp.status_code == 200 and resp.json() == {"ok": True, "resultado": "assinado"}
    assert _item(client_gestor, item_id)["status"] == "Indisponível"

    confirmacoes = _historico(item_id, "Confirmação Empréstimo")
    assert len(confirmacoes) == 1
    assert confirmacoes[0]["operador"] == "RevSign"
    chave = confirmacoes[0]["termo_assinado_anexo"]
    assert chave.startswith("termos_assinados/")
    assert storage.read(chave) == PDF_FAKE

    atual = _termo(termo["id"])
    assert atual["status"] == "assinado" and atual["assinado_em"] is not None
    assert atual["arquivo_assinado_key"] == chave
    for entidade, acao in (("emprestimo", "confirmar"), ("termo", "assinar")):
        linhas = linhas_audit(entidade, acao)
        assert len(linhas) == 1 and linhas[0]["username"] == "sistema"
    assert linhas_audit("termo", "assinar")[0]["depois"]["origem"] == "revsign"


def test_mesmo_evento_duas_vezes_confirma_uma_vez(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    revsign_fake.marcar_assinado("1001")
    evento = evento_webhook("termo.signed", "1001", termo["id"])

    assert postar_webhook(client, evento).json()["resultado"] == "assinado"
    assert postar_webhook(client, evento).json()["resultado"] == "ja_processado"
    assert len(_historico(item_id, "Confirmação Empréstimo")) == 1
    assert len(linhas_audit("emprestimo", "confirmar")) == 1


def test_trava_ocupada_nao_processa(client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake):
    from app.db.database_mysql import get_cursor
    from app.services import termos_retorno

    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    revsign_fake.marcar_assinado("1001")

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT GET_LOCK(%s, 0)", (termos_retorno.nome_da_trava(termo["id"]),))
        assert termos_retorno.sincronizar(termo["id"]) == "em_processamento"
        assert _item(client_gestor, item_id)["status"] == "Pendente"
        cur.execute("SELECT RELEASE_LOCK(%s)", (termos_retorno.nome_da_trava(termo["id"]),))
    assert termos_retorno.sincronizar(termo["id"]) == "assinado"


def test_nome_da_trava_cabe_no_limite_do_mysql():
    from app.services import termos_retorno

    assert len(termos_retorno.nome_da_trava(2147483647)) <= 64
    assert termos_retorno.nome_da_trava(7).startswith("patrimonio_termo_7:")


def test_segredo_errado_401(client, revsign_fake):
    resp = postar_webhook(client, evento_webhook("termo.signed", "1001", 1), segredo="errado")
    assert resp.status_code == 401
    assert resp.json()["detail"] == "Webhook não autorizado."


def test_timestamp_velho_ou_do_futuro_401(client, revsign_fake):
    import time

    for desvio in (-310, 310):
        resp = postar_webhook(
            client, evento_webhook("termo.signed", "1001", 1), timestamp=int(time.time()) + desvio
        )
        assert resp.status_code == 401, desvio
        assert resp.json()["detail"] == "Webhook não autorizado."


@pytest.mark.parametrize("digitos", [13, 400, 5000])
def test_timestamp_gigante_401(client, revsign_fake, digitos):
    """Timestamp com dígitos demais é 401 (não estoura em int()/float e vira 500)."""
    import json

    from app.services.revsign_client import assinar_webhook
    from tests.fixtures_termos import URL_WEBHOOK

    corpo = json.dumps(evento_webhook("termo.signed", "1001", 1)).encode()
    cab = assinar_webhook(SEGREDO_WEBHOOK, corpo, 1)
    cab["X-RevSign-Timestamp"] = "9" * digitos
    resp = client.post(URL_WEBHOOK, content=corpo, headers={**cab, "Content-Type": "application/json"})
    assert resp.status_code == 401


def test_assinatura_sem_prefixo_sha256_401(client, revsign_fake):
    import json
    import time

    from app.services.revsign_client import assinar_webhook
    from tests.fixtures_termos import URL_WEBHOOK

    corpo = json.dumps(evento_webhook("termo.signed", "1001", 1)).encode()
    cab = assinar_webhook(SEGREDO_WEBHOOK, corpo, int(time.time()))
    cab["X-RevSign-Signature"] = cab["X-RevSign-Signature"].removeprefix("sha256=")
    resp = client.post(URL_WEBHOOK, content=corpo, headers={**cab, "Content-Type": "application/json"})
    assert resp.status_code == 401


def test_sem_cabecalho_401(client, revsign_fake):
    resp = postar_webhook(client, evento_webhook("termo.signed", "1001", 1), segredo=None)
    assert resp.status_code == 401


def test_segredo_vazio_no_servidor_recusa_tudo(client, monkeypatch):
    from app.core import config
    from app.services.revsign_client import FakeRevSignClient, usar_cliente

    monkeypatch.setattr(config.settings, "REVSIGN_ENABLED", True)
    monkeypatch.setattr(config.settings, "REVSIGN_WEBHOOK_SECRET", "")
    usar_cliente(FakeRevSignClient(webhook_secret=""))
    try:
        resp = postar_webhook(client, evento_webhook("termo.signed", "1001", 1), segredo="")
        assert resp.status_code == 401
        resp = postar_webhook(client, evento_webhook("termo.signed", "1001", 1), segredo=None)
        assert resp.status_code == 401
    finally:
        usar_cliente(None)


def test_integracao_desligada_404(client):
    resp = postar_webhook(client, evento_webhook("termo.signed", "1001", 1))
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Integração com o RevSign desligada."


def test_corpo_invalido_400(client, revsign_fake):
    resp = postar_webhook(client, b"nao-json")
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Corpo do webhook inválido."


def test_corpo_grande_demais_413(client, revsign_fake):
    resp = postar_webhook(client, b"{" + b" " * (1024 * 1024 + 10) + b"}")
    assert resp.status_code == 413
    assert resp.json()["detail"] == "Corpo do webhook muito grande."


def test_documento_desconhecido_ignorado(client, revsign_fake):
    resp = postar_webhook(client, evento_webhook("termo.signed", "5555", 4321))
    assert resp.status_code == 200
    assert resp.json()["resultado"] == "nao_encontrado"


def test_id_de_documento_que_nao_bate_com_a_referencia_e_ignorado(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    termo = _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    revsign_fake.marcar_assinado("1001")
    resp = postar_webhook(client, evento_webhook("termo.signed", "9999", termo["id"]))
    assert resp.json()["resultado"] == "nao_encontrado"
    assert _termo(termo["id"])["status"] == "enviado"


def test_referencia_com_documento_diferente_registra_aviso(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake, caplog
):
    """O externalId é de uma solicitação nossa, mas o documento não é o guardado (ex.: criação
    que deu timeout e ficou manual): o evento é ignorado com aviso no log, sem segredos."""
    import logging

    termo = _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    with caplog.at_level(logging.WARNING, logger="app.services.termos_retorno"):
        resp = postar_webhook(client, evento_webhook("termo.signed", "9999", termo["id"]))
    assert resp.json()["resultado"] == "nao_encontrado"
    avisos = [r.getMessage() for r in caplog.records if r.name == "app.services.termos_retorno"]
    assert len(avisos) == 1
    assert "9999" in avisos[0] and str(termo["id"]) in avisos[0] and "1001" in avisos[0]
    assert SEGREDO_WEBHOOK not in avisos[0]


def test_referencia_sem_documento_guardado_registra_aviso(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake, caplog, monkeypatch
):
    import logging

    from app.services.revsign_client import RevSignError

    def criar(doc):
        raise RevSignError("RevSign indisponível.")  # timeout nas duas tentativas

    monkeypatch.setattr(revsign_fake, "criar_documento", criar)
    termo = _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    assert _termo(termo["id"])["status"] == "manual"
    with caplog.at_level(logging.WARNING, logger="app.services.termos_retorno"):
        resp = postar_webhook(client, evento_webhook("termo.sent", "envelope_orfao", termo["id"]))
    assert resp.json()["resultado"] == "nao_encontrado"
    avisos = [r.getMessage() for r in caplog.records if r.name == "app.services.termos_retorno"]
    assert len(avisos) == 1 and "envelope_orfao" in avisos[0] and "nenhum" in avisos[0]


def test_recusado_marca_recusado(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    revsign_fake.marcar("1001", "recusado", "Dados errados")
    resp = postar_webhook(client, evento_webhook("termo.rejected", "1001", termo["id"]))
    assert resp.json()["resultado"] == "recusado"
    atual = _termo(termo["id"])
    assert atual["status"] == "recusado" and atual["motivo"] == "Dados errados"
    assert _item(client_gestor, item_id)["status"] == "Pendente"
    assert linhas_audit("termo", "atualizar")[0]["username"] == "sistema"


def test_assinado_no_webhook_mas_pendente_no_revsign_nao_confirma(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    """O corpo do webhook diz COMPLETED, mas a fonte da verdade (consulta) diz pendente."""
    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    resp = postar_webhook(client, evento_webhook("termo.signed", "1001", termo["id"]))
    assert resp.json()["resultado"] == "pendente"
    assert _item(client_gestor, item_id)["status"] == "Pendente"


def test_assinatura_de_solicitacao_cancelada_e_ignorada(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    assert client_gestor.post(f"/api/termos/{termo['id']}/cancelar").status_code == 200
    revsign_fake.marcar_assinado("1001")

    resp = postar_webhook(client, evento_webhook("termo.signed", "1001", termo["id"]))
    assert resp.json()["resultado"] == "ja_processado"
    assert _item(client_gestor, item_id)["status"] == "Pendente"
    assert _historico(item_id, "Confirmação Empréstimo") == []


def test_confirmacao_recusada_vira_erro_com_pdf_guardado(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake, executar_sql
):
    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    executar_sql("UPDATE items SET status='Disponível' WHERE id=%s", (item_id,))
    revsign_fake.marcar_assinado("1001")

    resp = postar_webhook(client, evento_webhook("termo.signed", "1001", termo["id"]))
    assert resp.json()["resultado"] == "erro"
    atual = _termo(termo["id"])
    assert atual["status"] == "erro"
    assert atual["motivo"] == "Apenas itens com status 'Pendente' podem ser confirmados."
    assert atual["arquivo_assinado_key"].startswith("termos_assinados/")
    historico = client_gestor.get(f"/api/items/{item_id}/termos").json()["historico"]["items"]
    assert historico[0]["tem_assinado"] is True


def test_devolucao_assinada_pelo_revsign_conclui_devolucao(
    client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake,
    forcar_devolucao_iniciada, termo_devolucao_disponivel,
):
    if not termo_devolucao_disponivel:
        pytest.skip("Modelo do termo de devolução ausente neste ambiente.")
    item_id = item_disponivel["id"]
    _emprestar(client_gestor, item_id, colaborador["id"])
    assert client_gestor.post(
        f"/api/loans/{item_id}/confirm", files={"signed_pdf": ("t.pdf", b"%PDF-x", "application/pdf")}
    ).status_code == 200
    resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate")
    assert resp.status_code == 200, resp.text
    termo = resp.json()["termo"]
    assert termo["tipo"] == "devolucao" and termo["status"] == "enviado"
    documento_id = str(len(revsign_fake.criados) + 1000)
    revsign_fake.marcar_assinado(documento_id)

    resp = postar_webhook(client, evento_webhook("termo.signed", documento_id, termo["id"]))
    assert resp.json()["resultado"] == "assinado"
    item = _item(client_gestor, item_id)
    assert item["status"] == "Disponível" and item["colaborador_id"] is None
    chave = _historico(item_id, "Confirmação Devolução")[0]["termo_assinado_anexo"]
    assert chave.startswith("termos_devolucao_assinados/")


def test_falha_na_consulta_guarda_o_motivo_e_continua_enviado(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    from app.services import termos_retorno

    termo = _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    revsign_fake.falhar_em = {"consultar"}
    assert termos_retorno.sincronizar(termo["id"]) == "falha_consulta"
    atual = _termo(termo["id"])
    assert atual["status"] == "enviado"
    assert atual["motivo"].startswith("Falha ao consultar o RevSign")
    revsign_fake.falhar_em = set()
    assert termos_retorno.sincronizar(termo["id"]) == "pendente"
    assert _termo(termo["id"])["motivo"] is None


def test_anexo_manual_respeita_a_trava_da_solicitacao(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    from app.db.database_mysql import get_cursor
    from app.services import termos, termos_retorno
    from app.services.audit import AuditContext
    from app.services.termos_confirmacao import anexar_assinado

    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    ctx = AuditContext(1, "gestor.teste", "Gestor", "127.0.0.1")

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT GET_LOCK(%s, 0)", (termos_retorno.nome_da_trava(termo["id"]),))
        with pytest.raises(termos.TermoErro, match="está sendo processada"):
            anexar_assinado(termo["id"], b"%PDF-1.4 x", "a.pdf", "tecnico.teste", ctx)
        assert _item(client_gestor, item_id)["status"] == "Pendente"
        cur.execute("SELECT RELEASE_LOCK(%s)", (termos_retorno.nome_da_trava(termo["id"]),))
    anexar_assinado(termo["id"], b"%PDF-1.4 x", "a.pdf", "tecnico.teste", ctx)
    assert _item(client_gestor, item_id)["status"] == "Indisponível"
