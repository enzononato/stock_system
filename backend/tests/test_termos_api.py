"""API de termos: vigente, histórico, solicitação, cancelamento, reenvio e anexo."""
from datetime import datetime, timedelta

import pytest

from tests.fixtures_termos import (  # noqa: F401
    colaborador,
    corpo_emprestimo_colaborador,
    criar_colaborador,
    revsign_fake,
    unidade_juazeiro,
)

pytestmark = pytest.mark.integration

ARQUIVO_PDF = {"signed_pdf": ("termo.pdf", b"%PDF-1.4 conteudo", "application/pdf")}


def _emprestar(client, item_id, colaborador_id):
    resp = client.post("/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador_id))
    assert resp.status_code == 200, resp.text
    return resp.json()["termo"]


def _indisponivel_com_colaborador(client, item_id, colaborador_id):
    termo = _emprestar(client, item_id, colaborador_id)
    resp = client.post(f"/api/loans/{item_id}/confirm", files=ARQUIVO_PDF)
    assert resp.status_code == 200, resp.text
    return termo


def test_termos_do_item_traz_vigente_e_historico(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    item_id = item_disponivel["id"]
    _emprestar(client_gestor, item_id, colaborador["id"])

    resp = client_gestor.get(f"/api/items/{item_id}/termos")
    assert resp.status_code == 200
    dados = resp.json()
    assert dados["revsign_ativo"] is True
    vigente = dados["vigente"]
    assert vigente["tipo"] == "entrega"
    assert vigente["status"] == "enviado"
    assert vigente["modelo_versao"].startswith("termo.docx@")
    assert vigente["solicitado_por"] == "gestor.teste"
    assert vigente["colaborador_nome"] == "Maria Souza"
    assert dados["historico"]["total"] == 1
    for linha in [vigente, *dados["historico"]["items"]]:
        assert "documento_key" not in linha and "revsign_documento_id" not in linha
        assert "arquivo_assinado_key" not in linha


def test_termos_item_inexistente_404(client_gestor):
    resp = client_gestor.get("/api/items/9999/termos")
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Equipamento não encontrado."


def test_aprendiz_recebe_403(client_aprendiz, item_disponivel):
    resp = client_aprendiz.get(f"/api/items/{item_disponivel['id']}/termos")
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
    assert client_aprendiz.post(f"/api/items/{item_disponivel['id']}/termos").status_code == 403


def test_tecnico_pode_listar(client_tecnico, item_disponivel):
    assert client_tecnico.get(f"/api/items/{item_disponivel['id']}/termos").status_code == 200


def test_vigente_legado(client_gestor, item_indisponivel):
    resp = client_gestor.get(f"/api/items/{item_indisponivel['id']}/termos")
    vigente = resp.json()["vigente"]
    assert vigente["legado"] is True
    assert vigente["id"] is None
    assert vigente["status"] == "assinado"
    assert resp.json()["historico"]["total"] == 0


def test_solicitar_uso_em_item_indisponivel(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    item_id = item_disponivel["id"]
    _indisponivel_com_colaborador(client_gestor, item_id, colaborador["id"])

    resp = client_gestor.post(f"/api/items/{item_id}/termos")
    assert resp.status_code == 201, resp.text
    primeira = resp.json()
    assert primeira["tipo"] == "uso" and primeira["status"] == "enviado"

    segunda = client_gestor.post(f"/api/items/{item_id}/termos").json()
    historico = client_gestor.get(f"/api/items/{item_id}/termos").json()["historico"]["items"]
    por_id = {t["id"]: t for t in historico}
    assert por_id[primeira["id"]]["status"] == "cancelado"
    assert por_id[segunda["id"]]["status"] == "enviado"


def test_solicitar_em_item_disponivel_400(client_gestor, item_disponivel):
    resp = client_gestor.post(f"/api/items/{item_disponivel['id']}/termos")
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Só é possível solicitar assinatura de equipamento com colaborador."


def test_solicitar_item_inexistente_404(client_gestor):
    assert client_gestor.post("/api/items/9999/termos").status_code == 404


def test_agendar_com_integracao_desligada_400(client_gestor, item_indisponivel):
    futuro = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%dT%H:%M")
    resp = client_gestor.post(f"/api/items/{item_indisponivel['id']}/termos", json={"agendado_para": futuro})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Agendamento disponível só com a integração do RevSign ligada."


def test_agendar_no_passado_400(client_gestor, item_indisponivel, revsign_fake):
    passado = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%dT%H:%M")
    resp = client_gestor.post(f"/api/items/{item_indisponivel['id']}/termos", json={"agendado_para": passado})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "A data agendada precisa ser no futuro."


def test_agendar_converte_fuso_para_brasilia(client_gestor, item_indisponivel, colaborador, revsign_fake, inv_manager):
    inv_manager.vincular_colaborador(item_indisponivel["id"], colaborador["id"])
    from app.core.tempo import agora_local

    alvo = (agora_local() + timedelta(days=2)).replace(second=0, microsecond=0)
    com_fuso = alvo.isoformat() + "-03:00"  # Brasília é UTC-3 o ano todo
    resp = client_gestor.post(f"/api/items/{item_indisponivel['id']}/termos", json={"agendado_para": com_fuso})
    assert resp.status_code == 201, resp.text
    assert resp.json()["status"] == "agendado"
    assert resp.json()["agendado_para"].startswith(alvo.strftime("%Y-%m-%dT%H:%M"))


def test_cancelar_e_reenviar(client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake):
    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])

    resp = client_gestor.post(f"/api/termos/{termo['id']}/reenviar")
    assert resp.status_code == 200, resp.text
    assert revsign_fake.reenviados == ["1001"]

    resp = client_gestor.post(f"/api/termos/{termo['id']}/cancelar", json={"motivo": "Pedido duplicado"})
    assert resp.status_code == 200
    assert resp.json()["status"] == "cancelado"
    assert resp.json()["motivo"] == "Pedido duplicado"

    resp = client_gestor.post(f"/api/termos/{termo['id']}/cancelar")
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Esta solicitação não está pendente."
    resp = client_gestor.post(f"/api/termos/{termo['id']}/reenviar")
    assert resp.status_code == 400

    assert client_gestor.post("/api/termos/9999/cancelar").status_code == 404
    assert client_gestor.post("/api/termos/9999/reenviar").status_code == 404


def test_cancelar_sem_motivo_usa_o_padrao(client_gestor, item_disponivel, colaborador, unidade_juazeiro):
    termo = _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    resp = client_gestor.post(f"/api/termos/{termo['id']}/cancelar")
    assert resp.json()["motivo"] == "Cancelada por gestor.teste."


def test_anexar_assinado_e_arquivos(client_gestor, item_indisponivel, colaborador, inv_manager):
    item_id = item_indisponivel["id"]
    inv_manager.vincular_colaborador(item_id, colaborador["id"])
    termo = client_gestor.post(f"/api/items/{item_id}/termos").json()
    assert termo["tipo"] == "uso" and termo["status"] == "manual"

    resp = client_gestor.post(f"/api/termos/{termo['id']}/assinado", files=ARQUIVO_PDF)
    assert resp.status_code == 200, resp.text
    assert resp.json()["tem_assinado"] is True
    assert resp.json()["status"] == "assinado"

    arquivos = client_gestor.get(f"/api/termos/{termo['id']}/arquivos").json()
    assert arquivos["assinado_url"].startswith("/api/documents/files/termos_assinados/")
    if termo["tem_documento"]:
        assert arquivos["documento_url"].startswith("/api/documents/files/terms/")
    assert client_gestor.get("/api/termos/9999/arquivos").status_code == 404
    assert client_gestor.post("/api/termos/9999/assinado", files=ARQUIVO_PDF).status_code == 404


def test_documento_do_termo_manual_fica_disponivel(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro
):
    termo = _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    assert termo["tem_documento"] is True
    arquivos = client_gestor.get(f"/api/termos/{termo['id']}/arquivos").json()
    assert arquivos["documento_url"].startswith("/api/documents/files/terms/")
    assert arquivos["assinado_url"] is None


def test_anexar_assinado_acima_do_limite_413(client_gestor, item_disponivel, colaborador, unidade_juazeiro):
    termo = _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    grande = {"signed_pdf": ("termo.pdf", b"a" * (20 * 1024 * 1024 + 1), "application/pdf")}
    resp = client_gestor.post(f"/api/termos/{termo['id']}/assinado", files=grande)
    assert resp.status_code == 413
    assert resp.json()["detail"] == "Arquivo excede o tamanho máximo de 20MB."


def test_anexar_assinado_entrega_confirma_emprestimo(client_gestor, item_disponivel, colaborador, unidade_juazeiro):
    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    resp = client_gestor.post(f"/api/termos/{termo['id']}/assinado", files=ARQUIVO_PDF)
    assert resp.status_code == 200, resp.text
    assert client_gestor.get(f"/api/items/{item_id}").json()["status"] == "Indisponível"


def test_anexar_em_solicitacao_cancelada_400(client_gestor, item_disponivel, colaborador, unidade_juazeiro):
    termo = _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    client_gestor.post(f"/api/termos/{termo['id']}/cancelar")
    resp = client_gestor.post(f"/api/termos/{termo['id']}/assinado", files=ARQUIVO_PDF)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Esta solicitação não aceita anexo."


def test_termos_do_colaborador(client_gestor, item_disponivel, criar_item, colaborador, unidade_juazeiro):
    _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    outro = criar_item()
    _emprestar(client_gestor, outro["id"], colaborador["id"])

    resp = client_gestor.get(f"/api/colaboradores/{colaborador['id']}/termos?limit=1")
    assert resp.status_code == 200
    dados = resp.json()
    assert dados["total"] == 2 and len(dados["items"]) == 1
    assert dados["items"][0]["item_descricao"] == "Notebook Dell Latitude 5420"
    assert client_gestor.get("/api/colaboradores/9999/termos").status_code == 404


def test_solicitar_responde_409_com_a_pendente_em_processamento(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro
):
    from app.db.database_mysql import get_cursor
    from app.services.termos_trava import MSG_PROCESSANDO, nome_da_trava

    item_id = item_disponivel["id"]
    termo = _indisponivel_com_colaborador(client_gestor, item_id, colaborador["id"])
    uso = client_gestor.post(f"/api/items/{item_id}/termos").json()
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT GET_LOCK(%s, 0)", (nome_da_trava(uso["id"]),))
        resp = client_gestor.post(f"/api/items/{item_id}/termos")
        assert resp.status_code == 409
        assert resp.json()["detail"] == MSG_PROCESSANDO
        assert client_gestor.post(f"/api/termos/{uso['id']}/cancelar").status_code == 409
        cur.execute("SELECT RELEASE_LOCK(%s)", (nome_da_trava(uso["id"]),))
    assert termo is not None
