"""Serviço de solicitação de termos: solicitar, enviar, cancelar e reenviar."""
import hashlib
import json
from datetime import datetime, timedelta

import pytest

from app.services.audit import AuditContext
from tests.fixtures_termos import (  # noqa: F401
    colaborador,
    criar_colaborador,
    linhas_audit,
    revsign_fake,
    unidade_juazeiro,
)

CTX = AuditContext(1, "gestor.teste", "Gestor", "127.0.0.1")
CPF_COLABORADOR = "52998224725"


def _emprestar(inv, item_id, colaborador_id=None, cpf=CPF_COLABORADOR):
    ok, msg = inv.issue(
        item_id, "Maria Souza", cpf, "101 - Puxada", "Analista", "TI", "Revalle Juazeiro",
        datetime.now().strftime("%d/%m/%Y"), "fixture_teste", colaborador_id=colaborador_id,
    )
    assert ok, msg


@pytest.fixture
def emprestimo(inv_manager, item_disponivel, colaborador, unidade_juazeiro):
    """Item `Pendente` emprestado ao colaborador, com a unidade cadastrada (termo gerável)."""
    _emprestar(inv_manager, item_disponivel["id"], colaborador["id"])
    return item_disponivel["id"]


@pytest.mark.unit
def test_versao_modelo(tmp_path):
    from app.services.termos import versao_modelo

    arquivo = tmp_path / "modelo.docx"
    arquivo.write_bytes(b"abc")
    assert versao_modelo(str(arquivo)) == f"modelo.docx@{hashlib.sha256(b'abc').hexdigest()[:8]}"
    assert versao_modelo(str(tmp_path / "nao-existe.docx")) is None
    assert versao_modelo(None) is None


@pytest.mark.integration
def test_solicitar_desligado_fica_manual_com_motivo(emprestimo, colaborador):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    assert termo["status"] == "manual"
    assert termo["motivo"] == termos.MOTIVO_DESLIGADO
    assert termo["modelo_versao"].startswith("termo.docx@")
    assert termo["documento_key"].startswith("terms/")
    assert termo["colaborador_id"] == colaborador["id"]
    assert termo["solicitado_por"] == "gestor.teste"


@pytest.mark.integration
def test_solicitar_ligado_envia_ao_revsign(emprestimo, revsign_fake):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    assert termo["status"] == "enviado"
    assert termo["revsign_documento_id"] == "1001"
    assert termo["enviado_em"] is not None
    assert termo["motivo"] is None
    enviado = revsign_fake.criados[0]
    assert enviado.signatario_email == "maria.souza@gmail.com"
    assert enviado.nome_arquivo.endswith(".docx")
    assert enviado.arquivo[:2] == b"PK"
    assert enviado.referencia_externa == f"patrimonio-termo-{termo['id']}"
    assert "Termo de entrega" in enviado.titulo


@pytest.mark.integration
def test_sem_email_fica_manual(inv_manager, item_disponivel, criar_colaborador, unidade_juazeiro, revsign_fake):
    from app.services import termos

    sem_email = criar_colaborador(email=None)
    _emprestar(inv_manager, item_disponivel["id"], sem_email["id"])
    termo = termos.solicitar(item_disponivel["id"], "entrega", CTX)
    assert termo["status"] == "manual"
    assert termo["motivo"] == termos.MOTIVO_SEM_EMAIL
    assert revsign_fake.criados == []


@pytest.mark.integration
def test_payload_antigo_acha_colaborador_pelo_cpf(
    inv_manager, item_disponivel, colaborador, unidade_juazeiro, revsign_fake, criar_item
):
    from app.services import termos

    _emprestar(inv_manager, item_disponivel["id"], None, cpf=CPF_COLABORADOR)
    termo = termos.solicitar(item_disponivel["id"], "entrega", CTX)
    assert termo["colaborador_id"] == colaborador["id"]
    assert termo["status"] == "enviado"

    outro = criar_item()
    _emprestar(inv_manager, outro["id"], None, cpf="11144477735")
    termo = termos.solicitar(outro["id"], "entrega", CTX)
    assert termo["colaborador_id"] is None
    assert termo["motivo"] == termos.MOTIVO_SEM_COLABORADOR


@pytest.mark.integration
def test_falha_do_revsign_fica_manual_e_nao_propaga(emprestimo, revsign_fake):
    from app.services import termos

    revsign_fake.falhar_em = {"criar_documento"}
    termo = termos.solicitar(emprestimo, "entrega", CTX)
    assert termo["status"] == "manual"
    assert termo["motivo"].startswith("Falha ao enviar ao RevSign (")
    assert termo["revsign_documento_id"] is None


@pytest.mark.integration
def test_falha_ao_gerar_termo_fica_manual(inv_manager, item_disponivel, colaborador, revsign_fake):
    from app.services import termos

    _emprestar(inv_manager, item_disponivel["id"], colaborador["id"])  # sem unidade cadastrada
    termo = termos.solicitar(item_disponivel["id"], "entrega", CTX)
    assert termo["status"] == "manual"
    assert termo["motivo"].startswith(
        "Não foi possível gerar o termo: A unidade 'Revalle Juazeiro' não está cadastrada"
    )
    assert revsign_fake.criados == []


@pytest.mark.integration
def test_nova_solicitacao_cancela_pendente_anterior_tambem_no_revsign(emprestimo, revsign_fake):
    from app.services import termos

    primeira = termos.solicitar(emprestimo, "entrega", CTX)
    segunda = termos.solicitar(emprestimo, "entrega", CTX)
    primeira = termos.get_termo_db().obter(primeira["id"])
    assert primeira["status"] == "cancelado"
    assert primeira["motivo"] == f"Substituída pela solicitação #{segunda['id']}."
    assert primeira["cancelado_em"] is not None
    assert revsign_fake.cancelados == ["1001"]
    assert segunda["status"] == "enviado"
    assert segunda["revsign_documento_id"] == "1002"


@pytest.mark.integration
def test_cancelar_nao_pendente_levanta(emprestimo):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    termos.cancelar(termo["id"], CTX, "Teste")
    with pytest.raises(termos.TermoErro, match="Esta solicitação não está pendente."):
        termos.cancelar(termo["id"], CTX, "De novo")


@pytest.mark.integration
def test_cancelar_com_falha_no_revsign_anota_no_motivo(emprestimo, revsign_fake):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    revsign_fake.falhar_em = {"cancelar"}
    cancelado = termos.cancelar(termo["id"], CTX, "Pedido duplicado")
    assert cancelado["status"] == "cancelado"
    assert cancelado["motivo"] == "Pedido duplicado (não foi possível cancelar no RevSign)"


@pytest.mark.integration
def test_enviar_agendado_com_item_em_outra_situacao_cancela(
    inv_manager, item_indisponivel, colaborador, revsign_fake, executar_sql
):
    from app.services import termos

    item_id = item_indisponivel["id"]
    inv_manager.vincular_colaborador(item_id, colaborador["id"])
    termo = termos.solicitar(
        item_id, "uso", CTX, agendado_para=datetime.now() + timedelta(hours=1)
    )
    assert termo["status"] == "agendado"

    executar_sql("UPDATE items SET status='Disponível' WHERE id=%s", (item_id,))
    termo = termos.enviar(termo["id"], CTX)
    assert termo["status"] == "cancelado"
    assert termo["motivo"] == termos.MOTIVO_SITUACAO_MUDOU
    assert revsign_fake.criados == []


@pytest.mark.integration
def test_reenviar(emprestimo, revsign_fake, executar_sql, colaborador):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    termos.reenviar(termo["id"], CTX)
    assert revsign_fake.reenviados == ["1001"]

    # manual por falta de e-mail: depois de cadastrar o e-mail, reenviar envia de fato
    executar_sql("UPDATE colaboradores SET email=NULL WHERE id=%s", (colaborador["id"],))
    novo = termos.solicitar(emprestimo, "entrega", CTX)
    assert novo["status"] == "manual" and novo["motivo"] == termos.MOTIVO_SEM_EMAIL
    executar_sql(
        "UPDATE colaboradores SET email='maria.souza@gmail.com' WHERE id=%s", (colaborador["id"],)
    )
    reenviado = termos.reenviar(novo["id"], CTX)
    assert reenviado["status"] == "enviado"
    assert reenviado["motivo"] is None

    # outro status não reenvia
    with pytest.raises(termos.TermoErro, match="Só é possível reenviar"):
        termos.reenviar(termo["id"], CTX)  # a primeira foi cancelada pela segunda


@pytest.mark.integration
def test_reenviar_manual_com_integracao_desligada(emprestimo):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    assert termo["status"] == "manual"
    with pytest.raises(termos.TermoErro, match="Integração com o RevSign desligada: use a assinatura manual."):
        termos.reenviar(termo["id"], CTX)


@pytest.mark.integration
def test_auditoria_termo(emprestimo, revsign_fake):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    for acao in ("criar", "enviar"):
        linhas = linhas_audit("termo", acao)
        assert len(linhas) == 1
        assert linhas[0]["entidade_id"] == str(termo["id"])
        assert linhas[0]["username"] == "gestor.teste"
    todas = linhas_audit("termo")
    assert "maria.souza@gmail.com" not in json.dumps([(l["antes"], l["depois"]) for l in todas])


@pytest.mark.integration
def test_para_resposta_remove_chaves_internas(emprestimo, revsign_fake):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    resposta = termos.para_resposta(termo)
    for chave in ("documento_key", "arquivo_assinado_key", "revsign_documento_id", "destinatario_nome"):
        assert chave not in resposta
    assert resposta["tem_documento"] is True
    assert resposta["tem_assinado"] is False
    assert resposta["legado"] is False


# ── Solicitação em processamento (trava ocupada) ─────────────────────────────


@pytest.mark.integration
def test_solicitar_nao_cria_outra_enquanto_a_pendente_esta_em_processamento(emprestimo):
    from app.db.database_mysql import get_cursor
    from app.services import termos
    from app.services.termos_trava import MSG_PROCESSANDO, nome_da_trava

    primeira = termos.solicitar(emprestimo, "entrega", CTX)
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT GET_LOCK(%s, 0)", (nome_da_trava(primeira["id"]),))
        with pytest.raises(termos.TermoErro) as erro:
            termos.solicitar(emprestimo, "entrega", CTX)
        assert str(erro.value) == MSG_PROCESSANDO
        cur.execute("SELECT RELEASE_LOCK(%s)", (nome_da_trava(primeira["id"]),))
    linhas, total = termos.get_termo_db().listar_por_item(emprestimo, limit=10, offset=0)
    assert total == 1 and linhas[0]["status"] == "manual"


@pytest.mark.integration
def test_solicitar_cancela_a_nova_se_a_anterior_ficou_ocupada_no_intervalo(emprestimo, monkeypatch):
    """Corrida: a pré-checagem não viu a trava, mas ela estava ocupada na hora de cancelar."""
    from app.db.database_mysql import get_cursor
    from app.services import termos
    from app.services.termos_trava import MSG_PROCESSANDO, nome_da_trava

    primeira = termos.solicitar(emprestimo, "entrega", CTX)
    monkeypatch.setattr(termos, "_em_processamento", lambda pendentes: [])
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT GET_LOCK(%s, 0)", (nome_da_trava(primeira["id"]),))
        with pytest.raises(termos.TermoErro, match=MSG_PROCESSANDO):
            termos.solicitar(emprestimo, "entrega", CTX)
        cur.execute("SELECT RELEASE_LOCK(%s)", (nome_da_trava(primeira["id"]),))
    linhas, _ = termos.get_termo_db().listar_por_item(emprestimo, limit=10, offset=0)
    por_status = sorted(t["status"] for t in linhas)
    assert por_status == ["cancelado", "manual"]  # a nova foi cancelada; a anterior segue pendente


@pytest.mark.integration
def test_cancelar_por_estorno_registra_aviso_quando_a_trava_esta_ocupada(emprestimo, inv_manager, caplog):
    import logging

    from app.db.database_mysql import get_cursor
    from app.services import termos
    from app.services.termos_trava import nome_da_trava

    historico = inv_manager.ultimo_history_id(emprestimo, "Empréstimo")
    termo = termos.solicitar(emprestimo, "entrega", CTX, history_id=historico)
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT GET_LOCK(%s, 0)", (nome_da_trava(termo["id"]),))
        with caplog.at_level(logging.WARNING):
            assert termos.cancelar_por_estorno(historico, CTX) == []
        cur.execute("SELECT RELEASE_LOCK(%s)", (nome_da_trava(termo["id"]),))
    assert termos.get_termo_db().obter(termo["id"])["status"] == "manual"
    assert any(str(termo["id"]) in r.getMessage() and "processamento" in r.getMessage() for r in caplog.records)


@pytest.mark.integration
def test_reabrir_por_estorno_so_reabre_a_solicitacao_com_o_mesmo_pdf(emprestimo):
    from app.services import termos
    from app.services.termos_confirmacao import anexar_assinado

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    assinado = anexar_assinado(termo["id"], b"%PDF-1.4 x", "a.pdf", "tecnico.teste", CTX)
    chave = assinado["arquivo_assinado_key"]

    outra = {"id": 1, "operation": "Confirmação Empréstimo", "item_id": emprestimo, "termo_assinado_anexo": "termos_assinados/outro.pdf"}
    assert termos.reabrir_por_estorno(outra, CTX) is None
    assert termos.get_termo_db().obter(termo["id"])["status"] == "assinado"

    sem_chave = {**outra, "termo_assinado_anexo": None}
    assert termos.reabrir_por_estorno(sem_chave, CTX) is None
    assert termos.get_termo_db().obter(termo["id"])["status"] == "assinado"

    igual = {**outra, "termo_assinado_anexo": chave}
    assert termos.reabrir_por_estorno(igual, CTX) == termo["id"]
    assert termos.get_termo_db().obter(termo["id"])["status"] == "manual"
