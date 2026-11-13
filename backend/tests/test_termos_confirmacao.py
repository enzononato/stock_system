"""Confirmação de assinatura compartilhada pelo upload manual e pelo RevSign."""
from datetime import datetime

import pytest

from app.services.audit import AuditContext
from tests.fixtures_termos import (  # noqa: F401
    colaborador,
    criar_colaborador,
    linhas_audit,
    revsign_fake,
    unidade_juazeiro,
)

pytestmark = pytest.mark.integration

CTX = AuditContext(1, "gestor.teste", "Gestor", "127.0.0.1")
ARQUIVO_PDF = {"signed_pdf": ("termo.pdf", b"conteudo-fake-pdf", "application/pdf")}


def _emprestar(inv, item_id, colaborador_id):
    ok, msg = inv.issue(
        item_id, "Maria Souza", "52998224725", "101 - Puxada", "Analista", "TI", "Revalle Juazeiro",
        datetime.now().strftime("%d/%m/%Y"), "fixture_teste", colaborador_id=colaborador_id,
    )
    assert ok, msg


def _confirmacoes(item_id, operacao="Confirmação Empréstimo"):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT * FROM history WHERE item_id=%s AND operation=%s ORDER BY id", (item_id, operacao)
        )
        return list(cur.fetchall())


@pytest.fixture
def emprestimo(inv_manager, item_disponivel, colaborador, unidade_juazeiro):
    _emprestar(inv_manager, item_disponivel["id"], colaborador["id"])
    return item_disponivel["id"]


def test_confirmacao_manual_marca_entrega_assinada(client_gestor, emprestimo):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    resp = client_gestor.post(f"/api/loans/{emprestimo}/confirm", files=ARQUIVO_PDF)
    assert resp.status_code == 200
    assert resp.json()["detail"] == f"Empréstimo do item {emprestimo} confirmado."

    atual = termos.get_termo_db().obter(termo["id"])
    confirmacao = _confirmacoes(emprestimo)[0]
    assert atual["status"] == "assinado"
    assert atual["assinado_em"] is not None
    assert atual["arquivo_assinado_key"] == confirmacao["termo_assinado_anexo"]
    assert len(linhas_audit("emprestimo", "confirmar")) == 1
    assinar = linhas_audit("termo", "assinar")
    assert len(assinar) == 1 and assinar[0]["depois"]["origem"] == "manual"


def test_upload_manual_com_solicitacao_enviada_cancela_no_revsign(client_gestor, emprestimo, revsign_fake):
    from app.services import termos

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    assert termo["status"] == "enviado"
    resp = client_gestor.post(f"/api/loans/{emprestimo}/confirm", files=ARQUIVO_PDF)
    assert resp.status_code == 200
    assert termos.get_termo_db().obter(termo["id"])["status"] == "assinado"
    assert revsign_fake.cancelados == ["1001"]


def test_confirmacao_sem_solicitacao_continua_funcionando(client_gestor, item_pendente):
    from app.db.database_mysql import get_cursor

    resp = client_gestor.post(f"/api/loans/{item_pendente['id']}/confirm", files=ARQUIVO_PDF)
    assert resp.status_code == 200
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT COUNT(*) AS n FROM termo_solicitacoes")
        assert cur.fetchone()["n"] == 0


def test_confirmacao_recusada_devolve_a_chave_e_nao_marca_a_solicitacao(emprestimo, revsign_fake):
    from app.services import termos
    from app.services.termos_confirmacao import confirmar_emprestimo

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    ok, msg, chave = confirmar_emprestimo(emprestimo, b"%PDF-x", "a.pdf", "RevSign", CTX, termo_id=termo["id"], origem="revsign")
    assert ok
    # segunda confirmação: o item já não está Pendente
    ok, msg, chave = confirmar_emprestimo(emprestimo, b"%PDF-x", "b.pdf", "RevSign", CTX, termo_id=termo["id"], origem="revsign")
    assert not ok
    assert msg == "Apenas itens com status 'Pendente' podem ser confirmados."
    assert chave.startswith("termos_assinados/")
    assert revsign_fake.cancelados == []  # origem revsign não cancela nada


def test_confirmar_devolucao_marca_devolucao_assinada(
    client_gestor, inv_manager, emprestimo, forcar_devolucao_iniciada
):
    from app.services import termos
    from app.services.termos import DocumentoGerado

    assert inv_manager.confirm_loan(emprestimo, "fixture_teste", "termos_assinados/x.pdf")[0]
    forcar_devolucao_iniciada(emprestimo)
    termo = termos.solicitar(
        emprestimo, "devolucao", CTX,
        documento=DocumentoGerado(b"PK", "termo_devolucao.docx", "termos_devolucao/termo_devolucao.docx"),
    )
    assert termo["tipo"] == "devolucao" and termo["status"] == "manual"

    resp = client_gestor.post(f"/api/loans/{emprestimo}/return/confirm", files=ARQUIVO_PDF)
    assert resp.status_code == 200
    atual = termos.get_termo_db().obter(termo["id"])
    assert atual["status"] == "assinado"
    assert atual["arquivo_assinado_key"].startswith("termos_devolucao_assinados/")
    assert len(linhas_audit("devolucao", "confirmar")) == 1


def test_anexar_assinado_uso_grava_pdf(inv_manager, item_indisponivel, colaborador):
    from app.services import termos
    from app.services.termos import DocumentoGerado
    from app.services.termos_confirmacao import anexar_assinado

    item_id = item_indisponivel["id"]
    inv_manager.vincular_colaborador(item_id, colaborador["id"])
    termo = termos.solicitar(
        item_id, "uso", CTX,
        documento=DocumentoGerado(b"PK", "termo.docx", "terms/termo.docx"),
    )
    assinado = anexar_assinado(termo["id"], b"%PDF-1.4 x", "assinado.pdf", "tecnico.teste", CTX)
    assert assinado["status"] == "assinado"
    assert assinado["arquivo_assinado_key"].startswith("termos_assinados/termo_uso_")
    assert inv_manager.find(item_id)["status"] == "Indisponível"

    outro = termos.solicitar(item_id, "uso", CTX, documento=DocumentoGerado(b"PK", "t.docx", "terms/t.docx"))
    termos.cancelar(outro["id"], CTX, "Teste")
    with pytest.raises(termos.TermoErro, match="Esta solicitação não aceita anexo."):
        anexar_assinado(outro["id"], b"%PDF-1.4 x", "assinado.pdf", "tecnico.teste", CTX)


def test_anexar_assinado_entrega_confirma_emprestimo(inv_manager, emprestimo):
    from app.services import termos
    from app.services.termos_confirmacao import anexar_assinado

    termo = termos.solicitar(emprestimo, "entrega", CTX)
    anexar_assinado(termo["id"], b"%PDF-1.4 x", "assinado.pdf", "tecnico.teste", CTX)
    assert inv_manager.find(emprestimo)["status"] == "Indisponível"
    assert len(_confirmacoes(emprestimo)) == 1
