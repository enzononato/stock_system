"""Ganchos do InventoryDBManager usados pelos termos: vínculo do colaborador e geração do
termo sem mudar o status do item."""
from datetime import datetime

import pytest

from tests.fixtures_termos import colaborador, criar_colaborador, unidade_juazeiro  # noqa: F401

pytestmark = pytest.mark.integration


def _emprestar(inv, item_id, colaborador_id=None):
    return inv.issue(
        item_id, "Maria Souza", "52998224725", "101 - Puxada", "Analista", "TI",
        "Revalle Juazeiro", datetime.now().strftime("%d/%m/%Y"), "fixture_teste",
        colaborador_id=colaborador_id,
    )


def test_issue_grava_colaborador_id(inv_manager, item_disponivel, colaborador):
    ok, msg = _emprestar(inv_manager, item_disponivel["id"], colaborador["id"])
    assert ok, msg
    assert inv_manager.find(item_disponivel["id"])["colaborador_id"] == colaborador["id"]


def test_issue_sem_colaborador_grava_null(inv_manager, item_disponivel):
    ok, msg = _emprestar(inv_manager, item_disponivel["id"])
    assert ok, msg
    assert inv_manager.find(item_disponivel["id"])["colaborador_id"] is None


def test_confirm_return_limpa_colaborador_id(
    inv_manager, item_disponivel, colaborador, forcar_devolucao_iniciada
):
    item_id = item_disponivel["id"]
    assert _emprestar(inv_manager, item_id, colaborador["id"])[0]
    assert inv_manager.confirm_loan(item_id, "fixture_teste", "termos_assinados/x.pdf")[0]
    forcar_devolucao_iniciada(item_id)
    assert inv_manager.find(item_id)["colaborador_id"] == colaborador["id"]

    ok, msg = inv_manager.confirm_return(item_id, "fixture_teste", "termos_devolucao_assinados/x.pdf")
    assert ok, msg
    assert inv_manager.find(item_id)["colaborador_id"] is None


def test_estorno_do_emprestimo_limpa_colaborador_id(inv_manager, item_disponivel, colaborador):
    item_id = item_disponivel["id"]
    assert _emprestar(inv_manager, item_id, colaborador["id"])[0]
    history_id = inv_manager.ultimo_history_id(item_id, "Empréstimo")

    ok, msg = inv_manager.reverse_history_entry(history_id, "teste")
    assert ok, msg
    item = inv_manager.find(item_id)
    assert item["status"] == "Disponível"
    assert item["colaborador_id"] is None


def test_ultimo_history_id_ignora_estornados(inv_manager, item_disponivel):
    item_id = item_disponivel["id"]
    assert inv_manager.ultimo_history_id(item_id, "Empréstimo") is None
    assert _emprestar(inv_manager, item_id)[0]
    primeiro = inv_manager.ultimo_history_id(item_id, "Empréstimo")
    assert primeiro is not None
    assert inv_manager.reverse_history_entry(primeiro, "teste")[0]
    assert inv_manager.ultimo_history_id(item_id, "Empréstimo") is None

    assert _emprestar(inv_manager, item_id)[0]
    segundo = inv_manager.ultimo_history_id(item_id, "Empréstimo")
    assert segundo is not None and segundo != primeiro


def test_vincular_colaborador(inv_manager, item_indisponivel, colaborador):
    assert inv_manager.find(item_indisponivel["id"])["colaborador_id"] is None
    inv_manager.vincular_colaborador(item_indisponivel["id"], colaborador["id"])
    assert inv_manager.find(item_indisponivel["id"])["colaborador_id"] == colaborador["id"]


def test_render_loan_term_indisponivel_so_quando_permitido(
    inv_manager, item_indisponivel, unidade_juazeiro, termo_emprestimo_disponivel
):
    if not termo_emprestimo_disponivel:
        pytest.skip("Modelo do termo de responsabilidade ausente neste ambiente.")
    item_id = item_indisponivel["id"]
    assert inv_manager.render_loan_term(item_id) == (
        False, "Este equipamento não está pendente de empréstimo.", None,
    )
    ok, conteudo, nome = inv_manager.render_loan_term(item_id, status_permitidos=("Indisponível",))
    assert ok, conteudo
    assert conteudo[:2] == b"PK"
    assert nome.endswith(".docx")
    # Com outra lista de status a mensagem é a genérica.
    assert inv_manager.render_loan_term(item_id, status_permitidos=("Pendente Devolução",))[1] == (
        "Este equipamento não está com um colaborador."
    )


def test_generate_loan_term_bytes_continua_igual(inv_manager, item_pendente, unidade_juazeiro, termo_emprestimo_disponivel):
    if not termo_emprestimo_disponivel:
        pytest.skip("Modelo do termo de responsabilidade ausente neste ambiente.")
    ok, conteudo, nome = inv_manager.generate_loan_term_bytes(item_pendente["id"])
    assert ok and conteudo[:2] == b"PK" and nome.endswith(".docx")


def test_render_return_term_nao_muda_status(inv_manager, item_indisponivel, termo_devolucao_disponivel):
    if not termo_devolucao_disponivel:
        pytest.skip("Modelo do termo de devolução ausente neste ambiente.")
    item_id = item_indisponivel["id"]
    ok, conteudo, nome = inv_manager.render_return_term(item_id)
    assert ok, conteudo
    assert conteudo[:2] == b"PK"
    assert inv_manager.find(item_id)["status"] == "Indisponível"
    assert inv_manager.ultimo_history_id(item_id, "Devolução") is None

    ok, msg, _ = inv_manager.render_return_term(item_id, status_permitidos=("Pendente Devolução",))
    assert not ok
    assert msg == "Este equipamento não está com devolução pendente."
