"""F7 / Task 1 (Ruling 23): empréstimo que muda a unidade do equipamento registra a transferência."""
from datetime import date, datetime

import pytest

pytestmark = pytest.mark.integration

MOTIVO = "Empréstimo para colaborador de outra unidade"


def _consultar(sql, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


def _emprestar(inv_manager, item_id, revenda):
    hoje = datetime.now().strftime("%d/%m/%Y")
    return inv_manager.issue(
        item_id, "Fulano de Tal", "52998224725", "101 - Puxada", "Analista", "TI", revenda, hoje, "gestor.teste"
    )


@pytest.fixture
def item_com_local(client_gestor, criar_unidade, criar_item, executar_sql):
    unidade = criar_unidade(nome="Revalle Juazeiro")
    resp = client_gestor.post("/api/locais", json={"unidade_id": unidade["id"], "nome": "Sala TI"})
    assert resp.status_code in (200, 201), resp.text
    local_id = resp.json()["id"]
    item = criar_item()
    executar_sql("UPDATE items SET local_id=%s WHERE id=%s", (local_id, item["id"]))
    return item, local_id


def test_emprestimo_para_outra_unidade_grava_transferencia(inv_manager, item_com_local):
    item, local_id = item_com_local
    ok, msg = _emprestar(inv_manager, item["id"], "Revalle Petrolina")
    assert ok is True, msg

    linhas = _consultar("SELECT * FROM transferencias WHERE item_id=%s", (item["id"],))
    assert len(linhas) == 1
    t = linhas[0]
    assert t["de_unidade"] == "Revalle Juazeiro"
    assert t["para_unidade"] == "Revalle Petrolina"
    assert t["de_local_id"] == local_id
    assert t["para_local_id"] is None
    assert t["motivo"] == MOTIVO
    assert t["transferido_por"] == "gestor.teste"
    assert t["transferido_em"].date() == date.today()

    atual = inv_manager.find(item["id"])
    assert atual["revenda"] == "Revalle Petrolina"
    assert atual["local_id"] is None


def test_emprestimo_na_mesma_unidade_nao_grava(inv_manager, item_com_local):
    item, local_id = item_com_local
    ok, msg = _emprestar(inv_manager, item["id"], "REVALLE JUAZEIRO ")
    assert ok is True, msg
    assert _consultar("SELECT * FROM transferencias WHERE item_id=%s", (item["id"],)) == []
    assert inv_manager.find(item["id"])["local_id"] == local_id


def test_falha_na_transferencia_desfaz_o_emprestimo(inv_manager, item_com_local, executar_sql):
    item, local_id = item_com_local
    executar_sql("RENAME TABLE transferencias TO transferencias_f7_tmp")
    try:
        resultado = _emprestar(inv_manager, item["id"], "Revalle Petrolina")
    finally:
        executar_sql("RENAME TABLE transferencias_f7_tmp TO transferencias")
    assert resultado == (False, "Erro ao processar a operação. Tente novamente.")
    atual = inv_manager.find(item["id"])
    assert atual["status"] == "Disponível"
    assert atual["revenda"] == "Revalle Juazeiro"
    assert atual["local_id"] == local_id
    assert _consultar("SELECT id FROM history WHERE item_id=%s AND operation='Empréstimo'", (item["id"],)) == []


def test_linha_do_tempo_mostra_a_transferencia_do_emprestimo(client_gestor, inv_manager, item_com_local):
    item, _ = item_com_local
    assert _emprestar(inv_manager, item["id"], "Revalle Petrolina")[0] is True
    resp = client_gestor.get(f"/api/items/{item['id']}/linha-do-tempo")
    assert resp.status_code == 200, resp.text
    eventos = [e for e in resp.json()["eventos"] if e["tipo"] == "transferencia"]
    assert len(eventos) == 1
    assert eventos[0]["titulo"] == "Transferência Revalle Juazeiro → Revalle Petrolina"
    assert eventos[0]["descricao"].startswith(MOTIVO)


def test_estorno_do_emprestimo_mantem_a_transferencia(inv_manager, item_com_local):
    item, _ = item_com_local
    assert _emprestar(inv_manager, item["id"], "Revalle Petrolina")[0] is True
    emprestimo = _consultar(
        "SELECT id FROM history WHERE item_id=%s AND operation='Empréstimo'", (item["id"],)
    )[0]["id"]
    ok, msg = inv_manager.reverse_history_entry(emprestimo, "gestor.teste")
    assert ok, msg
    assert len(_consultar("SELECT id FROM transferencias WHERE item_id=%s", (item["id"],))) == 1
    assert inv_manager.find(item["id"])["revenda"] == "Revalle Petrolina"
