"""F7 / Task 3: GET /api/painel/resumo (indicadores calculados no servidor)."""
from datetime import datetime

import pytest

from tests.fixtures_linhas import dados_linha, nova_linha  # noqa: F401

pytestmark = pytest.mark.integration

STATUS = ["Disponível", "Indisponível", "Pendente", "Pendente Devolução", "Em Manutenção"]


def _resumo(client):
    resp = client.get("/api/painel/resumo")
    assert resp.status_code == 200, resp.text
    return resp.json()


def _consultar(sql, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


def test_resumo_vazio(client_gestor):
    r = _resumo(client_gestor)
    assert r["equipamentos"]["total"] == 0
    assert r["equipamentos"]["por_status"] == {
        "Disponível": 0, "Indisponível": 0, "Pendente": 0, "Pendente Devolução": 0, "Em Manutenção": 0,
    }
    assert r["equipamentos"]["por_unidade"] == []
    assert r["termos"] == {
        "pendentes": 0,
        "por_status": {"agendado": 0, "enviado": 0, "manual": 0, "erro": 0},
        "sem_solicitacao": 0,
    }
    assert r["manutencao"] == {"equipamentos_em_manutencao": 0, "ordens_em_andamento": 0}
    assert r["linhas"]["ativas"] == 0
    assert r["linhas"]["custo_mensal"] == "0.00"
    assert r["linhas"]["por_unidade"] == []
    assert r["gerado_em"]


def test_resumo_bate_com_consultas_diretas(client_gestor, criar_item, executar_sql):
    def com_status(status, **kw):
        item = criar_item(**kw)
        executar_sql("UPDATE items SET status=%s WHERE id=%s", (status, item["id"]))
        return item

    criar_item(revenda="Revalle Juazeiro")
    com_status("Indisponível", revenda="Revalle Juazeiro")
    com_status("Em Manutenção", revenda="Revalle Juazeiro")
    com_status("Pendente", revenda="Revalle Petrolina")
    com_status("Pendente Devolução", revenda="Revalle Petrolina")
    sem = criar_item(revenda="Revalle Juazeiro")
    executar_sql("UPDATE items SET revenda=NULL WHERE id=%s", (sem["id"],))
    baixado = criar_item(revenda="Revalle Juazeiro")
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (baixado["id"],))

    r = _resumo(client_gestor)
    eq = r["equipamentos"]
    assert eq["total"] == 6

    diretas = _consultar(
        "SELECT revenda, status, COUNT(*) AS n FROM items WHERE is_active=1 GROUP BY revenda, status"
    )
    por_unidade = {u["unidade"]: u for u in eq["por_unidade"]}
    for linha in diretas:
        assert por_unidade[linha["revenda"]]["por_status"][linha["status"]] == linha["n"]
    por_status = {s: sum(l["n"] for l in diretas if l["status"] == s) for s in STATUS}
    assert eq["por_status"] == por_status
    assert eq["por_status"]["Disponível"] == 2  # um Juazeiro + o sem unidade

    assert eq["por_unidade"][0]["unidade"] == "Revalle Juazeiro" and eq["por_unidade"][0]["total"] == 3
    assert eq["por_unidade"][1]["unidade"] == "Revalle Petrolina" and eq["por_unidade"][1]["total"] == 2
    assert eq["por_unidade"][-1]["unidade"] is None and eq["por_unidade"][-1]["total"] == 1
    assert r["manutencao"]["equipamentos_em_manutencao"] == 1


def test_resumo_agrupa_unidade_vazia_como_sem_unidade(client_gestor, criar_item, executar_sql):
    a = criar_item()
    b = criar_item()
    executar_sql("UPDATE items SET revenda='' WHERE id=%s", (a["id"],))
    executar_sql("UPDATE items SET revenda='   ' WHERE id=%s", (b["id"],))
    por_unidade = _resumo(client_gestor)["equipamentos"]["por_unidade"]
    assert por_unidade == [
        {"unidade": None, "total": 2, "por_status": {**{s: 0 for s in STATUS}, "Disponível": 2}}
    ]


def test_resumo_conta_ordens_e_termos(client_gestor, criar_item, item_pendente):
    from app.db.termo_db import TermoDBManager

    a = criar_item()
    b = criar_item()
    resp = client_gestor.post("/api/manutencoes", json={"item_id": a["id"], "defeito": "Tela"})
    assert resp.status_code == 200, resp.text
    resp = client_gestor.post("/api/manutencoes", json={"item_id": b["id"], "defeito": "Teclado"})
    client_gestor.post(f"/api/manutencoes/{resp.json()['id']}/cancelar", json={"motivo": "Teste"})

    db = TermoDBManager()
    agora = datetime.now()
    for status in ("manual", "enviado", "assinado"):
        item = criar_item()
        db.criar(
            item_id=item["id"], tipo="entrega", status=status, solicitado_por="gestor.teste",
            solicitado_em=agora,
        )

    r = _resumo(client_gestor)
    assert r["manutencao"]["ordens_em_andamento"] == 1
    assert r["termos"]["pendentes"] == 2
    assert r["termos"]["por_status"] == {"agendado": 0, "enviado": 1, "manual": 1, "erro": 0}
    assert r["termos"]["sem_solicitacao"] == 1  # item_pendente legado


def test_resumo_linhas_usa_relatorio_de_custo(client_gestor, nova_linha):
    nova_linha(custo_mensal="49.90")
    nova_linha(custo_mensal="30.10")
    suspensa = nova_linha(custo_mensal="10.00")
    resp = client_gestor.post(f"/api/linhas/{suspensa['id']}/status", json={"status": "suspensa"})
    assert resp.status_code == 200, resp.text

    linhas = _resumo(client_gestor)["linhas"]
    assert linhas["ativas"] == 2
    assert linhas["custo_mensal"] == "80.00"
    assert linhas["por_unidade"][0]["quantidade"] == 2
    assert linhas["por_unidade"][0]["custo_total"] == "80.00"


def test_resumo_liberado_para_todos_os_papeis(client, client_aprendiz, client_tecnico):
    assert client_aprendiz.get("/api/painel/resumo").status_code == 200
    assert client_tecnico.get("/api/painel/resumo").status_code == 200
    assert client.get("/api/painel/resumo").status_code == 401
