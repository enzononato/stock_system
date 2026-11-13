"""Inventário físico: encerramento, relatório e exportação Excel/PDF."""
import json
from io import BytesIO

import pytest

pytestmark = pytest.mark.integration

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _abrir(client, unidade_id):
    resp = client.post("/api/inventarios", json={"unidade_id": unidade_id})
    assert resp.status_code == 201, resp.text
    return resp.json()["id"]


def _ler(client, inv, codigo):
    resp = client.post(f"/api/inventarios/{inv}/leituras", json={"codigo": codigo})
    assert resp.status_code in (200, 201), resp.text
    return resp.json()


def _encerrar(client, inv):
    return client.post(f"/api/inventarios/{inv}/encerrar")


def _nao_encontrados(client, inv):
    linhas = client.get(f"/api/inventarios/{inv}/relatorio").json()["linhas"]
    return {linha["item_id"] for linha in linhas if linha["categoria"] == "nao_encontrado"}


def _cenario(client, unidade, criar_item, executar_sql):
    a, b, c, d = (criar_item(revenda=unidade["nome"]) for _ in range(4))
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (d["id"],))
    inv = _abrir(client, unidade["id"])
    return inv, a["id"], b["id"], c["id"], d["id"]


def test_encerrar_congela_nao_encontrados(client_tecnico, unidade, criar_item, executar_sql):
    inv, a, b, c, _ = _cenario(client_tecnico, unidade, criar_item, executar_sql)
    _ler(client_tecnico, inv, f"#{a}")

    resp = _encerrar(client_tecnico, inv)

    assert resp.status_code == 200, resp.text
    corpo = resp.json()
    assert corpo["status"] == "encerrado" and corpo["encerrado_por"] == "tecnico.teste"
    assert corpo["resumo"]["nao_encontrados"] == 2
    assert corpo["resumo"]["esperados"] == 3
    assert corpo["resumo"]["encontrados"] == 1
    relatorio = client_tecnico.get(f"/api/inventarios/{inv}/relatorio").json()
    antes = relatorio["linhas"]
    assert {linha["item_id"] for linha in antes if linha["categoria"] == "nao_encontrado"} == {b, c}

    executar_sql("UPDATE items SET revenda='Outra' WHERE id=%s", (c,))
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (b,))

    depois = client_tecnico.get(f"/api/inventarios/{inv}/relatorio").json()
    assert {linha["item_id"] for linha in depois["linhas"] if linha["categoria"] == "nao_encontrado"} == {b, c}
    assert depois["inventario"]["resumo"]["nao_encontrados"] == 2


def test_escritas_depois_de_encerrar_409(client_tecnico, unidade, criar_item, executar_sql):
    from app.db.database_mysql import get_cursor

    inv, a, b, _, _ = _cenario(client_tecnico, unidade, criar_item, executar_sql)
    leitura = _ler(client_tecnico, inv, f"#{a}")
    _encerrar(client_tecnico, inv)

    nova = client_tecnico.post(f"/api/inventarios/{inv}/leituras", json={"codigo": f"#{b}"})
    assert nova.status_code == 409
    assert nova.json()["detail"] == "Este inventário não está aberto: não aceita leituras."
    url = f"/api/inventarios/{inv}/leituras/{leitura['id']}"
    assert client_tecnico.patch(url, json={"observacao": "x"}).status_code == 409
    assert client_tecnico.delete(url).status_code == 409
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT COUNT(*) AS n FROM inventario_leituras")
        assert cur.fetchone()["n"] == 1


def test_encerrar_duas_vezes_e_cancelado_409(client_gestor, unidade, criar_unidade):
    inv = _abrir(client_gestor, unidade["id"])
    assert _encerrar(client_gestor, inv).status_code == 200
    again = _encerrar(client_gestor, inv)
    assert again.status_code == 409
    assert again.json()["detail"] == "Só inventários abertos podem ser encerrados."

    outra = criar_unidade()
    cancelado = _abrir(client_gestor, outra["id"])
    client_gestor.post(f"/api/inventarios/{cancelado}/cancelar", json={"motivo": "Teste"})
    assert _encerrar(client_gestor, cancelado).status_code == 409
    assert _encerrar(client_gestor, 999999).status_code == 404


def test_relatorio_aberto_lista_ainda_nao_lidos(client_tecnico, unidade, criar_item, executar_sql):
    inv, a, b, c, _ = _cenario(client_tecnico, unidade, criar_item, executar_sql)
    novo = criar_item(revenda=unidade["nome"])
    _ler(client_tecnico, inv, f"#{a}")

    assert _nao_encontrados(client_tecnico, inv) == {b, c, novo["id"]}
    detalhe = client_tecnico.get(f"/api/inventarios/{inv}").json()
    assert detalhe["resumo"]["esperados"] == 4 and detalhe["resumo"]["nao_encontrados"] == 3


def test_relatorio_ordena_por_categoria(client_tecnico, unidade, criar_unidade, criar_item, executar_sql):
    outra = criar_unidade()
    inv, a, b, c, _ = _cenario(client_tecnico, unidade, criar_item, executar_sql)
    de_fora = criar_item(revenda=outra["nome"])
    _ler(client_tecnico, inv, f"#{a}")
    _ler(client_tecnico, inv, f"#{de_fora['id']}")
    _ler(client_tecnico, inv, "DESCONHECIDO")

    linhas = client_tecnico.get(f"/api/inventarios/{inv}/relatorio").json()["linhas"]

    ordem = ["nao_encontrado", "nao_cadastrado", "outra_unidade", "usuario_diferente", "local_errado", "ok"]
    categorias = [linha["categoria"] for linha in linhas]
    assert categorias == sorted(categorias, key=ordem.index)
    assert categorias == ["nao_encontrado", "nao_encontrado", "nao_cadastrado", "outra_unidade", "ok"]
    assert [linha["item_id"] for linha in linhas[:2]] == [b, c]
    primeiro_ok = linhas[-1]
    assert primeiro_ok["tipo"] == "Notebook" and primeiro_ok["lido_por"] == "tecnico.teste"


def test_relatorio_do_cancelado_nao_tem_nao_encontrados(client_gestor, unidade, criar_item):
    criar_item(revenda=unidade["nome"])
    inv = _abrir(client_gestor, unidade["id"])
    client_gestor.post(f"/api/inventarios/{inv}/cancelar", json={"motivo": "Teste"})

    relatorio = client_gestor.get(f"/api/inventarios/{inv}/relatorio").json()

    assert relatorio["linhas"] == []
    assert relatorio["inventario"]["resumo"]["nao_encontrados"] == 0


def test_exportar_excel(client_tecnico, unidade, criar_item, executar_sql):
    from openpyxl import load_workbook

    inv, a, b, _, _ = _cenario(client_tecnico, unidade, criar_item, executar_sql)
    _ler(client_tecnico, inv, f"#{a}")
    _encerrar(client_tecnico, inv)

    resp = client_tecnico.get(f"/api/inventarios/{inv}/relatorio/exportar?formato=xlsx")

    assert resp.status_code == 200
    assert resp.headers["content-type"] == XLSX
    assert f"inventario_{inv}_" in resp.headers["content-disposition"]
    planilha = load_workbook(BytesIO(resp.content)).active
    celulas = {str(c.value) for linha in planilha.iter_rows() for c in linha if c.value is not None}
    assert "Não encontrado" in celulas
    assert f"#{b}" in celulas
    assert f"Inventário #{inv} — {unidade['nome'].upper()}" in celulas
    assert not any("(parcial)" in c for c in celulas)


def test_exportar_excel_aberto_diz_parcial_e_ainda_nao_lido(client_tecnico, unidade, criar_item, executar_sql):
    from openpyxl import load_workbook

    inv, *_ = _cenario(client_tecnico, unidade, criar_item, executar_sql)

    resp = client_tecnico.get(f"/api/inventarios/{inv}/relatorio/exportar?formato=xlsx")

    planilha = load_workbook(BytesIO(resp.content)).active
    celulas = {str(c.value) for linha in planilha.iter_rows() for c in linha if c.value is not None}
    assert f"Inventário #{inv} — {unidade['nome'].upper()} (parcial)" in celulas
    assert "Ainda não lido" in celulas


def test_exportar_pdf(client_tecnico, unidade, criar_item, executar_sql):
    inv, a, *_ = _cenario(client_tecnico, unidade, criar_item, executar_sql)
    _ler(client_tecnico, inv, f"#{a}")

    resp = client_tecnico.get(f"/api/inventarios/{inv}/relatorio/exportar?formato=pdf")

    assert resp.status_code == 200
    assert resp.content[:4] == b"%PDF"


def test_exportar_formato_invalido_422(client_tecnico, unidade):
    inv = _abrir(client_tecnico, unidade["id"])
    assert client_tecnico.get(f"/api/inventarios/{inv}/relatorio/exportar?formato=csv").status_code == 422
    assert client_tecnico.get(f"/api/inventarios/{inv}/relatorio/exportar").status_code == 422
    assert client_tecnico.get("/api/inventarios/999999/relatorio/exportar?formato=pdf").status_code == 404


def test_historico_do_item(client_tecnico, client_gestor, unidade, criar_item):
    from app.db.inventario_db import InventarioDBManager

    item = criar_item(revenda=unidade["nome"])
    primeiro = _abrir(client_tecnico, unidade["id"])
    _ler(client_tecnico, primeiro, f"#{item['id']}")
    _encerrar(client_tecnico, primeiro)
    segundo = _abrir(client_tecnico, unidade["id"])
    _encerrar(client_tecnico, segundo)
    cancelado = _abrir(client_tecnico, unidade["id"])
    client_gestor.post(f"/api/inventarios/{cancelado}/cancelar", json={"motivo": "Teste"})

    historico = InventarioDBManager().historico_do_item(item["id"])

    assert [e["evento"] for e in historico] == ["leitura", "nao_encontrado"]
    assert [e["inventario_id"] for e in historico] == [primeiro, segundo]
    assert historico[0]["resultado"] == "ok" and historico[0]["por"] == "tecnico.teste"
    assert historico[1]["resultado"] is None and historico[1]["status_inventario"] == "encerrado"
    assert set(historico[0]) == {
        "inventario_id", "unidade_nome", "status_inventario", "evento", "resultado",
        "data", "por", "local_encontrado_nome",
    }


def test_auditoria_encerrar(client_tecnico, unidade, criar_item, executar_sql):
    from app.db.database_mysql import get_cursor

    inv, a, *_ = _cenario(client_tecnico, unidade, criar_item, executar_sql)
    _ler(client_tecnico, inv, f"#{a}")
    _encerrar(client_tecnico, inv)

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT * FROM audit_log WHERE entidade='inventario' AND acao='encerrar'")
        (linha,) = cur.fetchall()
    depois = json.loads(linha["depois"])
    assert linha["entidade_id"] == str(inv) and linha["username"] == "tecnico.teste"
    assert depois["status"] == "encerrado" and depois["nao_encontrados"] == 2
