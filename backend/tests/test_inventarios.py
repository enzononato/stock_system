"""Inventário físico: abrir, listar, consultar e cancelar."""
import json

import pytest

pytestmark = pytest.mark.integration

ZEROS = {
    "lidos": 0, "encontrados": 0, "ok": 0, "local_errado": 0, "usuario_diferente": 0,
    "outra_unidade": 0, "nao_cadastrado": 0,
}


def _abrir(client, unidade_id, **extra):
    return client.post("/api/inventarios", json={"unidade_id": unidade_id, **extra})


def _auditoria(entidade, acao):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT * FROM audit_log WHERE entidade = %s AND acao = %s ORDER BY id",
            (entidade, acao),
        )
        return cur.fetchall()


def test_abrir_inventario_tecnico(client_tecnico, unidade, criar_unidade, criar_item, executar_sql):
    outra = criar_unidade()
    criar_item(revenda=unidade["nome"])
    criar_item(revenda=unidade["nome"])
    baixado = criar_item(revenda=unidade["nome"])
    criar_item(revenda=outra["nome"])
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (baixado["id"],))

    resp = _abrir(client_tecnico, unidade["id"], observacao="  Contagem anual  ")

    assert resp.status_code == 201, resp.text
    corpo = resp.json()
    assert corpo["status"] == "aberto"
    assert corpo["iniciado_por"] == "tecnico.teste"
    assert corpo["unidade_nome"] == unidade["nome"]
    assert corpo["observacao"] == "Contagem anual"
    assert corpo["resumo"] == {**ZEROS, "esperados": 2, "nao_encontrados": 2}


def test_segundo_aberto_na_mesma_unidade_409(client_tecnico, unidade):
    primeiro = _abrir(client_tecnico, unidade["id"]).json()["id"]

    resp = _abrir(client_tecnico, unidade["id"])

    assert resp.status_code == 409
    assert resp.json()["detail"] == f"Já existe um inventário aberto para esta unidade (#{primeiro})."


def test_cancelar_libera_novo_inventario(client_gestor, unidade):
    primeiro = _abrir(client_gestor, unidade["id"]).json()["id"]
    assert client_gestor.post(f"/api/inventarios/{primeiro}/cancelar", json={"motivo": "Teste"}).status_code == 200

    segundo = _abrir(client_gestor, unidade["id"])

    assert segundo.status_code == 201
    assert segundo.json()["id"] != primeiro


def test_unidade_inexistente_404_e_inativa_422(client_tecnico, unidade, executar_sql):
    resp = _abrir(client_tecnico, 999999)
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Unidade não encontrada."

    executar_sql("UPDATE unidades SET is_active=0 WHERE id=%s", (unidade["id"],))
    resp = _abrir(client_tecnico, unidade["id"])
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Unidade inativa não pode ser inventariada."


def test_aprendiz_recebe_403(client_aprendiz, unidade):
    for resp in (client_aprendiz.get("/api/inventarios"), _abrir(client_aprendiz, unidade["id"])):
        assert resp.status_code == 403
        assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."


def test_cancelar_so_gestor(client_tecnico, client_gestor, unidade):
    inv = _abrir(client_tecnico, unidade["id"]).json()["id"]
    url = f"/api/inventarios/{inv}/cancelar"

    resp = client_tecnico.post(url, json={"motivo": "Contagem refeita"})
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito ao Gestor."

    resp = client_gestor.post(url, json={"motivo": " "})
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Motivo: Informe o motivo do cancelamento."

    resp = client_gestor.post(url, json={"motivo": "Contagem refeita"})
    assert resp.status_code == 200
    assert resp.json()["status"] == "cancelado"
    assert resp.json()["motivo_cancelamento"] == "Contagem refeita"
    assert resp.json()["encerrado_por"] == "gestor.teste"

    resp = client_gestor.post(url, json={"motivo": "Contagem refeita"})
    assert resp.status_code == 409
    assert resp.json()["detail"] == "Só inventários abertos podem ser cancelados."


def test_listar_filtra_e_pagina(client_tecnico, client_gestor, criar_unidade):
    u1, u2 = criar_unidade(), criar_unidade()
    a = _abrir(client_tecnico, u1["id"]).json()["id"]
    client_gestor.post(f"/api/inventarios/{a}/cancelar", json={"motivo": "Refazer"})
    b = _abrir(client_tecnico, u1["id"]).json()["id"]
    c = _abrir(client_tecnico, u2["id"]).json()["id"]

    todos = client_tecnico.get("/api/inventarios").json()
    assert todos["total"] == 3
    assert [i["id"] for i in todos["items"]] == [c, b, a]

    da_u1 = client_tecnico.get(f"/api/inventarios?unidade_id={u1['id']}").json()
    assert da_u1["total"] == 2 and [i["id"] for i in da_u1["items"]] == [b, a]

    abertos = client_tecnico.get("/api/inventarios?status=aberto").json()
    assert abertos["total"] == 2

    pagina = client_tecnico.get("/api/inventarios?limit=1&offset=1").json()
    assert pagina["total"] == 3 and [i["id"] for i in pagina["items"]] == [b]

    assert client_tecnico.get("/api/inventarios?limit=101").status_code == 422
    assert client_tecnico.get("/api/inventarios?status=xyz").status_code == 422


def test_inventario_inexistente_404(client_tecnico):
    resp = client_tecnico.get("/api/inventarios/999999")
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Inventário não encontrado."


def test_auditoria_criar_e_cancelar(client_gestor, unidade):
    inv = _abrir(client_gestor, unidade["id"], observacao="Anual").json()["id"]
    client_gestor.post(f"/api/inventarios/{inv}/cancelar", json={"motivo": "Contagem refeita"})

    (criar,) = _auditoria("inventario", "criar")
    (cancelar,) = _auditoria("inventario", "cancelar")
    assert criar["entidade_id"] == str(inv) and criar["username"] == "gestor.teste"
    assert json.loads(criar["depois"]) == {
        "unidade_id": unidade["id"], "unidade_nome": unidade["nome"], "observacao": "Anual",
    }
    assert json.loads(cancelar["antes"]) == {"status": "aberto"}
    assert json.loads(cancelar["depois"]) == {"status": "cancelado", "motivo": "Contagem refeita"}
