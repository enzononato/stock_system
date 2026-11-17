"""Etiquetas com QR: dados da folha e seleção de todos os itens da unidade."""
import pytest

pytestmark = pytest.mark.integration


def test_url_do_qr_so_tem_endereco_e_id(client_tecnico, unidade, criar_item, executar_sql, monkeypatch):
    from app.core import config

    monkeypatch.setattr(config.settings, "APP_PUBLIC_URL", "https://patrimonio.revalle.com.br/")
    item = criar_item(revenda=unidade["nome"])
    executar_sql(
        "UPDATE items SET assigned_to=%s, cpf=%s WHERE id=%s",
        ("Fulano de Tal", "12345678909", item["id"]),
    )

    resp = client_tecnico.post("/api/etiquetas", json={"item_ids": [item["id"]]})

    assert resp.status_code == 200, resp.text
    corpo = resp.json()
    assert corpo["base_url"] == "https://patrimonio.revalle.com.br"
    etiqueta = corpo["etiquetas"][0]
    assert etiqueta["url"] == f"https://patrimonio.revalle.com.br/e/{item['id']}"
    assert "Fulano" not in resp.text
    assert "12345678909" not in resp.text
    assert set(etiqueta) == {"id", "tipo", "brand", "model", "identificador", "revenda", "url"}
    assert etiqueta["tipo"] == "Notebook" and etiqueta["model"] == "Latitude 5420"
    assert corpo["ignorados"] == []


def test_ordena_e_ignora_baixados_e_inexistentes(client_tecnico, unidade, criar_item, executar_sql):
    a = criar_item(revenda=unidade["nome"])
    c = criar_item(revenda=unidade["nome"])
    baixado = criar_item(revenda=unidade["nome"])
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (baixado["id"],))

    resp = client_tecnico.post(
        "/api/etiquetas", json={"item_ids": [c["id"], a["id"], baixado["id"], 999999, a["id"]]}
    )

    corpo = resp.json()
    assert [e["id"] for e in corpo["etiquetas"]] == sorted([a["id"], c["id"]])
    assert corpo["ignorados"] == sorted([baixado["id"], 999999])


def test_limites_da_selecao(client_tecnico):
    resp = client_tecnico.post("/api/etiquetas", json={"item_ids": []})
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Equipamentos: Selecione ao menos um equipamento."

    resp = client_tecnico.post("/api/etiquetas", json={"item_ids": list(range(1, 2002))})
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Equipamentos: Selecione no máximo 2000 equipamentos por vez."

    assert client_tecnico.post("/api/etiquetas", json={"item_ids": list(range(1, 2001))}).status_code == 200


def test_ids_da_unidade_com_filtros(client_tecnico, unidade, criar_unidade, criar_item):
    outra = criar_unidade()
    n1 = criar_item(revenda=unidade["nome"], tipo="Notebook")
    n2 = criar_item(revenda=unidade["nome"], tipo="Notebook")
    m = criar_item(revenda=unidade["nome"], tipo="Monitor")
    criar_item(revenda=outra["nome"], tipo="Notebook")

    todos = client_tecnico.get("/api/etiquetas/ids", params={"revenda": unidade["nome"]}).json()
    assert todos["total"] == 3
    assert todos["ids"] == sorted([n1["id"], n2["id"], m["id"]])

    notebooks = client_tecnico.get(
        "/api/etiquetas/ids", params={"revenda": unidade["nome"], "tipo": "Notebook"}
    ).json()
    assert notebooks["total"] == 2 and notebooks["ids"] == sorted([n1["id"], n2["id"]])

    resp = client_tecnico.get("/api/etiquetas/ids")
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Escolha a unidade."
    assert client_tecnico.get("/api/etiquetas/ids", params={"revenda": " "}).status_code == 422


def test_aprendiz_403(client_aprendiz):
    for resp in (
        client_aprendiz.post("/api/etiquetas", json={"item_ids": [1]}),
        client_aprendiz.get("/api/etiquetas/ids", params={"revenda": "X"}),
    ):
        assert resp.status_code == 403
        assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
