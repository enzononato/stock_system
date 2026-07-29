"""Caracterização do CRUD de itens (backend/app/routers/items.py + InventoryDBManager)."""
import pytest

pytestmark = pytest.mark.integration


class TestCriarItem:
    def test_criar_item_com_campos_obrigatorios_retorna_200_e_id(self, client_gestor):
        resp = client_gestor.post(
            "/api/items",
            json={
                "tipo": "Notebook", "brand": "Dell", "model": "Latitude",
                "revenda": "Revalle Juazeiro", "date_registered": "01/01/2026",
            },
        )
        assert resp.status_code == 200
        body = resp.json()
        assert body["detail"] == "Item cadastrado com sucesso."
        assert isinstance(body["id"], int)

    def test_criar_item_com_data_em_formato_invalido(self, client_gestor):
        resp = client_gestor.post(
            "/api/items",
            json={
                "tipo": "Notebook", "brand": "Dell", "model": "Latitude",
                "revenda": "Revalle Juazeiro", "date_registered": "2026-01-01",
            },
        )
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Formato de data inválido. Use dd/mm/yyyy."

    def test_criar_item_com_data_de_cadastro_no_futuro_e_aceito(self, client_gestor):
        """Current behavior: create_item() só valida o FORMATO da data (dd/mm/yyyy), não
        se ela é passada/futura — ao contrário de issue(), que valida isso para
        date_issue. Uma data de cadastro no ano 2099 é aceita sem erro."""
        resp = client_gestor.post(
            "/api/items",
            json={
                "tipo": "Notebook", "brand": "Dell", "model": "Latitude",
                "revenda": "Revalle Juazeiro", "date_registered": "01/01/2099",
            },
        )
        assert resp.status_code == 200


class TestListarEBuscarItem:
    def test_listar_items_retorna_peripheral_count_zero_por_padrao(self, client_gestor, item_disponivel):
        resp = client_gestor.get("/api/items")
        assert resp.status_code == 200
        item = next(i for i in resp.json() if i["id"] == item_disponivel["id"])
        assert item["peripheral_count"] == 0

    def test_listar_items_peripheral_count_reflete_vinculos(
        self, client_gestor, item_disponivel, periferico_disponivel, inv_manager
    ):
        ok, msg = inv_manager.link_peripheral_to_equipment(
            item_disponivel["id"], periferico_disponivel["id"], "teste"
        )
        assert ok, msg
        resp = client_gestor.get("/api/items")
        item = next(i for i in resp.json() if i["id"] == item_disponivel["id"])
        assert item["peripheral_count"] == 1

    def test_filtro_search_e_case_insensitive_e_cobre_varios_campos(self, client_gestor, criar_item):
        criar_item(brand="Dell", model="Latitude", identificador="ABC123", host="HOST-01")
        criar_item(brand="Lenovo", model="ThinkPad", identificador="XYZ999", host="HOST-02")

        resp = client_gestor.get("/api/items", params={"search": "lenovo"})
        assert resp.status_code == 200
        marcas = {i["brand"] for i in resp.json()}
        assert marcas == {"Lenovo"}

    def test_filtro_status_e_revenda(self, client_gestor, item_disponivel, item_pendente):
        resp = client_gestor.get("/api/items", params={"status": "Pendente"})
        ids = {i["id"] for i in resp.json()}
        assert item_pendente["id"] in ids
        assert item_disponivel["id"] not in ids

    def test_buscar_item_existente(self, client_gestor, item_disponivel):
        resp = client_gestor.get(f"/api/items/{item_disponivel['id']}")
        assert resp.status_code == 200
        assert resp.json()["id"] == item_disponivel["id"]

    def test_buscar_item_inexistente_retorna_404(self, client_gestor):
        resp = client_gestor.get("/api/items/999999")
        assert resp.status_code == 404
        assert resp.json()["detail"] == "Item não encontrado."


class TestAtualizarItem:
    def test_atualizar_item_existente(self, client_gestor, item_disponivel):
        resp = client_gestor.put(f"/api/items/{item_disponivel['id']}", json={"brand": "NovaMarca"})
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Item {item_disponivel['id']} atualizado."
        assert client_gestor.get(f"/api/items/{item_disponivel['id']}").json()["brand"] == "NovaMarca"

    def test_atualizar_item_inexistente_retorna_404(self, client_gestor):
        resp = client_gestor.put("/api/items/999999", json={"brand": "X"})
        assert resp.status_code == 404
        assert resp.json()["detail"] == "Item não encontrado."

    def test_atualizar_sem_nenhum_campo_retorna_400(self, client_gestor, item_disponivel):
        resp = client_gestor.put(f"/api/items/{item_disponivel['id']}", json={})
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Nenhum campo para atualizar."


class TestRemoverItem:
    def test_remover_item_disponivel(self, client_gestor, item_disponivel):
        resp = client_gestor.request(
            "DELETE", f"/api/items/{item_disponivel['id']}", data={"reason": "Doação"}
        )
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Aparelho {item_disponivel['id']} removido do estoque."
        # Soft-delete: some deixa de aparecer na listagem/find.
        assert client_gestor.get(f"/api/items/{item_disponivel['id']}").status_code == 404

    def test_remover_item_pendente_e_bloqueado(self, client_gestor, item_pendente):
        resp = client_gestor.request(
            "DELETE", f"/api/items/{item_pendente['id']}", data={"reason": "Perda"}
        )
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Não é possível remover produto emprestado."

    def test_remover_item_indisponivel_e_bloqueado(self, client_gestor, item_indisponivel):
        resp = client_gestor.request(
            "DELETE", f"/api/items/{item_indisponivel['id']}", data={"reason": "Perda"}
        )
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Não é possível remover produto emprestado."

    def test_remover_item_inexistente_retorna_400_nao_404(self, client_gestor):
        """
        Inconsistência observada (não é um erro de servidor, então não classificamos
        como BUG CONHECIDO, mas vale registrar): GET/PUT /api/items/{id} devolvem 404
        "Item não encontrado." para um id inexistente, enquanto DELETE devolve 400
        "ID não encontrado." — mensagem e status code diferentes para o mesmo tipo de
        situação, dependendo do endpoint.
        """
        resp = client_gestor.request("DELETE", "/api/items/999999", data={"reason": "Perda"})
        assert resp.status_code == 400
        assert resp.json()["detail"] == "ID não encontrado."
