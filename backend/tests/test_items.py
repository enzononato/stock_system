"""Caracterização do CRUD de itens (backend/app/routers/items.py + InventoryDBManager)."""
from datetime import datetime

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
        """
        MUDANÇA INTENCIONAL (Módulo 6/refatoração): a validação do formato dd/mm/aaaa de
        `date_registered` saiu do router (que devolvia 400 com mensagem própria) e foi
        para `ItemCreate` (Pydantic, via `validar_data_br`). O FastAPI agora responde 422,
        e o handler global `erro_de_validacao` (app/main.py) reduz a lista de erros do
        Pydantic a uma única frase em português citando o rótulo amigável do campo — daí
        "Data de cadastro: Formato de data de cadastro inválido. Use dd/mm/aaaa." em vez
        do antigo "Formato de data inválido. Use dd/mm/yyyy.".
        """
        resp = client_gestor.post(
            "/api/items",
            json={
                "tipo": "Notebook", "brand": "Dell", "model": "Latitude",
                "revenda": "Revalle Juazeiro", "date_registered": "2026-01-01",
            },
        )
        assert resp.status_code == 422
        body = resp.json()
        assert body["detail"] == "Data de cadastro: Formato de data de cadastro inválido. Use dd/mm/aaaa."
        assert isinstance(body["detail"], str)
        assert isinstance(body["errors"], list) and body["errors"]

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
    """
    MUDANÇA INTENCIONAL (T10, paginação server-side): GET /api/items deixou de
    devolver um array JSON cru e passou a devolver o envelope
    `{"items": [...], "total": N}` (schema `Paginated`), aceitando também `limit`
    e `offset`. Os testes abaixo foram atualizados de `resp.json()` (lista) para
    `resp.json()["items"]`; a paginação em si ganhou uma classe de testes dedicada
    logo abaixo (TestPaginacaoDeItems).
    """

    def test_listar_items_retorna_peripheral_count_zero_por_padrao(self, client_gestor, item_disponivel):
        resp = client_gestor.get("/api/items")
        assert resp.status_code == 200
        body = resp.json()
        item = next(i for i in body["items"] if i["id"] == item_disponivel["id"])
        assert item["peripheral_count"] == 0

    def test_listar_items_peripheral_count_reflete_vinculos(
        self, client_gestor, item_disponivel, periferico_disponivel, inv_manager
    ):
        ok, msg = inv_manager.link_peripheral_to_equipment(
            item_disponivel["id"], periferico_disponivel["id"], "teste"
        )
        assert ok, msg
        resp = client_gestor.get("/api/items")
        item = next(i for i in resp.json()["items"] if i["id"] == item_disponivel["id"])
        assert item["peripheral_count"] == 1

    def test_filtro_search_e_case_insensitive_e_cobre_varios_campos(self, client_gestor, criar_item):
        criar_item(brand="Dell", model="Latitude", identificador="ABC123", host="HOST-01")
        criar_item(brand="Lenovo", model="ThinkPad", identificador="XYZ999", host="HOST-02")

        resp = client_gestor.get("/api/items", params={"search": "lenovo"})
        assert resp.status_code == 200
        body = resp.json()
        marcas = {i["brand"] for i in body["items"]}
        assert marcas == {"Lenovo"}
        assert body["total"] == 1

    def test_filtro_status_e_revenda(self, client_gestor, item_disponivel, inv_manager, criar_item):
        """
        CORREÇÃO DE TESTE (não é bug de produto): este teste pedia os fixtures
        `item_disponivel` E `item_pendente` esperando DOIS itens distintos, mas
        `item_pendente` (conftest.py) depende de `item_disponivel` e MUTA esse mesmo
        item para status 'Pendente' via inv_manager.issue() — não cria um item
        separado. Com os dois fixtures juntos, item_disponivel["id"] ==
        item_pendente["id"] (mesma linha, só que o dict `item_disponivel` já foi
        capturado antes da mutação, então ele mesmo ainda mostra "Disponível" em
        memória, mesmo com o banco já tendo mudado). Isso ficava mascarado pela
        falha antiga do envelope de paginação (Grupo A) e só apareceu depois da
        correção. Aqui criamos dois itens de fato independentes: um que permanece
        Disponível e outro que vira Pendente via issue() direto.
        """
        item_pendente_de_verdade = criar_item()
        ok, msg = inv_manager.issue(
            item_pendente_de_verdade["id"], "Fulano de Tal", "11122233344",
            "101 - Puxada", "Analista", "TI", "Revalle Juazeiro",
            datetime.now().strftime("%d/%m/%Y"), "fixture_teste",
        )
        assert ok, msg

        resp = client_gestor.get("/api/items", params={"status": "Pendente"})
        body = resp.json()
        ids = {i["id"] for i in body["items"]}
        assert item_pendente_de_verdade["id"] in ids
        assert item_disponivel["id"] not in ids
        assert body["total"] == 1


class TestPaginacaoDeItems:
    """T10: cobertura da paginação em si (limit/offset/total), nova neste ciclo."""

    def test_total_reflete_quantidade_real_independente_do_limit(self, client_gestor, criar_item):
        for _ in range(3):
            criar_item()
        resp = client_gestor.get("/api/items", params={"limit": 2})
        assert resp.status_code == 200
        body = resp.json()
        assert len(body["items"]) == 2
        assert body["total"] == 3

    def test_offset_pagina_pelos_itens_seguintes_sem_repetir(self, client_gestor, criar_item):
        ids_criados = [criar_item()["id"] for _ in range(3)]
        pagina1 = client_gestor.get("/api/items", params={"limit": 2, "offset": 0}).json()
        pagina2 = client_gestor.get("/api/items", params={"limit": 2, "offset": 2}).json()

        ids_pagina1 = {i["id"] for i in pagina1["items"]}
        ids_pagina2 = {i["id"] for i in pagina2["items"]}
        assert len(pagina1["items"]) == 2
        assert len(pagina2["items"]) == 1
        assert ids_pagina1.isdisjoint(ids_pagina2)
        assert ids_pagina1 | ids_pagina2 == set(ids_criados)

    def test_offset_alem_do_fim_retorna_lista_vazia_mas_total_correto(self, client_gestor, item_disponivel):
        resp = client_gestor.get("/api/items", params={"limit": 10, "offset": 999})
        assert resp.status_code == 200
        body = resp.json()
        assert body["items"] == []
        assert body["total"] == 1

    def test_limit_acima_do_maximo_e_rejeitado_com_422(self, client_gestor):
        from app.core.config import settings

        resp = client_gestor.get("/api/items", params={"limit": settings.MAX_PAGE_SIZE + 1})
        assert resp.status_code == 422

    def test_limit_default_e_o_configurado_em_settings(self, client_gestor):
        from app.core.config import settings

        resp = client_gestor.get("/api/items")
        assert resp.status_code == 200
        # Não há como observar limit/offset na resposta diretamente; a garantia aqui é
        # que a rota não rejeita a ausência dos parâmetros (usa o default de settings).
        assert settings.DEFAULT_PAGE_SIZE > 0

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
