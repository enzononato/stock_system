"""Caracterização de GET /api/history: envelope de paginação (T10) e busca
server-side (T2, novo neste ciclo) -- ver InventoryDBManager.list_history().

A busca (`?search=`) precisa varrer o BANCO, não a página já carregada: com
paginação server-side, filtrar no cliente só enxergaria as linhas da página atual
e daria a falsa impressão de ter pesquisado o histórico inteiro (daí `total`
também refletir o filtro, não a contagem geral).
"""
from datetime import datetime

import pytest

pytestmark = pytest.mark.integration


class TestPaginacaoDoHistorico:
    def test_envelope_tem_items_e_total(self, client_gestor, item_disponivel):
        resp = client_gestor.get("/api/history")
        assert resp.status_code == 200
        body = resp.json()
        assert "items" in body and "total" in body
        assert body["total"] >= 1  # ao menos o 'Cadastro' do item_disponivel

    def test_limit_pagina_e_total_reflete_contagem_geral(self, client_gestor, criar_item):
        for _ in range(3):
            criar_item()  # 3 entradas 'Cadastro' no histórico
        resp = client_gestor.get("/api/history", params={"limit": 2})
        assert resp.status_code == 200
        body = resp.json()
        assert len(body["items"]) == 2
        assert body["total"] == 3

    def test_offset_alem_do_fim_retorna_vazio_com_total_correto(self, client_gestor, item_disponivel):
        resp = client_gestor.get("/api/history", params={"limit": 10, "offset": 999})
        assert resp.status_code == 200
        body = resp.json()
        assert body["items"] == []
        assert body["total"] == 1

    def test_limit_acima_do_maximo_e_rejeitado_com_422(self, client_gestor):
        from app.core.config import settings

        resp = client_gestor.get("/api/history", params={"limit": settings.MAX_PAGE_SIZE + 1})
        assert resp.status_code == 422

    def test_ordenacao_mais_recente_primeiro(self, client_gestor, criar_item):
        primeiro = criar_item()
        segundo = criar_item()
        resp = client_gestor.get("/api/history", params={"limit": 2})
        ids_item = [row["item_id"] for row in resp.json()["items"]]
        assert ids_item == [segundo["id"], primeiro["id"]]


class TestBuscaServerSideDoHistorico:
    """`search` cobre: operador, usuário, operação, revenda, tipo, marca, modelo e
    identificador (ver campos_busca em InventoryDBManager.list_history)."""

    def test_busca_por_marca_e_case_insensitive(self, client_gestor, criar_item):
        criar_item(brand="Dell", model="Latitude")
        criar_item(brand="Lenovo", model="ThinkPad")

        resp = client_gestor.get("/api/history", params={"search": "lenovo"})
        assert resp.status_code == 200
        body = resp.json()
        assert body["total"] == 1
        assert all(row["marca"] == "Lenovo" for row in body["items"])

    def test_busca_por_modelo(self, client_gestor, criar_item):
        criar_item(brand="Dell", model="Latitude 5420")
        criar_item(brand="Dell", model="Vostro 3510")

        resp = client_gestor.get("/api/history", params={"search": "vostro"})
        body = resp.json()
        assert body["total"] == 1
        assert body["items"][0]["modelo"] == "Vostro 3510"

    def test_busca_por_identificador(self, client_gestor, criar_item):
        criar_item(identificador="SN-BUSCA-XYZ-001")
        criar_item(identificador="SN-OUTRO-002")

        resp = client_gestor.get("/api/history", params={"search": "busca-xyz"})
        body = resp.json()
        assert body["total"] == 1
        assert body["items"][0]["identificador"] == "SN-BUSCA-XYZ-001"

    def test_busca_por_operador(self, client_gestor, criar_item):
        """O 'Cadastro' feito via HTTP registra o operador como o username do
        usuário logado (aqui, 'gestor.teste'), diferente da fábrica `criar_item`
        (que usa `inv.add_item(..., "fixture_teste")` diretamente)."""
        criar_item()  # operador = "fixture_teste"
        resp_post = client_gestor.post(
            "/api/items",
            json={
                "tipo": "Notebook", "brand": "HP", "model": "EliteBook",
                "revenda": "Revalle Juazeiro", "date_registered": "01/01/2026",
            },
        )
        assert resp_post.status_code == 200

        resp = client_gestor.get("/api/history", params={"search": "gestor.teste"})
        body = resp.json()
        assert body["total"] == 1
        assert body["items"][0]["operador"] == "gestor.teste"

    def test_busca_por_usuario_do_emprestimo(self, client_gestor, inv_manager, item_disponivel, criar_item):
        outro_item = criar_item()
        ok, msg = inv_manager.issue(
            item_disponivel["id"], "Ciclano Testável", "11144477735",
            "101 - Puxada", "Analista", "TI", "Revalle Juazeiro",
            datetime.now().strftime("%d/%m/%Y"), "teste",
        )
        assert ok, msg

        resp = client_gestor.get("/api/history", params={"search": "ciclano"})
        body = resp.json()
        encontrados = [row for row in body["items"] if row["usuario"] == "Ciclano Testável"]
        assert encontrados
        assert all(row["item_id"] == item_disponivel["id"] for row in encontrados)

    def test_busca_por_revenda_do_emprestimo(self, client_gestor, inv_manager, item_disponivel):
        """O 'Cadastro' não grava revenda no histórico (só o 'Empréstimo' grava, ver
        InventoryDBManager.issue); por isso o cenário usa um empréstimo."""
        ok, msg = inv_manager.issue(
            item_disponivel["id"], "Fulano", "11144477735",
            "101 - Puxada", "Analista", "TI", "Revalle Ribeira",
            datetime.now().strftime("%d/%m/%Y"), "teste",
        )
        assert ok, msg

        resp = client_gestor.get("/api/history", params={"search": "ribeira"})
        body = resp.json()
        assert body["total"] == 1
        assert body["items"][0]["operation"] == "Empréstimo"
        assert body["items"][0]["revenda"] == "Revalle Ribeira"

    def test_busca_por_operacao(self, client_gestor, criar_periferico):
        """'Cadastro Periférico' não colide com 'Cadastro' de item porque o termo de
        busca inclui a palavra 'Periférico'."""
        criar_periferico()
        resp = client_gestor.get("/api/history", params={"search": "cadastro periférico"})
        body = resp.json()
        assert body["total"] == 1
        assert body["items"][0]["operation"] == "Cadastro Periférico"

    def test_busca_sem_correspondencia_retorna_lista_vazia_e_total_zero(self, client_gestor, item_disponivel):
        resp = client_gestor.get("/api/history", params={"search": "nada-deveria-bater-com-isto-123456"})
        assert resp.status_code == 200
        body = resp.json()
        assert body["items"] == []
        assert body["total"] == 0

    def test_total_reflete_o_filtro_nao_o_total_geral(self, client_gestor, criar_item):
        criar_item(brand="Dell")
        criar_item(brand="Dell")
        criar_item(brand="Lenovo")

        geral = client_gestor.get("/api/history").json()
        filtrado = client_gestor.get("/api/history", params={"search": "dell"}).json()
        assert geral["total"] == 3
        assert filtrado["total"] == 2
        assert filtrado["total"] < geral["total"]
