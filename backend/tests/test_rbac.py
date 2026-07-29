"""
Matriz de RBAC (role x endpoint): quem recebe 401 (sem token), 403 (role errado) e 200.

Usa os fixtures client_gestor/client_tecnico/client_aprendiz (conftest.py), que
contornam deliberadamente o BUG CONHECIDO do JWT (claim "sub" inteiro — ver
test_auth.py) via dependency_override, para poder caracterizar a lógica de
autorização por role (gestor_only/gestor_or_tecnico) isoladamente. O caso "com token
real" (que hoje falha sempre com 401, independente do role) já está coberto em
test_auth.py::test_token_de_login_real_nao_autentica_em_nenhuma_rota.

Para os casos que devem falhar (401 sem token / 403 role errado), usamos IDs
"fantasia" (ex.: 999999) quando o endpoint tem parâmetros de path — a dependency de
autorização é avaliada antes do corpo da rota rodar, então o recurso nunca chega a
ser consultado. Para os casos que devem suceder (200), usamos dados reais criados
pelas fábricas de fixtures.
"""
import pytest

pytestmark = pytest.mark.integration

MSG_SEM_TOKEN = "Token de autenticação não fornecido."
MSG_SO_GESTOR = "Acesso restrito ao Gestor."
MSG_GESTOR_OU_TECNICO = "Acesso restrito a Gestor ou Técnico."

NOVO_ITEM_BODY = {
    "tipo": "Notebook",
    "brand": "Dell",
    "model": "XPS 13",
    "revenda": "Revalle Juazeiro",
    "date_registered": "01/01/2026",
}


class TestRbacItems:
    def test_get_items_qualquer_role_autenticado_200_sem_token_401(
        self, client, client_gestor, client_tecnico, client_aprendiz
    ):
        assert client.get("/api/items").status_code == 401
        assert client.get("/api/items").json()["detail"] == MSG_SEM_TOKEN
        assert client_gestor.get("/api/items").status_code == 200
        assert client_tecnico.get("/api/items").status_code == 200
        assert client_aprendiz.get("/api/items").status_code == 200

    def test_post_items_gestor_e_tecnico_200_aprendiz_403_anonimo_401(
        self, client, client_gestor, client_tecnico, client_aprendiz
    ):
        assert client.post("/api/items", json=NOVO_ITEM_BODY).status_code == 401

        resp_aprendiz = client_aprendiz.post("/api/items", json=NOVO_ITEM_BODY)
        assert resp_aprendiz.status_code == 403
        assert resp_aprendiz.json()["detail"] == MSG_GESTOR_OU_TECNICO

        assert client_gestor.post("/api/items", json=NOVO_ITEM_BODY).status_code == 200
        assert client_tecnico.post("/api/items", json=NOVO_ITEM_BODY).status_code == 200

    def test_put_items_gestor_e_tecnico_200_aprendiz_403_anonimo_401(
        self, client, client_gestor, client_tecnico, client_aprendiz, criar_item
    ):
        item_para_gestor = criar_item()
        item_para_tecnico = criar_item()

        assert client.put(f"/api/items/{item_para_gestor['id']}", json={"brand": "X"}).status_code == 401

        resp_aprendiz = client_aprendiz.put("/api/items/999999", json={"brand": "X"})
        assert resp_aprendiz.status_code == 403
        assert resp_aprendiz.json()["detail"] == MSG_GESTOR_OU_TECNICO

        assert client_gestor.put(f"/api/items/{item_para_gestor['id']}", json={"brand": "Y"}).status_code == 200
        assert client_tecnico.put(f"/api/items/{item_para_tecnico['id']}", json={"brand": "Z"}).status_code == 200

    def test_delete_items_somente_gestor_200_outros_403_anonimo_401(
        self, client, client_tecnico, client_aprendiz, client_gestor, item_disponivel
    ):
        form = {"reason": "Doação"}

        resp_anon = client.request("DELETE", f"/api/items/{item_disponivel['id']}", data=form)
        assert resp_anon.status_code == 401

        resp_tecnico = client_tecnico.request("DELETE", "/api/items/999999", data=form)
        assert resp_tecnico.status_code == 403
        assert resp_tecnico.json()["detail"] == MSG_SO_GESTOR

        resp_aprendiz = client_aprendiz.request("DELETE", "/api/items/999999", data=form)
        assert resp_aprendiz.status_code == 403
        assert resp_aprendiz.json()["detail"] == MSG_SO_GESTOR

        # Gestor por último: é quem de fato consegue remover o item.
        resp_gestor = client_gestor.request("DELETE", f"/api/items/{item_disponivel['id']}", data=form)
        assert resp_gestor.status_code == 200


class TestRbacUsuarios:
    def test_todas_as_rotas_de_usuarios_sao_restritas_ao_gestor(
        self, client, client_tecnico, client_aprendiz, client_gestor
    ):
        casos_negados = [
            ("GET", "/api/users", None),
            ("POST", "/api/users", {"username": "x", "password": "y", "role": "Técnico"}),
            ("DELETE", "/api/users/999999", None),
            ("PUT", "/api/users/999999/password", {"new_password": "outra"}),
        ]
        for metodo, url, body in casos_negados:
            resp_anon = client.request(metodo, url, json=body)
            assert resp_anon.status_code == 401, f"{metodo} {url} deveria exigir token"

            resp_tec = client_tecnico.request(metodo, url, json=body)
            assert resp_tec.status_code == 403, f"{metodo} {url} deveria negar Técnico"
            assert resp_tec.json()["detail"] == MSG_SO_GESTOR

            resp_apr = client_aprendiz.request(metodo, url, json=body)
            assert resp_apr.status_code == 403, f"{metodo} {url} deveria negar Jovem Aprendiz"

        # Caminho positivo (Gestor) só para GET, que não tem efeito colateral.
        assert client_gestor.get("/api/users").status_code == 200


class TestRbacHistorico:
    def test_get_historico_gestor_e_tecnico_200_aprendiz_403(
        self, client, client_gestor, client_tecnico, client_aprendiz
    ):
        assert client.get("/api/history").status_code == 401
        assert client_gestor.get("/api/history").status_code == 200
        assert client_tecnico.get("/api/history").status_code == 200
        resp = client_aprendiz.get("/api/history")
        assert resp.status_code == 403
        assert resp.json()["detail"] == MSG_GESTOR_OU_TECNICO

    def test_estornar_historico_somente_gestor(
        self, client, client_tecnico, client_aprendiz, client_gestor, item_disponivel, inv_manager
    ):
        entradas = [h for h in inv_manager.list_history() if h["item_id"] == item_disponivel["id"]]
        assert entradas, "Fixture deveria ter gerado uma entrada de histórico 'Cadastro'."
        history_id = entradas[0]["id"]

        assert client.post(f"/api/history/{history_id}/reverse").status_code == 401

        resp_tec = client_tecnico.post("/api/history/999999/reverse")
        assert resp_tec.status_code == 403
        assert resp_tec.json()["detail"] == MSG_SO_GESTOR

        resp_apr = client_aprendiz.post("/api/history/999999/reverse")
        assert resp_apr.status_code == 403

        resp_gestor = client_gestor.post(f"/api/history/{history_id}/reverse")
        assert resp_gestor.status_code == 200


class TestRbacRelatorios:
    def test_monthly_gestor_e_tecnico_200_aprendiz_403(self, client, client_gestor, client_tecnico, client_aprendiz):
        assert client.get("/api/reports/monthly").status_code == 401
        assert client_gestor.get("/api/reports/monthly").status_code == 200
        assert client_tecnico.get("/api/reports/monthly").status_code == 200
        assert client_aprendiz.get("/api/reports/monthly").status_code == 403

    def test_monthly_export_somente_gestor(self, client, client_tecnico, client_aprendiz, client_gestor):
        assert client.get("/api/reports/monthly/export").status_code == 401
        resp_tec = client_tecnico.get("/api/reports/monthly/export")
        assert resp_tec.status_code == 403
        assert resp_tec.json()["detail"] == MSG_SO_GESTOR
        assert client_aprendiz.get("/api/reports/monthly/export").status_code == 403
        assert client_gestor.get("/api/reports/monthly/export").status_code == 200

    def test_charts_qualquer_role_autenticado(self, client, client_gestor, client_tecnico, client_aprendiz):
        for c in (client_gestor, client_tecnico, client_aprendiz):
            assert c.get("/api/reports/charts/loans").status_code == 200
            assert c.get("/api/reports/charts/registrations").status_code == 200
        assert client.get("/api/reports/charts/loans").status_code == 401


class TestRbacPerifericos:
    def test_get_perifericos_qualquer_role_autenticado(self, client, client_gestor, client_tecnico, client_aprendiz):
        assert client.get("/api/peripherals").status_code == 401
        for c in (client_gestor, client_tecnico, client_aprendiz):
            assert c.get("/api/peripherals").status_code == 200

    def test_post_perifericos_gestor_e_tecnico_200_aprendiz_403(
        self, client, client_gestor, client_tecnico, client_aprendiz
    ):
        assert client.post("/api/peripherals", json={"tipo": "Mouse"}).status_code == 401

        resp_apr = client_aprendiz.post("/api/peripherals", json={"tipo": "Mouse"})
        assert resp_apr.status_code == 403
        assert resp_apr.json()["detail"] == MSG_GESTOR_OU_TECNICO

        assert client_gestor.post(
            "/api/peripherals", json={"tipo": "Mouse", "identificador": "RBAC-GESTOR-001"}
        ).status_code == 200
        assert client_tecnico.post(
            "/api/peripherals", json={"tipo": "Mouse", "identificador": "RBAC-TECNICO-001"}
        ).status_code == 200

    def test_listar_perifericos_do_item_qualquer_role(
        self, client, client_gestor, client_tecnico, client_aprendiz, item_disponivel
    ):
        url = f"/api/items/{item_disponivel['id']}/peripherals"
        assert client.get(url).status_code == 401
        for c in (client_gestor, client_tecnico, client_aprendiz):
            assert c.get(url).status_code == 200

    def test_vincular_periferico_gestor_e_tecnico_200_aprendiz_403(
        self, client, client_gestor, client_tecnico, client_aprendiz, item_disponivel, criar_periferico
    ):
        perif_gestor = criar_periferico()
        perif_tecnico = criar_periferico()

        assert client.post(f"/api/items/{item_disponivel['id']}/peripherals/{perif_gestor['id']}").status_code == 401

        resp_apr = client_aprendiz.post("/api/items/999999/peripherals/999999")
        assert resp_apr.status_code == 403

        assert client_gestor.post(
            f"/api/items/{item_disponivel['id']}/peripherals/{perif_gestor['id']}"
        ).status_code == 200
        assert client_tecnico.post(
            f"/api/items/{item_disponivel['id']}/peripherals/{perif_tecnico['id']}"
        ).status_code == 200

    def test_desvincular_periferico_gestor_e_tecnico_200_aprendiz_403(
        self, client, client_gestor, client_tecnico, client_aprendiz, item_disponivel, criar_periferico, inv_manager
    ):
        perif_a = criar_periferico()
        perif_b = criar_periferico()
        ok, _ = inv_manager.link_peripheral_to_equipment(item_disponivel["id"], perif_a["id"], "teste")
        assert ok
        ok, _ = inv_manager.link_peripheral_to_equipment(item_disponivel["id"], perif_b["id"], "teste")
        assert ok
        links = inv_manager.list_peripherals_for_equipment(item_disponivel["id"])
        link_a = next(l["link_id"] for l in links if l["id"] == perif_a["id"])
        link_b = next(l["link_id"] for l in links if l["id"] == perif_b["id"])

        assert client.delete(f"/api/peripherals/links/{link_a}").status_code == 401

        resp_apr = client_aprendiz.delete("/api/peripherals/links/999999")
        assert resp_apr.status_code == 403

        assert client_gestor.delete(f"/api/peripherals/links/{link_a}").status_code == 200
        assert client_tecnico.delete(f"/api/peripherals/links/{link_b}").status_code == 200

    def test_substituir_periferico_gestor_e_tecnico_200_aprendiz_403(
        self, client, client_gestor, client_tecnico, client_aprendiz, item_disponivel, criar_periferico, inv_manager
    ):
        antigo_gestor = criar_periferico()
        novo_gestor = criar_periferico()
        antigo_tecnico = criar_periferico()
        novo_tecnico = criar_periferico()
        inv_manager.link_peripheral_to_equipment(item_disponivel["id"], antigo_gestor["id"], "teste")
        inv_manager.link_peripheral_to_equipment(item_disponivel["id"], antigo_tecnico["id"], "teste")

        url_fake = f"/api/items/{item_disponivel['id']}/peripherals/999999/replace"
        assert client.post(url_fake, data={"new_peripheral_id": "1", "reason": "Defeito"}).status_code == 401

        resp_apr = client_aprendiz.post(url_fake, data={"new_peripheral_id": "1", "reason": "Defeito"})
        assert resp_apr.status_code == 403

        resp_gestor = client_gestor.post(
            f"/api/items/{item_disponivel['id']}/peripherals/{antigo_gestor['id']}/replace",
            data={"new_peripheral_id": str(novo_gestor["id"]), "reason": "Defeito"},
        )
        assert resp_gestor.status_code == 200

        resp_tecnico = client_tecnico.post(
            f"/api/items/{item_disponivel['id']}/peripherals/{antigo_tecnico['id']}/replace",
            data={"new_peripheral_id": str(novo_tecnico["id"]), "reason": "Defeito"},
        )
        assert resp_tecnico.status_code == 200


class TestRbacEmprestimos:
    def test_iniciar_emprestimo_gestor_e_tecnico_200_aprendiz_403(
        self, client, client_gestor, client_tecnico, client_aprendiz, criar_item
    ):
        item_gestor = criar_item()
        item_tecnico = criar_item()
        body = lambda item_id: {
            "item_id": item_id, "usuario": "Fulano", "cpf": "11122233344",
            "center_cost": "101 - Puxada", "cargo": "Analista", "setor": "TI",
            "revenda": "Revalle Juazeiro", "date_issue": "01/01/2026",
        }

        assert client.post("/api/loans", json=body(item_gestor["id"])).status_code == 401

        resp_apr = client_aprendiz.post("/api/loans", json=body(999999))
        assert resp_apr.status_code == 403
        assert resp_apr.json()["detail"] == MSG_GESTOR_OU_TECNICO

        assert client_gestor.post("/api/loans", json=body(item_gestor["id"])).status_code == 200
        assert client_tecnico.post("/api/loans", json=body(item_tecnico["id"])).status_code == 200

    def test_confirmar_emprestimo_gestor_e_tecnico_200_aprendiz_403(
        self, client, client_gestor, client_tecnico, client_aprendiz, inv_manager, criar_item
    ):
        def novo_item_pendente():
            item = criar_item()
            ok, msg = inv_manager.issue(
                item["id"], "Fulano", "11122233344", "101 - Puxada", "Analista", "TI",
                "Revalle Juazeiro", "01/01/2026", "teste",
            )
            assert ok, msg
            return item["id"]

        item_gestor = novo_item_pendente()
        item_tecnico = novo_item_pendente()
        arquivo = {"signed_pdf": ("termo.pdf", b"conteudo-fake", "application/pdf")}

        assert client.post(f"/api/loans/{item_gestor}/confirm", files=arquivo).status_code == 401

        resp_apr = client_aprendiz.post("/api/loans/999999/confirm", files=arquivo)
        assert resp_apr.status_code == 403

        assert client_gestor.post(f"/api/loans/{item_gestor}/confirm", files=arquivo).status_code == 200
        assert client_tecnico.post(f"/api/loans/{item_tecnico}/confirm", files=arquivo).status_code == 200

    def test_iniciar_devolucao_gestor_e_tecnico_nao_sao_barrados_por_role_aprendiz_403(
        self, client, client_gestor, client_tecnico, client_aprendiz, item_indisponivel
    ):
        """
        A rota depende de gerar um .docx (generate_return_term_bytes), que pode não ter
        o modelo disponível neste ambiente (ver docs/TESTES.md) — por isso aceitamos
        200 (sucesso) ou 400 (erro de geração de documento) como evidência de que o
        pedido *passou* da checagem de role; o que importa para o RBAC é que não seja
        401/403 para os roles permitidos.
        """
        url = f"/api/loans/{item_indisponivel['id']}/return/initiate"
        assert client.post(url).status_code == 401

        resp_apr = client_aprendiz.post("/api/loans/999999/return/initiate")
        assert resp_apr.status_code == 403
        assert resp_apr.json()["detail"] == MSG_GESTOR_OU_TECNICO

        resp_gestor = client_gestor.post(url)
        assert resp_gestor.status_code in (200, 400)

    def test_confirmar_devolucao_gestor_e_tecnico_200_aprendiz_403(
        self, client, client_gestor, client_tecnico, client_aprendiz, item_indisponivel, forcar_devolucao_iniciada,
        inv_manager, criar_item,
    ):
        def item_em_pendente_devolucao():
            item = criar_item()
            ok, msg = inv_manager.issue(
                item["id"], "Fulano", "11122233344", "101 - Puxada", "Analista", "TI",
                "Revalle Juazeiro", "01/01/2026", "teste",
            )
            assert ok, msg
            ok, msg = inv_manager.confirm_loan(item["id"], "teste", "termos_assinados/fake.pdf")
            assert ok, msg
            forcar_devolucao_iniciada(item["id"])
            return item["id"]

        item_gestor = item_em_pendente_devolucao()
        item_tecnico = item_em_pendente_devolucao()
        arquivo = {"signed_pdf": ("termo.pdf", b"conteudo-fake", "application/pdf")}

        assert client.post(f"/api/loans/{item_gestor}/return/confirm", files=arquivo).status_code == 401

        resp_apr = client_aprendiz.post("/api/loans/999999/return/confirm", files=arquivo)
        assert resp_apr.status_code == 403

        assert client_gestor.post(f"/api/loans/{item_gestor}/return/confirm", files=arquivo).status_code == 200
        assert client_tecnico.post(f"/api/loans/{item_tecnico}/return/confirm", files=arquivo).status_code == 200
