"""Caracterização de periféricos: vincular, desvincular, substituir, duplicidade e o
efeito de confirm_loan/confirm_return sobre o status dos periféricos."""
import pytest

pytestmark = pytest.mark.integration

ARQUIVO_PDF = {"signed_pdf": ("termo.pdf", b"conteudo-fake-pdf", "application/pdf")}


class TestCadastroDePeriferico:
    def test_identificador_duplicado_retorna_mensagem_amigavel_nao_500(self, client_gestor):
        corpo = {"tipo": "Mouse", "brand": "Logitech", "model": "M90", "identificador": "DUP-001"}
        primeira = client_gestor.post("/api/peripherals", json=corpo)
        assert primeira.status_code == 200

        segunda = client_gestor.post("/api/peripherals", json=corpo)
        assert segunda.status_code == 400
        assert segunda.json()["detail"] == "Já existe um periférico com este Identificador (Nº de Série)."


class TestVincularEDesvincular:
    def test_vincular_periferico_marca_em_uso(self, client_gestor, item_disponivel, periferico_disponivel):
        resp = client_gestor.post(
            f"/api/items/{item_disponivel['id']}/peripherals/{periferico_disponivel['id']}"
        )
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Periférico vinculado com sucesso."

        listagem = client_gestor.get("/api/peripherals").json()
        perif = next(p for p in listagem if p["id"] == periferico_disponivel["id"])
        assert perif["status"] == "Em Uso"

    def test_vincular_o_mesmo_par_duas_vezes_retorna_erro_tecnico_nao_amigavel(
        self, client_gestor, item_disponivel, periferico_disponivel
    ):
        """
        BUG CONHECIDO: link_peripheral_to_equipment() não trata a violação do UNIQUE
        (equipment_id, peripheral_id) de forma amigável (ao contrário de add_peripheral,
        que trata o UNIQUE do identificador com uma mensagem clara). O usuário recebe a
        exceção crua do pymysql/MySQL prefixada só por "Erro ao vincular: ".
        """
        url = f"/api/items/{item_disponivel['id']}/peripherals/{periferico_disponivel['id']}"
        primeira = client_gestor.post(url)
        assert primeira.status_code == 200

        segunda = client_gestor.post(url)
        assert segunda.status_code == 400
        assert segunda.json()["detail"].startswith("Erro ao vincular:")

    def test_desvincular_periferico_marca_disponivel(
        self, client_gestor, item_disponivel, periferico_disponivel
    ):
        client_gestor.post(f"/api/items/{item_disponivel['id']}/peripherals/{periferico_disponivel['id']}")
        link_id = client_gestor.get(f"/api/items/{item_disponivel['id']}/peripherals").json()[0]["link_id"]

        resp = client_gestor.delete(f"/api/peripherals/links/{link_id}")
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Periférico desvinculado com sucesso."

        listagem = client_gestor.get("/api/peripherals").json()
        perif = next(p for p in listagem if p["id"] == periferico_disponivel["id"])
        assert perif["status"] == "Disponível"
        assert client_gestor.get(f"/api/items/{item_disponivel['id']}/peripherals").json() == []

    def test_desvincular_link_inexistente(self, client_gestor):
        resp = client_gestor.delete("/api/peripherals/links/999999")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Vínculo não encontrado."


class TestSubstituirPeriferico:
    def test_substituicao_marca_antigo_substituido_e_novo_em_uso(
        self, client_gestor, item_disponivel, criar_periferico
    ):
        antigo = criar_periferico(identificador="ANTIGO-001")
        novo = criar_periferico(identificador="NOVO-001")
        client_gestor.post(f"/api/items/{item_disponivel['id']}/peripherals/{antigo['id']}")

        resp = client_gestor.post(
            f"/api/items/{item_disponivel['id']}/peripherals/{antigo['id']}/replace",
            data={"new_peripheral_id": str(novo["id"]), "reason": "Defeito de fábrica"},
        )
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Substituição realizada com sucesso."

        todos = client_gestor.get("/api/peripherals", params={"include_inactive": True}).json()
        antigo_atualizado = next(p for p in todos if p["id"] == antigo["id"])
        novo_atualizado = next(p for p in todos if p["id"] == novo["id"])
        assert antigo_atualizado["status"] == "Substituido"
        assert antigo_atualizado["motivo_substituicao"] == "Defeito de fábrica"
        assert novo_atualizado["status"] == "Em Uso"

        vinculados = client_gestor.get(f"/api/items/{item_disponivel['id']}/peripherals").json()
        ids_vinculados = {p["id"] for p in vinculados}
        assert ids_vinculados == {novo["id"]}


class TestEfeitoDoFluxoDeEmprestimoSobrePerifericos:
    def test_confirmar_emprestimo_mantem_periferico_vinculado_em_uso(
        self, client_gestor, inv_manager, item_pendente, periferico_disponivel
    ):
        ok, msg = inv_manager.link_peripheral_to_equipment(item_pendente["id"], periferico_disponivel["id"], "teste")
        assert ok, msg

        resp = client_gestor.post(f"/api/loans/{item_pendente['id']}/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 200

        listagem = client_gestor.get("/api/peripherals").json()
        perif = next(p for p in listagem if p["id"] == periferico_disponivel["id"])
        assert perif["status"] == "Em Uso"

    def test_confirmar_devolucao_libera_periferico_mas_nao_remove_o_vinculo(
        self, client_gestor, inv_manager, item_indisponivel, periferico_disponivel, forcar_devolucao_iniciada
    ):
        """
        BUG CONHECIDO (item 16): confirm_return() atualiza o status de cada periférico
        vinculado para 'Disponível', mas não executa nenhum DELETE em
        equipment_peripherals. Ou seja, depois da devolução confirmada, o periférico
        aparece como "Disponível" (correto) só que CONTINUA aparecendo como vinculado
        àquele equipamento em GET /api/items/{id}/peripherals (incorreto) — o vínculo
        nunca é desfeito automaticamente pela devolução, só por um desvincular manual.
        """
        ok, msg = inv_manager.link_peripheral_to_equipment(
            item_indisponivel["id"], periferico_disponivel["id"], "teste"
        )
        assert ok, msg
        forcar_devolucao_iniciada(item_indisponivel["id"])

        resp = client_gestor.post(f"/api/loans/{item_indisponivel['id']}/return/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 200

        listagem = client_gestor.get("/api/peripherals").json()
        perif = next(p for p in listagem if p["id"] == periferico_disponivel["id"])
        assert perif["status"] == "Disponível"  # correto

        vinculados = client_gestor.get(f"/api/items/{item_indisponivel['id']}/peripherals").json()
        ids_vinculados = {p["id"] for p in vinculados}
        assert periferico_disponivel["id"] in ids_vinculados  # BUG CONHECIDO (item 16)
