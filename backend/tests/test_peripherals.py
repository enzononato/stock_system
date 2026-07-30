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


class TestInativarPeriferico:
    """T2 (novo): DELETE /api/peripherals/{id} (T9 no código) inativa (soft-delete)
    um periférico disponível; recusa com 400 se o periférico estiver 'Em Uso' ou
    vinculado a um equipamento; RBAC permite Gestor ou Técnico (gestor_or_tecnico)."""

    def test_inativa_periferico_disponivel(self, client_gestor, periferico_disponivel):
        resp = client_gestor.delete(f"/api/peripherals/{periferico_disponivel['id']}")
        assert resp.status_code == 200
        assert resp.json()["detail"] == (
            f"Periférico 'Mouse' (ID {periferico_disponivel['id']}) removido do estoque."
        )
        # Soft-delete: some da listagem padrão, mas continua existindo com include_inactive.
        ids_ativos = {p["id"] for p in client_gestor.get("/api/peripherals").json()}
        assert periferico_disponivel["id"] not in ids_ativos
        todos = client_gestor.get("/api/peripherals", params={"include_inactive": True}).json()
        inativo = next(p for p in todos if p["id"] == periferico_disponivel["id"])
        assert inativo["status"] == "Disponível"  # status não muda, só is_active

    def test_recusa_com_400_se_periferico_estiver_em_uso(
        self, client_gestor, item_disponivel, periferico_disponivel
    ):
        """
        deactivate_peripheral() checa 'Em Uso' ANTES de checar o vínculo em
        equipment_peripherals -- e link_peripheral_to_equipment() já marca o
        periférico como 'Em Uso' no momento do vínculo. Por isso, vincular um
        periférico dispara a primeira checagem (mensagem "...em uso.") antes da
        segunda ("...vinculado a um equipamento.") ter a chance de rodar; a
        segunda mensagem existiria, na prática, para um estado inconsistente em
        que o periférico está vinculado mas não está mais 'Em Uso'.
        """
        link = client_gestor.post(
            f"/api/items/{item_disponivel['id']}/peripherals/{periferico_disponivel['id']}"
        )
        assert link.status_code == 200

        resp = client_gestor.delete(f"/api/peripherals/{periferico_disponivel['id']}")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Não é possível remover periférico em uso."

    def test_periferico_inexistente_retorna_400(self, client_gestor):
        resp = client_gestor.delete("/api/peripherals/999999")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Periférico não encontrado."

    def test_rbac_gestor_e_tecnico_podem_aprendiz_nao(
        self, client, client_gestor, client_tecnico, client_aprendiz, criar_periferico
    ):
        perif_gestor = criar_periferico(identificador="RBAC-DEL-GESTOR")
        perif_tecnico = criar_periferico(identificador="RBAC-DEL-TECNICO")

        assert client.delete(f"/api/peripherals/{perif_gestor['id']}").status_code == 401

        resp_apr = client_aprendiz.delete(f"/api/peripherals/{perif_gestor['id']}")
        assert resp_apr.status_code == 403
        assert resp_apr.json()["detail"] == "Acesso restrito a Gestor ou Técnico."

        assert client_gestor.delete(f"/api/peripherals/{perif_gestor['id']}").status_code == 200
        assert client_tecnico.delete(f"/api/peripherals/{perif_tecnico['id']}").status_code == 200


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

    def test_vincular_o_mesmo_par_duas_vezes_retorna_mensagem_amigavel(
        self, client_gestor, item_disponivel, periferico_disponivel
    ):
        """
        CORRIGIDO NESTE CICLO (era o BUG CONHECIDO nº 5, ver docs/TESTES.md):
        `link_peripheral_to_equipment()` capturava a violação do UNIQUE
        (equipment_id, peripheral_id) e devolvia a exceção crua do pymysql/MySQL,
        prefixada só por "Erro ao vincular: " — algo como
        "Erro ao vincular: (1062, \"Duplicate entry...\")". O código atual já trata
        `pymysql.MySQLError` de forma genérica (mesmo padrão usado em outros métodos
        deste manager) e devolve uma mensagem limpa em português, sem vazar a exceção
        do driver. Não é uma mensagem específica para "já vinculado" (ainda não
        distingue esse UNIQUE do resto dos erros de banco, ao contrário de
        add_peripheral() com o identificador duplicado) — mas não há mais vazamento
        técnico, que era o bug real caracterizado aqui.
        """
        url = f"/api/items/{item_disponivel['id']}/peripherals/{periferico_disponivel['id']}"
        primeira = client_gestor.post(url)
        assert primeira.status_code == 200

        segunda = client_gestor.post(url)
        assert segunda.status_code == 400
        # Passou a distinguir a violação do UNIQUE (erro 1062) dos demais erros de
        # banco, como add_peripheral() já fazia para o identificador duplicado.
        assert segunda.json()["detail"] == "Este periférico já está vinculado a este equipamento."

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

    def test_confirmar_devolucao_libera_periferico_e_remove_o_vinculo(
        self, client_gestor, inv_manager, item_indisponivel, periferico_disponivel, forcar_devolucao_iniciada
    ):
        """
        CORRIGIDO NESTE CICLO (era o item nº 4/16 de "Mudanças esperadas", ver
        docs/TESTES.md): antes, `confirm_return()` atualizava o status de cada
        periférico vinculado para 'Disponível', mas não apagava a linha em
        `equipment_peripherals` — o periférico ficava "Disponível" e, ao mesmo
        tempo, continuava aparecendo como vinculado ao equipamento devolvido em
        GET /api/items/{id}/peripherals. O código atual
        (`InventoryDBManager.confirm_return`) passou a também fazer
        `DELETE FROM equipment_peripherals WHERE equipment_id=%s` e registrar
        'Desvínculo Periférico' no histórico para cada periférico que estava
        vinculado — este teste agora afirma o comportamento CORRIGIDO e funciona
        como teste de regressão para que o vínculo não volte a "vazar" após a
        devolução.
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
        assert perif["status"] == "Disponível"

        vinculados = client_gestor.get(f"/api/items/{item_indisponivel['id']}/peripherals").json()
        ids_vinculados = {p["id"] for p in vinculados}
        assert periferico_disponivel["id"] not in ids_vinculados  # vínculo removido (corrigido)

        # Regressão: o desvínculo automático fica registrado no histórico.
        linhas, _total = inv_manager.list_history()
        desvinculos = [
            h for h in linhas
            if h["item_id"] == item_indisponivel["id"]
            and h["peripheral_id"] == periferico_disponivel["id"]
            and h["operation"] == "Desvínculo Periférico"
        ]
        assert desvinculos, "Esperava um registro 'Desvínculo Periférico' no histórico após a devolução."
