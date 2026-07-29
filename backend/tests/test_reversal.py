"""Caracterização do estorno (InventoryDBManager.reverse_history_entry / POST /api/history/{id}/reverse)."""
import pytest

pytestmark = pytest.mark.integration

ARQUIVO_PDF = {"signed_pdf": ("termo.pdf", b"conteudo-fake-pdf", "application/pdf")}


def _history_id(inv_manager, item_id: int, operation: str) -> int:
    entradas = [h for h in inv_manager.list_history() if h["item_id"] == item_id and h["operation"] == operation]
    assert entradas, f"Não encontrei entrada de histórico '{operation}' para o item {item_id}."
    return entradas[0]["id"]


class TestEstornoCadastro:
    def test_estorna_cadastro_e_soft_deleta_o_item(self, client_gestor, inv_manager, item_disponivel):
        hid = _history_id(inv_manager, item_disponivel["id"], "Cadastro")
        resp = client_gestor.post(f"/api/history/{hid}/reverse")
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Operação 'Cadastro' do item {item_disponivel['id']} estornada com sucesso."
        assert client_gestor.get(f"/api/items/{item_disponivel['id']}").status_code == 404


class TestEstornoEmprestimo:
    def test_estorna_emprestimo_volta_para_disponivel(self, client_gestor, inv_manager, item_pendente):
        item_id = item_pendente["id"]
        hid = _history_id(inv_manager, item_id, "Empréstimo")
        resp = client_gestor.post(f"/api/history/{hid}/reverse")
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Operação 'Empréstimo' do item {item_id} estornada com sucesso."
        item = client_gestor.get(f"/api/items/{item_id}").json()
        assert item["status"] == "Disponível"
        assert item["assigned_to"] is None
        assert item["cpf"] is None


class TestEstornoConfirmacaoEmprestimo:
    def test_estorna_confirmacao_emprestimo_volta_para_pendente(self, client_gestor, inv_manager, item_indisponivel):
        item_id = item_indisponivel["id"]
        hid = _history_id(inv_manager, item_id, "Confirmação Empréstimo")
        resp = client_gestor.post(f"/api/history/{hid}/reverse")
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Operação 'Confirmação Empréstimo' do item {item_id} estornada com sucesso."
        item = client_gestor.get(f"/api/items/{item_id}").json()
        assert item["status"] == "Pendente"
        # assigned_to/cpf não são alterados por este ramo do estorno (só o status).
        assert item["assigned_to"] == "Fulano de Tal"


class TestEstornoDevolucao:
    def test_estorna_devolucao_volta_para_indisponivel_com_dados_do_ultimo_emprestimo(
        self, client_gestor, inv_manager, item_indisponivel, forcar_devolucao_iniciada
    ):
        item_id = item_indisponivel["id"]
        forcar_devolucao_iniciada(item_id)
        hid = _history_id(inv_manager, item_id, "Devolução")

        resp = client_gestor.post(f"/api/history/{hid}/reverse")
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Operação 'Devolução' do item {item_id} estornada com sucesso."
        item = client_gestor.get(f"/api/items/{item_id}").json()
        assert item["status"] == "Indisponível"
        # Restaurado a partir do último registro de Empréstimo/Confirmação Empréstimo
        # anterior ao lançamento de Devolução (aqui, a Confirmação Empréstimo).
        assert item["assigned_to"] == "Fulano de Tal"
        assert item["cpf"] == "11122233344"

    def test_estornar_devolucao_sem_emprestimo_anterior_no_historico(
        self, client_gestor, inv_manager, item_indisponivel, forcar_devolucao_iniciada
    ):
        """
        Caso extremo: se as entradas de Empréstimo/Confirmação Empréstimo anteriores já
        tiverem sido apagadas/estornadas antes da Devolução, reverse_history_entry não
        encontra o empréstimo original e recusa o estorno com uma mensagem própria.
        Aqui simulamos isso estornando primeiro a Confirmação Empréstimo e o Empréstimo,
        e só então tentando estornar a Devolução.
        """
        item_id = item_indisponivel["id"]
        forcar_devolucao_iniciada(item_id)
        hid_confirmacao = _history_id(inv_manager, item_id, "Confirmação Empréstimo")
        hid_emprestimo = _history_id(inv_manager, item_id, "Empréstimo")
        hid_devolucao = _history_id(inv_manager, item_id, "Devolução")

        ok, msg = inv_manager.reverse_history_entry(hid_confirmacao, "teste")
        assert ok, msg
        ok, msg = inv_manager.reverse_history_entry(hid_emprestimo, "teste")
        assert ok, msg

        resp = client_gestor.post(f"/api/history/{hid_devolucao}/reverse")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Não foi possível encontrar o empréstimo original."


class TestEstornoConfirmacaoDevolucao:
    def test_estorna_confirmacao_devolucao_volta_para_pendente_devolucao(
        self, client_gestor, inv_manager, item_indisponivel, forcar_devolucao_iniciada
    ):
        item_id = item_indisponivel["id"]
        forcar_devolucao_iniciada(item_id)
        ok, msg = inv_manager.confirm_return(item_id, "teste", "termos_devolucao_assinados/fake.pdf")
        assert ok, msg
        hid = _history_id(inv_manager, item_id, "Confirmação Devolução")

        resp = client_gestor.post(f"/api/history/{hid}/reverse")
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Operação 'Confirmação Devolução' do item {item_id} estornada com sucesso."
        assert client_gestor.get(f"/api/items/{item_id}").json()["status"] == "Pendente Devolução"


class TestEstornoCasosGerais:
    def test_estornar_duas_vezes_a_mesma_entrada(self, client_gestor, inv_manager, item_pendente):
        hid = _history_id(inv_manager, item_pendente["id"], "Empréstimo")
        primeira = client_gestor.post(f"/api/history/{hid}/reverse")
        assert primeira.status_code == 200

        segunda = client_gestor.post(f"/api/history/{hid}/reverse")
        assert segunda.status_code == 400
        assert segunda.json()["detail"] == "Esta operação já foi estornada."

    def test_estornar_operacao_nao_estornavel_edicao(self, client_gestor, inv_manager, item_disponivel):
        ok, msg = inv_manager.update_item(item_disponivel["id"], {"brand": "Outra"}, "teste")
        assert ok, msg
        hid = _history_id(inv_manager, item_disponivel["id"], "Edição")

        resp = client_gestor.post(f"/api/history/{hid}/reverse")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Não é possível estornar uma operação do tipo 'Edição'."

    def test_estornar_lancamento_inexistente(self, client_gestor):
        resp = client_gestor.post("/api/history/999999/reverse")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Lançamento não encontrado."
