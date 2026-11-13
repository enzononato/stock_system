"""Caracterização do estorno (InventoryDBManager.reverse_history_entry / POST /api/history/{id}/reverse).

ATENÇÃO — BUG DE PRODUTO ENCONTRADO NESTE CICLO (reportado, não corrigido; fora da
posse deste módulo, que só tem backend/tests/**, backend/pytest.ini e docs/TESTES.md):

`app/routers/history.py::reverse_entry()` (endpoint POST /api/history/{id}/reverse,
feature nova "estorno com senha") busca o usuário logado com
`user_db.get_user_by_id(current_user.id)` para verificar a senha:

    user = user_db.get_user_by_id(current_user.id)
    if not user or not verify_password(body.password, user["password"]):
        ...

Só que `UserDBManager.get_user_by_id()` seleciona DE PROPÓSITO sem a coluna
`password` (`"SELECT id, username, role FROM usuarios WHERE id = %s"` — o próprio
docstring do método diz "sem o hash da senha"; é o método usado por
`users.py::remove_user` para a checagem de lockout, que não precisa de senha).
O método certo para obter o hash da senha é `get_user_by_username()` (usado
corretamente em `auth.py::login()`), que SELECIONA `password`.

Resultado: `user["password"]` sempre levanta `KeyError: 'password'`, e todo
POST /api/history/{id}/reverse autenticado como Gestor — com senha certa OU
errada — quebra com 500 Internal Server Error em vez de 200/403. A feature
"estorno com senha" (T6) está, no estado atual do código, completamente
inoperante via HTTP.

Por isso, as caracterizações da MÁQUINA DE ESTADOS do estorno abaixo chamam
`InventoryDBManager.reverse_history_entry()` diretamente (camada não afetada por
este bug — o bug está só no router, na checagem de senha) em vez de passar pelo
endpoint HTTP quebrado. A classe `TestBugDeProdutoSenhaDoEstorno`, no fim deste
arquivo, documenta e comprova o bug via HTTP com testes `xfail` dedicados (ver
docs/TESTES.md, seção de bugs de produto encontrados nesta rodada).
"""
import pytest

pytestmark = pytest.mark.integration

ARQUIVO_PDF = {"signed_pdf": ("termo.pdf", b"conteudo-fake-pdf", "application/pdf")}


def _history_id(inv_manager, item_id: int, operation: str) -> int:
    # MUDANÇA INTENCIONAL (T10, paginação server-side): InventoryDBManager.list_history()
    # deixou de devolver só a lista de linhas e passou a devolver a tupla
    # (linhas, total) — o "total" é necessário para o envelope {"items", "total"}
    # de GET /api/history. Este helper precisa desempacotar a tupla antes de
    # filtrar; sem isso, iterar a tupla devolve a lista e depois o int, e
    # `h["item_id"]` sobre a lista estoura TypeError.
    linhas, _total = inv_manager.list_history()
    entradas = [h for h in linhas if h["item_id"] == item_id and h["operation"] == operation]
    assert entradas, f"Não encontrei entrada de histórico '{operation}' para o item {item_id}."
    return entradas[0]["id"]


class TestEstornoCadastro:
    def test_estorna_cadastro_e_soft_deleta_o_item(self, client_gestor, inv_manager, item_disponivel):
        hid = _history_id(inv_manager, item_disponivel["id"], "Cadastro")
        ok, msg = inv_manager.reverse_history_entry(hid, "teste")
        assert ok, msg
        assert msg == f"Operação 'Cadastro' do item {item_disponivel['id']} estornada com sucesso."
        assert client_gestor.get(f"/api/items/{item_disponivel['id']}").status_code == 404


class TestEstornoEmprestimo:
    def test_estorna_emprestimo_volta_para_disponivel(self, client_gestor, inv_manager, item_pendente):
        item_id = item_pendente["id"]
        hid = _history_id(inv_manager, item_id, "Empréstimo")
        ok, msg = inv_manager.reverse_history_entry(hid, "teste")
        assert ok, msg
        assert msg == f"Operação 'Empréstimo' do item {item_id} estornada com sucesso."
        item = client_gestor.get(f"/api/items/{item_id}").json()
        assert item["status"] == "Disponível"
        assert item["assigned_to"] is None
        assert item["cpf"] is None


class TestEstornoConfirmacaoEmprestimo:
    def test_estorna_confirmacao_emprestimo_volta_para_pendente(self, client_gestor, inv_manager, item_indisponivel):
        item_id = item_indisponivel["id"]
        hid = _history_id(inv_manager, item_id, "Confirmação Empréstimo")
        ok, msg = inv_manager.reverse_history_entry(hid, "teste")
        assert ok, msg
        assert msg == f"Operação 'Confirmação Empréstimo' do item {item_id} estornada com sucesso."
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

        ok, msg = inv_manager.reverse_history_entry(hid, "teste")
        assert ok, msg
        assert msg == f"Operação 'Devolução' do item {item_id} estornada com sucesso."
        item = client_gestor.get(f"/api/items/{item_id}").json()
        assert item["status"] == "Indisponível"
        # Restaurado a partir do último registro de Empréstimo/Confirmação Empréstimo
        # anterior ao lançamento de Devolução (aqui, a Confirmação Empréstimo).
        assert item["assigned_to"] == "Fulano de Tal"
        assert item["cpf"] == "11122233344"

    def test_estornar_devolucao_sem_emprestimo_anterior_no_historico(
        self, client_gestor, inv_manager, item_indisponivel, forcar_devolucao_iniciada, executar_sql
    ):
        """
        Caso extremo: se as entradas de Empréstimo/Confirmação Empréstimo anteriores já
        tiverem sido apagadas/estornadas antes da Devolução, reverse_history_entry não
        encontra o empréstimo original e recusa o estorno com uma mensagem própria.
        Aqui simulamos isso marcando como estornados a Confirmação Empréstimo e o
        Empréstimo direto no banco (pela API o estorno deles é recusado enquanto o item
        está em Pendente Devolução), e só então tentando estornar a Devolução.
        """
        item_id = item_indisponivel["id"]
        forcar_devolucao_iniciada(item_id)
        hid_devolucao = _history_id(inv_manager, item_id, "Devolução")
        executar_sql(
            "UPDATE history SET is_reversed=1 WHERE item_id=%s "
            "AND operation IN ('Empréstimo','Confirmação Empréstimo')",
            (item_id,),
        )

        ok, msg = inv_manager.reverse_history_entry(hid_devolucao, "teste")
        assert not ok
        assert msg == "Não foi possível encontrar o empréstimo original."


class TestEstornoConfirmacaoDevolucao:
    def test_estorna_confirmacao_devolucao_volta_para_pendente_devolucao(
        self, client_gestor, inv_manager, item_indisponivel, forcar_devolucao_iniciada
    ):
        item_id = item_indisponivel["id"]
        forcar_devolucao_iniciada(item_id)
        ok, msg = inv_manager.confirm_return(item_id, "teste", "termos_devolucao_assinados/fake.pdf")
        assert ok, msg
        hid = _history_id(inv_manager, item_id, "Confirmação Devolução")

        ok, msg = inv_manager.reverse_history_entry(hid, "teste")
        assert ok, msg
        assert msg == f"Operação 'Confirmação Devolução' do item {item_id} estornada com sucesso."
        assert client_gestor.get(f"/api/items/{item_id}").json()["status"] == "Pendente Devolução"


class TestEstornoCasosGerais:
    def test_estornar_duas_vezes_a_mesma_entrada(self, client_gestor, inv_manager, item_pendente):
        hid = _history_id(inv_manager, item_pendente["id"], "Empréstimo")
        ok, _ = inv_manager.reverse_history_entry(hid, "teste")
        assert ok

        ok, msg = inv_manager.reverse_history_entry(hid, "teste")
        assert not ok
        assert msg == "Esta operação já foi estornada."

    def test_estornar_operacao_nao_estornavel_edicao(self, client_gestor, inv_manager, item_disponivel):
        ok, msg = inv_manager.update_item(item_disponivel["id"], {"brand": "Outra"}, "teste")
        assert ok, msg
        hid = _history_id(inv_manager, item_disponivel["id"], "Edição")

        ok, msg = inv_manager.reverse_history_entry(hid, "teste")
        assert not ok
        assert msg == "Não é possível estornar uma operação do tipo 'Edição'."

    def test_estornar_lancamento_inexistente(self, client_gestor, inv_manager):
        ok, msg = inv_manager.reverse_history_entry(999999, "teste")
        assert not ok
        assert msg == "Lançamento não encontrado."


class TestEstornoComSenhaSchemaHttp:
    """
    T2 (novo, "estorno com senha"): a parte do contrato HTTP que NÃO depende do
    código com bug (a validação do corpo pelo schema ReverseRequest acontece antes
    de reverse_entry() sequer rodar, então não chega perto de user["password"]).
    """

    def test_sem_o_campo_password_retorna_422(self, client_gestor, item_disponivel, inv_manager):
        hid = _history_id(inv_manager, item_disponivel["id"], "Cadastro")
        resp = client_gestor.post(f"/api/history/{hid}/reverse", json={})
        assert resp.status_code == 422

    def test_password_em_branco_retorna_422(self, client_gestor, item_disponivel, inv_manager):
        hid = _history_id(inv_manager, item_disponivel["id"], "Cadastro")
        resp = client_gestor.post(f"/api/history/{hid}/reverse", json={"password": ""})
        assert resp.status_code == 422


class TestSenhaDoEstorno:
    """
    Regressão de um bug que quebrava o estorno por completo.

    `reverse_entry()` checava a senha com `user_db.get_user_by_id(...)`, cujo
    SELECT não inclui a coluna `password` (por design — ver
    `UserDBManager.get_user_by_id`). O acesso a `user["password"]` levantava
    KeyError, então TODA chamada autenticada respondia 500, com senha certa ou
    errada: a funcionalidade de estorno estava inteiramente inoperante. A
    correção passou a usar `get_user_by_username(current_user.username)`, o
    mesmo método que `auth.py::login()` já usava corretamente.

    Estes testes cobrem os dois lados da checagem, para que o endpoint não volte
    a ser publicado sem que a validação de senha funcione de fato.
    """

    def test_senha_correta_estorna_a_operacao(
        self, client_gestor, senha_padrao_teste, inv_manager, item_disponivel
    ):
        hid = _history_id(inv_manager, item_disponivel["id"], "Cadastro")
        resp = client_gestor.post(f"/api/history/{hid}/reverse", json={"password": senha_padrao_teste})
        assert resp.status_code == 200, resp.text

    def test_senha_errada_da_403_e_nao_estorna(
        self, client_gestor, inv_manager, item_disponivel
    ):
        hid = _history_id(inv_manager, item_disponivel["id"], "Cadastro")
        resp = client_gestor.post(f"/api/history/{hid}/reverse", json={"password": "senha-errada-com-certeza"})
        assert resp.status_code == 403
        assert resp.json()["detail"] == "Senha incorreta. Ação não autorizada."

        # O item precisa continuar disponível: senha errada não pode ter efeito colateral.
        item = inv_manager.find(item_disponivel["id"])
        assert item is not None, "O item foi removido apesar de a senha estar errada."


MSG_MUDOU = "O equipamento mudou de situação depois deste lançamento."


class TestEstornoConfereSituacaoAtual:
    """O estorno só desfaz um lançamento se o equipamento ainda está na situação que
    aquele lançamento produziu; senão recusa (400) sem mexer em nada."""

    def _ciclo_devolvido(self, inv_manager, item_id, forcar_devolucao_iniciada):
        forcar_devolucao_iniciada(item_id)
        ok, msg = inv_manager.confirm_return(item_id, "teste", "termos_devolucao_assinados/fake.pdf")
        assert ok, msg
        return _history_id(inv_manager, item_id, "Confirmação Devolução")

    def test_confirmacao_devolucao_antiga_com_manutencao_aberta_400_e_ordem_segue(
        self, client_gestor, senha_padrao_teste, inv_manager, item_indisponivel, forcar_devolucao_iniciada
    ):
        item_id = item_indisponivel["id"]
        hid = self._ciclo_devolvido(inv_manager, item_id, forcar_devolucao_iniciada)
        resp = client_gestor.post("/api/manutencoes", json={"item_id": item_id, "defeito": "Tela"})
        assert resp.status_code == 200, resp.text
        mid = resp.json()["id"]

        resp = client_gestor.post(f"/api/history/{hid}/reverse", json={"password": senha_padrao_teste})
        assert resp.status_code == 400
        assert resp.json()["detail"] == MSG_MUDOU
        assert inv_manager.get_history_entry(hid)["is_reversed"] == 0
        assert inv_manager.find(item_id)["status"] == "Em Manutenção"

        resp = client_gestor.post(f"/api/manutencoes/{mid}/retorno", data={"resultado": "retornada"})
        assert resp.status_code == 200, resp.text
        assert inv_manager.find(item_id)["status"] == "Disponível"

    def test_confirmacao_devolucao_antiga_permite_cancelar_a_ordem(
        self, client_gestor, inv_manager, item_indisponivel, forcar_devolucao_iniciada
    ):
        item_id = item_indisponivel["id"]
        hid = self._ciclo_devolvido(inv_manager, item_id, forcar_devolucao_iniciada)
        mid = client_gestor.post("/api/manutencoes", json={"item_id": item_id, "defeito": "Tela"}).json()["id"]

        assert inv_manager.reverse_history_entry(hid, "teste") == (False, MSG_MUDOU)
        resp = client_gestor.post(f"/api/manutencoes/{mid}/cancelar", json={"motivo": "Engano"})
        assert resp.status_code == 200, resp.text
        assert inv_manager.find(item_id)["status"] == "Disponível"

    def test_confirmacao_devolucao_com_ordem_em_andamento_e_item_disponivel_recusa(
        self, inv_manager, item_indisponivel, forcar_devolucao_iniciada, executar_sql
    ):
        item_id = item_indisponivel["id"]
        hid = self._ciclo_devolvido(inv_manager, item_id, forcar_devolucao_iniciada)
        executar_sql(
            "INSERT INTO manutencoes (item_id, status, defeito, aberta_por, aberta_em) "
            "VALUES (%s,'aberta','x','teste',NOW())",
            (item_id,),
        )
        assert inv_manager.reverse_history_entry(hid, "teste") == (False, MSG_MUDOU)

    def test_cadastro_de_item_baixado_400(
        self, client_gestor, senha_padrao_teste, inv_manager, item_disponivel
    ):
        item_id = item_disponivel["id"]
        hid = _history_id(inv_manager, item_id, "Cadastro")
        ok, msg = inv_manager.remove(item_id, "teste", "Obsolescência")
        assert ok, msg

        resp = client_gestor.post(f"/api/history/{hid}/reverse", json={"password": senha_padrao_teste})
        assert resp.status_code == 400
        assert resp.json()["detail"] == MSG_MUDOU
        assert inv_manager.get_history_entry(hid)["is_reversed"] == 0

    def test_confirmacao_devolucao_de_item_baixado_400(
        self, inv_manager, item_indisponivel, forcar_devolucao_iniciada
    ):
        item_id = item_indisponivel["id"]
        hid = self._ciclo_devolvido(inv_manager, item_id, forcar_devolucao_iniciada)
        ok, msg = inv_manager.remove(item_id, "teste", "Obsolescência")
        assert ok, msg
        assert inv_manager.reverse_history_entry(hid, "teste") == (False, MSG_MUDOU)
        assert inv_manager.get_history_entry(hid)["is_reversed"] == 0

    def test_cadastro_de_item_emprestado_400(self, inv_manager, item_pendente):
        hid = _history_id(inv_manager, item_pendente["id"], "Cadastro")
        assert inv_manager.reverse_history_entry(hid, "teste") == (False, MSG_MUDOU)
        assert inv_manager.find(item_pendente["id"])["status"] == "Pendente"

    def test_emprestimo_ja_confirmado_400(self, inv_manager, item_indisponivel):
        hid = _history_id(inv_manager, item_indisponivel["id"], "Empréstimo")
        assert inv_manager.reverse_history_entry(hid, "teste") == (False, MSG_MUDOU)
        item = inv_manager.find(item_indisponivel["id"])
        assert item["status"] == "Indisponível"
        assert item["assigned_to"] == "Fulano de Tal"

    def test_confirmacao_emprestimo_com_devolucao_iniciada_400(
        self, inv_manager, item_indisponivel, forcar_devolucao_iniciada
    ):
        item_id = item_indisponivel["id"]
        forcar_devolucao_iniciada(item_id)
        hid = _history_id(inv_manager, item_id, "Confirmação Empréstimo")
        assert inv_manager.reverse_history_entry(hid, "teste") == (False, MSG_MUDOU)
        assert inv_manager.find(item_id)["status"] == "Pendente Devolução"

    def test_devolucao_ja_confirmada_400(self, inv_manager, item_indisponivel, forcar_devolucao_iniciada):
        item_id = item_indisponivel["id"]
        self._ciclo_devolvido(inv_manager, item_id, forcar_devolucao_iniciada)
        hid = _history_id(inv_manager, item_id, "Devolução")
        assert inv_manager.reverse_history_entry(hid, "teste") == (False, MSG_MUDOU)
        assert inv_manager.find(item_id)["status"] == "Disponível"


class TestEmprestimoSoDeItemDisponivel:
    def test_issue_recusa_quando_o_status_mudou_depois_da_leitura(
        self, inv_manager, item_disponivel, executar_sql, monkeypatch
    ):
        """`issue()` lê o item antes de gravar; se outro fluxo mudou o status nesse
        meio tempo, o UPDATE condicional não casa e nada é gravado."""
        from datetime import datetime

        item_id = item_disponivel["id"]
        lido = inv_manager.find(item_id)
        executar_sql("UPDATE items SET status='Em Manutenção' WHERE id=%s", (item_id,))
        monkeypatch.setattr(inv_manager, "find", lambda _id: dict(lido))

        ok, msg = inv_manager.issue(
            item_id, "Fulano de Tal", "11122233344", "101 - Puxada", "Analista", "TI",
            "Outra Unidade", datetime.now().strftime("%d/%m/%Y"), "teste",
        )
        assert (ok, msg) == (False, "Este item não está disponível para empréstimo.")
        from app.db.database_mysql import get_cursor

        with get_cursor(commit=False) as cur:
            cur.execute("SELECT status, assigned_to FROM items WHERE id=%s", (item_id,))
            assert cur.fetchone() == {"status": "Em Manutenção", "assigned_to": None}
            cur.execute("SELECT COUNT(*) AS n FROM history WHERE item_id=%s AND operation='Empréstimo'", (item_id,))
            assert cur.fetchone()["n"] == 0
            cur.execute("SELECT COUNT(*) AS n FROM transferencias WHERE item_id=%s", (item_id,))
            assert cur.fetchone()["n"] == 0
