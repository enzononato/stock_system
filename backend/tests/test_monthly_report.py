"""
Caracterização do relatório mensal (InventoryDBManager.generate_monthly_report),
um UNION ALL de 4 blocos (Empréstimo, Cadastro, Periféricos, Exclusão) com subqueries
correlacionadas para achar data_confirmacao/data_devolucao. É a query mais frágil do
sistema — qualquer mudança na forma das colunas ou no cálculo dessas datas deve
aparecer aqui.
"""
from datetime import datetime

import pytest

pytestmark = pytest.mark.integration


class TestRelatorioMensal:
    def test_relatorio_agrega_os_quatro_blocos_e_preenche_datas_correlacionadas(
        self, client_gestor, inv_manager, criar_item, criar_periferico, forcar_devolucao_iniciada
    ):
        agora = datetime.now()

        # Bloco "Empréstimo": item que percorre o fluxo inteiro (empréstimo, confirmação,
        # devolução, confirmação de devolução) dentro do mesmo mês.
        item_emprestimo = criar_item()
        ok, msg = inv_manager.issue(
            item_emprestimo["id"], "Ciclano", "22233344455", "101 - Puxada",
            "Analista", "TI", "Revalle Juazeiro", agora.strftime("%d/%m/%Y"), "teste",
        )
        assert ok, msg
        ok, msg = inv_manager.confirm_loan(item_emprestimo["id"], "teste", "termos_assinados/fake.pdf")
        assert ok, msg
        # Contorna a dependência de .docx só para poder alcançar 'Pendente Devolução'
        # (ver conftest.forcar_devolucao_iniciada) e caracterizar data_devolucao.
        forcar_devolucao_iniciada(item_emprestimo["id"])
        ok, msg = inv_manager.confirm_return(item_emprestimo["id"], "teste", "termos_devolucao_assinados/fake.pdf")
        assert ok, msg

        # Bloco "Periféricos": vínculo de um periférico a um equipamento.
        periferico = criar_periferico()
        ok, msg = inv_manager.link_peripheral_to_equipment(item_emprestimo["id"], periferico["id"], "teste")
        assert ok, msg

        # Bloco "Exclusão": um segundo item, removido do estoque.
        item_excluido = criar_item()
        ok, msg = inv_manager.remove(item_excluido["id"], "teste", "Obsolescência")
        assert ok, msg

        resp = client_gestor.get("/api/reports/monthly", params={"year": agora.year, "month": agora.month})
        assert resp.status_code == 200
        linhas = resp.json()

        # Bloco "Cadastro": os 2 itens acima geram, cada um, sua própria linha de Cadastro.
        tipos_de_operacao = {linha["operation_type"] for linha in linhas}
        assert {"Empréstimo", "Cadastro", "Vínculo Periférico", "Exclusão"} <= tipos_de_operacao

        # -- Linha do bloco Empréstimo: confere o preenchimento das datas correlacionadas.
        linhas_emprestimo = [
            l for l in linhas if l["operation_type"] == "Empréstimo" and l["item_id"] == item_emprestimo["id"]
        ]
        assert len(linhas_emprestimo) == 1
        linha = linhas_emprestimo[0]
        assert linha["usuario"] == "Ciclano"
        assert linha["cpf"] == "22233344455"
        assert linha["data_emprestimo"] is not None
        assert linha["data_confirmacao"] is not None, (
            "data_confirmacao deveria vir preenchida pela subquery correlacionada "
            "(MIN de 'Confirmação Empréstimo' com id maior para o mesmo item_id)."
        )
        assert linha["data_devolucao"] is not None, (
            "data_devolucao deveria vir preenchida pela subquery correlacionada "
            "(MIN de 'Devolução' com id maior para o mesmo item_id)."
        )

        # -- Linha do bloco Cadastro para o mesmo item: não tem essas duas colunas
        # preenchidas (elas só existem no bloco Empréstimo).
        linhas_cadastro = [
            l for l in linhas if l["operation_type"] == "Cadastro" and l["item_id"] == item_emprestimo["id"]
        ]
        assert len(linhas_cadastro) == 1
        assert linhas_cadastro[0]["data_confirmacao"] is None
        assert linhas_cadastro[0]["data_devolucao"] is None
        assert linhas_cadastro[0]["usuario"] is None  # coluna só faz sentido no bloco Empréstimo

        # -- Linha do bloco Periféricos.
        linhas_periferico = [
            l for l in linhas
            if l["operation_type"] == "Vínculo Periférico" and l["peripheral_id"] == periferico["id"]
        ]
        assert len(linhas_periferico) == 1
        assert linhas_periferico[0]["tipo"] == periferico["tipo"]

        # -- Linha do bloco Exclusão.
        linhas_exclusao = [
            l for l in linhas if l["operation_type"] == "Exclusão" and l["item_id"] == item_excluido["id"]
        ]
        assert len(linhas_exclusao) == 1
        assert linhas_exclusao[0]["details"] == "Obsolescência"

    def test_relatorio_de_mes_sem_dados_retorna_lista_vazia(self, client_gestor):
        resp = client_gestor.get("/api/reports/monthly", params={"year": 1999, "month": 1})
        assert resp.status_code == 200
        assert resp.json() == []

    def test_mes_invalido_retorna_400(self, client_gestor):
        resp = client_gestor.get("/api/reports/monthly", params={"year": 2026, "month": 13})
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Mês inválido."

    def test_export_csv_gestor_tem_content_type_e_cabecalho_csv(self, client_gestor, item_disponivel):
        agora = datetime.now()
        resp = client_gestor.get(
            "/api/reports/monthly/export", params={"year": agora.year, "month": agora.month}
        )
        assert resp.status_code == 200
        assert resp.headers["content-type"].startswith("text/csv")
        assert f"relatorio_{agora.year}_{agora.month:02d}.csv" in resp.headers["content-disposition"]
        primeira_linha = resp.text.splitlines()[0]
        assert "operation_type" in primeira_linha
        assert "history_id" in primeira_linha


class TestGraficos:
    def test_chart_loans_tem_um_valor_por_dia_do_mes(self, client_gestor, inv_manager, item_disponivel):
        agora = datetime.now()
        ok, msg = inv_manager.issue(
            item_disponivel["id"], "Fulano", "11122233344", "101 - Puxada", "Analista",
            "TI", "Revalle Juazeiro", agora.strftime("%d/%m/%Y"), "teste",
        )
        assert ok, msg

        resp = client_gestor.get(
            "/api/reports/charts/loans", params={"year": agora.year, "month": agora.month}
        )
        assert resp.status_code == 200
        body = resp.json()
        assert len(body["days"]) == len(body["values"]) == len(body["values2"])
        indice_hoje = agora.day - 1
        assert body["values"][indice_hoje] >= 1

    def test_chart_registrations_tem_um_valor_por_dia_do_mes(self, client_gestor, item_disponivel):
        agora = datetime.now()
        resp = client_gestor.get(
            "/api/reports/charts/registrations", params={"year": agora.year, "month": agora.month}
        )
        assert resp.status_code == 200
        body = resp.json()
        assert len(body["days"]) == len(body["values"])
        assert body["values2"] is None
        indice_hoje = agora.day - 1
        assert body["values"][indice_hoje] >= 1
