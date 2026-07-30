"""
Caracterização da máquina de estados do empréstimo:

    Disponível --(issue)--> Pendente --(confirm_loan + PDF)--> Indisponível
        --(initiate return)--> Pendente Devolução --(confirm_return + PDF)--> Disponível

Endpoints: POST /api/loans, POST /api/loans/{id}/confirm,
POST /api/loans/{id}/return/initiate, POST /api/loans/{id}/return/confirm.
"""
from datetime import datetime, timedelta

import pytest

pytestmark = pytest.mark.integration

ARQUIVO_PDF = {"signed_pdf": ("termo.pdf", b"conteudo-fake-pdf", "application/pdf")}

# MUDANÇA INTENCIONAL (T7, validação de domínio migrou para o schema): este corpo
# usava o CPF "11122233344", que NUNCA foi um CPF de fato válido — os dígitos
# verificadores não conferem (11144477735 é o exemplo clássico de CPF válido para
# testes: 111.444.777-35). Antes da migração, `issue()`/`InventoryDBManager` não
# validavam CPF, então o valor "furado" passava despercebido. Agora `LoanRequest`
# (Pydantic, via `validar_cpf()`) valida o corpo inteiro ANTES do router, e todo
# POST /api/loans com um CPF inválido leva 422 mesmo em cenários que deveriam
# testar outra coisa (item pendente, item indisponível, data no futuro etc.) —
# por isso a troca para um CPF válido aqui, no helper compartilhado.
CPF_VALIDO_TESTE = "11144477735"


def _corpo_emprestimo(item_id: int, data_issue: str = None) -> dict:
    return {
        "item_id": item_id,
        "usuario": "Fulano de Tal",
        "cpf": CPF_VALIDO_TESTE,
        "center_cost": "101 - Puxada",
        "cargo": "Analista",
        "setor": "TI",
        "revenda": "Revalle Juazeiro",
        "date_issue": data_issue or datetime.now().strftime("%d/%m/%Y"),
    }


class TestCaminhoFeliz:
    def test_fluxo_completo_disponivel_ate_devolvido(
        self, client_gestor, item_disponivel, termo_devolucao_disponivel
    ):
        item_id = item_disponivel["id"]

        resp = client_gestor.post("/api/loans", json=_corpo_emprestimo(item_id))
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Empréstimo do item {item_id} para Fulano de Tal iniciado. Status: Pendente."
        assert client_gestor.get(f"/api/items/{item_id}").json()["status"] == "Pendente"

        resp = client_gestor.post(f"/api/loans/{item_id}/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Empréstimo do item {item_id} confirmado."
        assert client_gestor.get(f"/api/items/{item_id}").json()["status"] == "Indisponível"

        resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate")
        if not termo_devolucao_disponivel:
            pytest.skip(
                "Modelo de termo de devolução (.docx) não presente em backend/modelos/devolucao/ "
                "neste ambiente; não é possível caracterizar a geração bem-sucedida do termo nem "
                "o restante do fluxo de devolução por este caminho. Ver docs/TESTES.md "
                "(este mesmo teste, quando rodado com os modelos presentes, cobre o fluxo inteiro)."
            )
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Termo de devolução gerado."
        assert "download_url" in resp.json()
        assert client_gestor.get(f"/api/items/{item_id}").json()["status"] == "Pendente Devolução"

        resp = client_gestor.post(f"/api/loans/{item_id}/return/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Devolução do item {item_id} confirmada."
        item_final = client_gestor.get(f"/api/items/{item_id}").json()
        assert item_final["status"] == "Disponível"
        assert item_final["assigned_to"] is None
        assert item_final["cpf"] is None


class TestIniciarEmprestimoTransicoesInvalidas:
    def test_emprestar_item_ja_pendente(self, client_gestor, item_pendente):
        resp = client_gestor.post("/api/loans", json=_corpo_emprestimo(item_pendente["id"]))
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Este item não está disponível para empréstimo."

    def test_emprestar_item_indisponivel(self, client_gestor, item_indisponivel):
        resp = client_gestor.post("/api/loans", json=_corpo_emprestimo(item_indisponivel["id"]))
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Este item não está disponível para empréstimo."

    def test_emprestar_item_inexistente(self, client_gestor):
        resp = client_gestor.post("/api/loans", json=_corpo_emprestimo(999999))
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Item não encontrado."

    def test_data_de_emprestimo_em_formato_invalido(self, client_gestor, item_disponivel):
        """
        MUDANÇA INTENCIONAL (T7): a validação do FORMATO dd/mm/aaaa de `date_issue`
        saiu de `InventoryDBManager.issue()` (400, mensagem própria) e foi para
        `LoanRequest` (Pydantic, via `validar_data_br`). Agora é 422, com o handler
        global reduzindo a lista de erros do Pydantic a uma frase única citando o
        rótulo amigável do campo ("Data do empréstimo").
        """
        resp = client_gestor.post(
            "/api/loans", json=_corpo_emprestimo(item_disponivel["id"], data_issue="2026-01-01")
        )
        assert resp.status_code == 422
        body = resp.json()
        assert body["detail"] == "Data do empréstimo: Formato de data de empréstimo inválido. Use dd/mm/aaaa."
        assert isinstance(body["errors"], list) and body["errors"]

    def test_data_de_emprestimo_no_futuro(self, client_gestor, item_disponivel):
        amanha = (datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y")
        resp = client_gestor.post(
            "/api/loans", json=_corpo_emprestimo(item_disponivel["id"], data_issue=amanha)
        )
        assert resp.status_code == 400
        assert resp.json()["detail"] == "A data de empréstimo não pode ser no futuro."

    def test_data_de_emprestimo_anterior_ao_cadastro(self, client_gestor, criar_item):
        """
        BUG CONHECIDO (ver test_unit_formatacao.py::TestFormatDate::
        test_string_iso_com_hora_nao_e_reformatada): a mensagem de erro deveria mostrar a
        data de cadastro como dd/mm/aaaa (é isso que format_date() faz), mas
        issue() chama format_date(str(item['date_registered'])) — o str() já convertido
        antes faz o parse interno de format_date falhar (o padrão "%Y-%m-%d" não bate
        com o "YYYY-MM-DD HH:MM:SS" resultante), então a função devolve a string crua do
        MySQL sem reformatar. O usuário vê "2026-06-01 09:00:00" em vez de "01/06/2026".
        """
        item = criar_item(date_registered=datetime(2026, 6, 1, 9, 0, 0))
        resp = client_gestor.post(
            "/api/loans", json=_corpo_emprestimo(item["id"], data_issue="01/05/2026")
        )
        assert resp.status_code == 400
        assert resp.json()["detail"] == (
            "Data de empréstimo não pode ser anterior ao cadastro (2026-06-01 09:00:00)."
        )


class TestConfirmarEmprestimoTransicoesInvalidas:
    MSG_APENAS_PENDENTE = "Apenas itens com status 'Pendente' podem ser confirmados."

    def test_confirmar_item_disponivel(self, client_gestor, item_disponivel):
        resp = client_gestor.post(f"/api/loans/{item_disponivel['id']}/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 400
        assert resp.json()["detail"] == self.MSG_APENAS_PENDENTE

    def test_confirmar_item_ja_indisponivel(self, client_gestor, item_indisponivel):
        resp = client_gestor.post(f"/api/loans/{item_indisponivel['id']}/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 400
        assert resp.json()["detail"] == self.MSG_APENAS_PENDENTE

    def test_confirmar_item_inexistente(self, client_gestor):
        resp = client_gestor.post("/api/loans/999999/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 400
        assert resp.json()["detail"] == self.MSG_APENAS_PENDENTE


class TestIniciarDevolucaoTransicoesInvalidas:
    MSG_APENAS_INDISPONIVEL = "Apenas itens com status 'Indisponível' podem ter termo de devolução gerado."

    def test_iniciar_devolucao_de_item_disponivel(self, client_gestor, item_disponivel):
        resp = client_gestor.post(f"/api/loans/{item_disponivel['id']}/return/initiate")
        assert resp.status_code == 400
        assert resp.json()["detail"] == self.MSG_APENAS_INDISPONIVEL

    def test_iniciar_devolucao_de_item_pendente(self, client_gestor, item_pendente):
        resp = client_gestor.post(f"/api/loans/{item_pendente['id']}/return/initiate")
        assert resp.status_code == 400
        assert resp.json()["detail"] == self.MSG_APENAS_INDISPONIVEL

    def test_iniciar_devolucao_de_item_inexistente(self, client_gestor):
        resp = client_gestor.post("/api/loans/999999/return/initiate")
        assert resp.status_code == 400
        assert resp.json()["detail"] == self.MSG_APENAS_INDISPONIVEL

    def test_iniciar_devolucao_com_arquivo_de_modelo_ausente_retorna_mensagem_amigavel(
        self, client_gestor, item_indisponivel, monkeypatch
    ):
        """
        CORRIGIDO NESTE CICLO (era o item nº 3 de "Mudanças esperadas", ver
        docs/TESTES.md): antes, para uma revenda VÁLIDA (presente em
        TERMO_DEVOLUCAO_MODELOS), `generate_return_term_bytes()` só caía no ramo
        amigável ("Modelo de termo de devolução não encontrado para {revenda}.")
        quando a revenda era desconhecida — se o .docx de uma revenda válida
        estivesse ausente/corrompido, a função pulava direto para `Document(modelo_path)`
        e devolvia a exceção técnica crua ("Erro ao gerar documento: {e}"). O código
        atual (`app/db/inventory_manager_db.py::generate_return_term_bytes`) chama
        `os.path.exists(modelo_path)` incondicionalmente, então o ramo amigável agora
        dispara também para revenda válida com arquivo ausente.

        Este teste não depende mais de o ambiente ter ou não os .docx reais (o
        worktree deste módulo já os tem, ao contrário do Módulo 7a que escreveu a
        suíte original) — usa `monkeypatch` para forçar, de forma determinística, o
        caminho do modelo de uma revenda válida e conhecida para um arquivo que não
        existe, e confere que a mensagem amigável aparece (não mais a crua).
        """
        from app.db import inventory_manager_db

        monkeypatch.setitem(
            inventory_manager_db.TERMO_DEVOLUCAO_MODELOS,
            "Revalle Juazeiro",
            "/caminho/inexistente/termo_devolucao_juazeiro.docx",
        )
        resp = client_gestor.post(f"/api/loans/{item_indisponivel['id']}/return/initiate")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Modelo de termo de devolução não encontrado para Revalle Juazeiro."
        # E, importante: a transação não é aplicada -- o item continua Indisponível.
        assert client_gestor.get(f"/api/items/{item_indisponivel['id']}").json()["status"] == "Indisponível"

    def test_iniciar_devolucao_com_revenda_desconhecida_retorna_mensagem_amigavel(
        self, client_gestor, inv_manager, item_indisponivel
    ):
        """Caracterização determinística (não depende de arquivos em disco): uma revenda
        que não existe em TERMO_DEVOLUCAO_MODELOS sempre cai no ramo amigável."""
        inv_manager.update_item(item_indisponivel["id"], {"revenda": "Revenda Sem Modelo Cadastrado"}, "teste")
        resp = client_gestor.post(f"/api/loans/{item_indisponivel['id']}/return/initiate")
        assert resp.status_code == 400
        assert resp.json()["detail"] == (
            "Modelo de termo de devolução não encontrado para Revenda Sem Modelo Cadastrado."
        )


class TestConfirmarDevolucaoTransicoesInvalidas:
    MSG_APENAS_PENDENTE_DEVOLUCAO = "Apenas itens com status 'Pendente Devolução' podem ser confirmados."

    def test_confirmar_devolucao_de_item_indisponivel(self, client_gestor, item_indisponivel):
        resp = client_gestor.post(f"/api/loans/{item_indisponivel['id']}/return/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 400
        assert resp.json()["detail"] == self.MSG_APENAS_PENDENTE_DEVOLUCAO

    def test_confirmar_devolucao_de_item_disponivel(self, client_gestor, item_disponivel):
        resp = client_gestor.post(f"/api/loans/{item_disponivel['id']}/return/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 400
        assert resp.json()["detail"] == self.MSG_APENAS_PENDENTE_DEVOLUCAO

    def test_confirmar_devolucao_de_item_em_pendente_devolucao_tem_sucesso(
        self, client_gestor, item_indisponivel, forcar_devolucao_iniciada
    ):
        """Usa o utilitário de teste forcar_devolucao_iniciada (ver conftest.py) para
        alcançar 'Pendente Devolução' sem depender dos arquivos .docx, e assim
        caracterizar confirm_return isoladamente (que em si não usa templates)."""
        forcar_devolucao_iniciada(item_indisponivel["id"])
        resp = client_gestor.post(f"/api/loans/{item_indisponivel['id']}/return/confirm", files=ARQUIVO_PDF)
        assert resp.status_code == 200
        assert resp.json()["detail"] == f"Devolução do item {item_indisponivel['id']} confirmada."
