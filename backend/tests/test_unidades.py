"""
Testes do Módulo 9 — Backend da funcionalidade de Unidades.

Cobre: CRUD completo (app/routers/unidades.py + app/db/unidade_db.py), RBAC
das 6 rotas, validação de CNPJ/CEP/UF (app/schemas/unidades.py), a
renomeação de unidade cascateando para items/history na mesma transação, a
inativação recusada com itens ativos, os indicadores agregados por unidade
(app/db/unidade_db.py::UnidadeDBManager.get_indicadores) e /api/constants
passando a servir a lista de revendas a partir do banco.

NOTA sobre `devolucoes_concluidas`: InventoryDBManager.confirm_return() e
generate_return_term_bytes() (fora da posse deste módulo) NÃO gravam
history.revenda ao inserir 'Devolução'/'Confirmação Devolução' — só issue()
e confirm_loan() preenchem essa coluna. Por isso get_indicadores() usa JOIN
com items (via item_id) para essas contagens, em vez de history.revenda
direto; os testes abaixo montam o cenário exatamente para expor esse detalhe
caso a implementação regressione para o filtro ingênuo.
"""
from datetime import datetime
from uuid import uuid4

import pytest

pytestmark = pytest.mark.integration

MSG_SEM_TOKEN = "Token de autenticação não fornecido."
MSG_SO_GESTOR = "Acesso restrito ao Gestor."
MSG_GESTOR_OU_TECNICO = "Acesso restrito a Gestor ou Técnico."

DATA_HOJE = datetime.now().strftime("%d/%m/%Y")


def _nome_unico(prefixo: str = "Unidade Teste") -> str:
    return f"{prefixo} {uuid4().hex[:8]}"


def _payload_unidade(**overrides) -> dict:
    dados = {
        "nome": _nome_unico(),
        "razao_social": "REVENDA VALLE DA INTEGRAÇÃO LTDA",
        "cnpj": "04.690.106/0001-15",
        "endereco": "Centro Industrial São Francisco, Quadra QID, Lotes 5/6 s/n",
        "cep": "48905-630",
        "cidade": "Juazeiro",
        "uf": "BA",
    }
    dados.update(overrides)
    return dados


# ────────────────────────────────────────────────────────────────────────────
# Criação
# ────────────────────────────────────────────────────────────────────────────


class TestCriarUnidade:
    def test_criar_unidade_com_sucesso_retorna_200_e_id(self, client_gestor):
        resp = client_gestor.post("/api/unidades", json=_payload_unidade())
        assert resp.status_code == 200
        body = resp.json()
        assert "cadastrada com sucesso" in body["detail"]
        assert isinstance(body["id"], int)

    def test_criar_unidade_normaliza_cnpj_e_cep(self, client_gestor):
        nome = _nome_unico()
        resp = client_gestor.post(
            "/api/unidades",
            json=_payload_unidade(nome=nome, cnpj="04690106000115", cep="48905630"),
        )
        assert resp.status_code == 200
        criada = client_gestor.get(f"/api/unidades/{resp.json()['id']}").json()
        assert criada["cnpj"] == "04.690.106/0001-15"
        assert criada["cep"] == "48905-630"

    def test_criar_unidade_cnpj_com_digito_verificador_errado_retorna_422(self, client_gestor):
        resp = client_gestor.post(
            "/api/unidades", json=_payload_unidade(cnpj="04.690.106/0001-16")
        )
        assert resp.status_code == 422
        assert isinstance(resp.json()["detail"], str)

    def test_criar_unidade_cnpj_com_todos_digitos_iguais_retorna_422(self, client_gestor):
        resp = client_gestor.post(
            "/api/unidades", json=_payload_unidade(cnpj="11.111.111/1111-11")
        )
        assert resp.status_code == 422

    def test_criar_unidade_cnpj_com_quantidade_errada_de_digitos_retorna_422(self, client_gestor):
        resp = client_gestor.post("/api/unidades", json=_payload_unidade(cnpj="123"))
        assert resp.status_code == 422

    def test_criar_unidade_uf_invalida_retorna_422(self, client_gestor):
        resp = client_gestor.post("/api/unidades", json=_payload_unidade(uf="XX"))
        assert resp.status_code == 422

    def test_criar_unidade_cep_com_quantidade_errada_de_digitos_retorna_422(self, client_gestor):
        resp = client_gestor.post("/api/unidades", json=_payload_unidade(cep="123"))
        assert resp.status_code == 422

    def test_criar_unidade_sem_cep_e_aceito(self, client_gestor):
        """Caso real do seed: Revalle Petrolina não tem CEP na origem."""
        payload = _payload_unidade()
        del payload["cep"]
        resp = client_gestor.post("/api/unidades", json=payload)
        assert resp.status_code == 200

    def test_criar_unidade_nome_duplicado_retorna_400(self, client_gestor, unidade):
        resp = client_gestor.post("/api/unidades", json=_payload_unidade(nome=unidade["nome"]))
        assert resp.status_code == 400
        assert "já existe" in resp.json()["detail"].lower()

    def test_criar_unidade_sem_nome_retorna_422(self, client_gestor):
        payload = _payload_unidade()
        del payload["nome"]
        resp = client_gestor.post("/api/unidades", json=payload)
        assert resp.status_code == 422

    def test_rbac_post_unidades_somente_gestor(
        self, client, client_gestor, client_tecnico, client_aprendiz
    ):
        resp_anonimo = client.post("/api/unidades", json=_payload_unidade())
        assert resp_anonimo.status_code == 401
        assert resp_anonimo.json()["detail"] == MSG_SEM_TOKEN

        resp_aprendiz = client_aprendiz.post("/api/unidades", json=_payload_unidade())
        assert resp_aprendiz.status_code == 403

        resp_tecnico = client_tecnico.post("/api/unidades", json=_payload_unidade())
        assert resp_tecnico.status_code == 403
        assert resp_tecnico.json()["detail"] == MSG_SO_GESTOR

        resp_gestor = client_gestor.post("/api/unidades", json=_payload_unidade())
        assert resp_gestor.status_code == 200


# ────────────────────────────────────────────────────────────────────────────
# Listagem e busca
# ────────────────────────────────────────────────────────────────────────────


class TestListarEBuscarUnidade:
    def test_listar_unidades_qualquer_role_autenticado_200_sem_token_401(
        self, client, client_gestor, client_tecnico, client_aprendiz, unidade
    ):
        assert client.get("/api/unidades").status_code == 401
        for c in (client_gestor, client_tecnico, client_aprendiz):
            resp = c.get("/api/unidades")
            assert resp.status_code == 200
            nomes = [u["nome"] for u in resp.json()]
            assert unidade["nome"] in nomes

    def test_listar_unidades_omite_inativa_por_padrao_e_inclui_com_flag(
        self, client_gestor, unidade_db, unidade
    ):
        ok, msg = unidade_db.deactivate_unidade(unidade["id"])
        assert ok, msg

        resp = client_gestor.get("/api/unidades")
        assert unidade["nome"] not in [u["nome"] for u in resp.json()]

        resp_com_inativas = client_gestor.get("/api/unidades?include_inactive=true")
        assert unidade["nome"] in [u["nome"] for u in resp_com_inativas.json()]

    def test_buscar_unidade_por_id(self, client_tecnico, unidade):
        resp = client_tecnico.get(f"/api/unidades/{unidade['id']}")
        assert resp.status_code == 200
        assert resp.json()["nome"] == unidade["nome"]
        assert resp.json()["cnpj"] == unidade["cnpj"]

    def test_buscar_unidade_inexistente_retorna_404(self, client_gestor):
        resp = client_gestor.get("/api/unidades/999999")
        assert resp.status_code == 404

    def test_get_unidade_por_id_sem_token_401(self, client, unidade):
        resp = client.get(f"/api/unidades/{unidade['id']}")
        assert resp.status_code == 401


# ────────────────────────────────────────────────────────────────────────────
# Atualização (com o caso crítico: renomear cascateia para items/history)
# ────────────────────────────────────────────────────────────────────────────


class TestAtualizarUnidade:
    def test_atualizar_campos_simples(self, client_gestor, unidade):
        resp = client_gestor.put(
            f"/api/unidades/{unidade['id']}", json={"cidade": "Nova Cidade", "uf": "SP"}
        )
        assert resp.status_code == 200
        atualizado = client_gestor.get(f"/api/unidades/{unidade['id']}").json()
        assert atualizado["cidade"] == "Nova Cidade"
        assert atualizado["uf"] == "SP"
        # Nome não foi tocado
        assert atualizado["nome"] == unidade["nome"]

    def test_atualizar_sem_nenhum_campo_retorna_400(self, client_gestor, unidade):
        resp = client_gestor.put(f"/api/unidades/{unidade['id']}", json={})
        assert resp.status_code == 400

    def test_atualizar_para_nome_ja_existente_retorna_400(
        self, client_gestor, unidade, criar_unidade
    ):
        outra = criar_unidade()
        resp = client_gestor.put(f"/api/unidades/{unidade['id']}", json={"nome": outra["nome"]})
        assert resp.status_code == 400
        assert "já existe" in resp.json()["detail"].lower()
        # Nome original preservado
        assert client_gestor.get(f"/api/unidades/{unidade['id']}").json()["nome"] == unidade["nome"]

    def test_atualizar_unidade_inexistente_retorna_404(self, client_gestor):
        resp = client_gestor.put("/api/unidades/999999", json={"cidade": "X"})
        assert resp.status_code == 404

    def test_atualizar_cnpj_invalido_retorna_422(self, client_gestor, unidade):
        resp = client_gestor.put(f"/api/unidades/{unidade['id']}", json={"cnpj": "123"})
        assert resp.status_code == 422

    def test_rbac_put_unidades_somente_gestor(
        self, client, client_gestor, client_tecnico, client_aprendiz, unidade
    ):
        assert client.put(f"/api/unidades/{unidade['id']}", json={"cidade": "X"}).status_code == 401
        assert client_aprendiz.put(f"/api/unidades/{unidade['id']}", json={"cidade": "X"}).status_code == 403
        assert client_tecnico.put(f"/api/unidades/{unidade['id']}", json={"cidade": "X"}).status_code == 403
        assert client_gestor.put(f"/api/unidades/{unidade['id']}", json={"cidade": "X"}).status_code == 200

    def test_renomear_unidade_cascateia_para_items_e_history(
        self, client_gestor, unidade, criar_item, inv_manager
    ):
        """O caso crítico do contrato T2: renomear uma unidade precisa
        reapontar `items.revenda` e `history.revenda` na mesma transação,
        senão itens e histórico "perdem" a unidade a que pertenciam."""
        nome_antigo = unidade["nome"]
        item1 = criar_item(revenda=nome_antigo)
        item2 = criar_item(revenda=nome_antigo)

        ok, msg = inv_manager.issue(
            item1["id"], "Fulano de Tal", "11122233344", "101 - Puxada",
            "Analista", "TI", nome_antigo, DATA_HOJE, "teste",
        )
        assert ok, msg

        novo_nome = f"{nome_antigo} Renomeada"
        resp = client_gestor.put(f"/api/unidades/{unidade['id']}", json={"nome": novo_nome})
        assert resp.status_code == 200
        assert "reapontados" in resp.json()["detail"]

        # items seguem a unidade renomeada
        item1_atual = inv_manager.find(item1["id"])
        item2_atual = inv_manager.find(item2["id"])
        assert item1_atual["revenda"] == novo_nome
        assert item2_atual["revenda"] == novo_nome

        # history também segue (ao menos o registro de Empréstimo, que grava revenda)
        historico, _total = inv_manager.list_history()
        registros_emprestimo = [
            h for h in historico if h["operation"] == "Empréstimo" and h["item_id"] == item1["id"]
        ]
        assert registros_emprestimo, "Registro de Empréstimo do item1 não encontrado no histórico."
        assert all(h["revenda"] == novo_nome for h in registros_emprestimo)

        # a unidade em si reflete o novo nome
        atualizada = client_gestor.get(f"/api/unidades/{unidade['id']}").json()
        assert atualizada["nome"] == novo_nome

    def test_atualizar_sem_mudar_nome_nao_cascateia(
        self, client_gestor, unidade, criar_item, inv_manager
    ):
        """Atualizar outros campos sem tocar em `nome` não deve reapontar nada
        — a mensagem de sucesso não deve mencionar reapontamento."""
        item = criar_item(revenda=unidade["nome"])
        resp = client_gestor.put(f"/api/unidades/{unidade['id']}", json={"cidade": "Outra Cidade"})
        assert resp.status_code == 200
        assert "reapontados" not in resp.json()["detail"]
        assert inv_manager.find(item["id"])["revenda"] == unidade["nome"]


# ────────────────────────────────────────────────────────────────────────────
# Inativação (nunca apaga)
# ────────────────────────────────────────────────────────────────────────────


class TestInativarUnidade:
    def test_inativar_sem_itens_ativos_e_permitido(self, client_gestor, unidade):
        resp = client_gestor.delete(f"/api/unidades/{unidade['id']}")
        assert resp.status_code == 200
        assert "inativada" in resp.json()["detail"]

        atual = client_gestor.get(f"/api/unidades/{unidade['id']}").json()
        assert atual["is_active"] is False

    def test_inativar_com_itens_ativos_e_recusada(self, client_gestor, unidade, criar_item):
        criar_item(revenda=unidade["nome"])
        resp = client_gestor.delete(f"/api/unidades/{unidade['id']}")
        assert resp.status_code == 400
        assert "item" in resp.json()["detail"].lower()

        # Continua ativa — a recusa realmente impediu a mudança
        atual = client_gestor.get(f"/api/unidades/{unidade['id']}").json()
        assert atual["is_active"] is True

    def test_inativar_unidade_inexistente_retorna_404(self, client_gestor):
        resp = client_gestor.delete("/api/unidades/999999")
        assert resp.status_code == 404

    def test_rbac_delete_unidades_somente_gestor(
        self, client, client_gestor, client_tecnico, client_aprendiz, criar_unidade
    ):
        u1, u2, u3, u4 = (criar_unidade() for _ in range(4))
        assert client.delete(f"/api/unidades/{u1['id']}").status_code == 401
        assert client_aprendiz.delete(f"/api/unidades/{u2['id']}").status_code == 403
        assert client_tecnico.delete(f"/api/unidades/{u3['id']}").status_code == 403
        assert client_gestor.delete(f"/api/unidades/{u4['id']}").status_code == 200


# ────────────────────────────────────────────────────────────────────────────
# Reativação (adição pedida pelo coordenador: sem isso, uma inativação por
# engano não tinha caminho de volta pela API)
# ────────────────────────────────────────────────────────────────────────────


class TestReativarUnidade:
    def test_reativar_unidade_volta_a_listar_e_a_aparecer_em_constants(
        self, client_gestor, unidade, unidade_db
    ):
        ok, msg = unidade_db.deactivate_unidade(unidade["id"])
        assert ok, msg
        assert unidade["nome"] not in [u["nome"] for u in client_gestor.get("/api/unidades").json()]
        assert unidade["nome"] not in client_gestor.get("/api/constants").json()["revendas"]

        resp = client_gestor.post(f"/api/unidades/{unidade['id']}/reativar")
        assert resp.status_code == 200
        assert "reativada" in resp.json()["detail"]

        atual = client_gestor.get(f"/api/unidades/{unidade['id']}").json()
        assert atual["is_active"] is True
        assert unidade["nome"] in [u["nome"] for u in client_gestor.get("/api/unidades").json()]
        assert unidade["nome"] in client_gestor.get("/api/constants").json()["revendas"]

    def test_reativar_unidade_ja_ativa_e_recusada(self, client_gestor, unidade):
        """DECISÃO DE PRODUTO (ver docstring de reactivate_unidade): recusa em
        vez de tratar como sucesso idempotente, por simetria com
        deactivate_unidade (que também recusa inativar o que já está
        inativo)."""
        resp = client_gestor.post(f"/api/unidades/{unidade['id']}/reativar")
        assert resp.status_code == 400
        assert "já está ativa" in resp.json()["detail"].lower()

    def test_reativar_unidade_inexistente_retorna_404(self, client_gestor):
        resp = client_gestor.post("/api/unidades/999999/reativar")
        assert resp.status_code == 404

    def test_rbac_post_reativar_somente_gestor(
        self, client, client_gestor, client_tecnico, client_aprendiz, unidade_db, criar_unidade
    ):
        u1, u2, u3, u4 = (criar_unidade() for _ in range(4))
        for u in (u1, u2, u3, u4):
            ok, msg = unidade_db.deactivate_unidade(u["id"])
            assert ok, msg

        assert client.post(f"/api/unidades/{u1['id']}/reativar").status_code == 401
        resp_aprendiz = client_aprendiz.post(f"/api/unidades/{u2['id']}/reativar")
        assert resp_aprendiz.status_code == 403
        resp_tecnico = client_tecnico.post(f"/api/unidades/{u3['id']}/reativar")
        assert resp_tecnico.status_code == 403
        assert client_gestor.post(f"/api/unidades/{u4['id']}/reativar").status_code == 200


# ────────────────────────────────────────────────────────────────────────────
# Indicadores
# ────────────────────────────────────────────────────────────────────────────


class TestIndicadores:
    def test_indicadores_com_cenario_montado_de_proposito(
        self, client_gestor, inv_manager, criar_item, criar_periferico,
        unidade, forcar_devolucao_iniciada,
    ):
        nome = unidade["nome"]

        # item1: emprestado + confirmado (Indisponível), com periférico vinculado
        item1 = criar_item(revenda=nome)
        periferico1 = criar_periferico()
        ok, msg = inv_manager.link_peripheral_to_equipment(item1["id"], periferico1["id"], "teste")
        assert ok, msg
        ok, msg = inv_manager.issue(
            item1["id"], "Fulano", "11122233344", "101 - Puxada", "Analista", "TI",
            nome, DATA_HOJE, "teste",
        )
        assert ok, msg
        ok, msg = inv_manager.confirm_loan(item1["id"], "teste", "termos/assinado1.pdf")
        assert ok, msg

        # item2: emprestado + confirmado + devolução confirmada -> volta a Disponível
        item2 = criar_item(revenda=nome)
        ok, msg = inv_manager.issue(
            item2["id"], "Beltrano", "22233344455", "101 - Puxada", "Analista", "TI",
            nome, DATA_HOJE, "teste",
        )
        assert ok, msg
        ok, msg = inv_manager.confirm_loan(item2["id"], "teste", "termos/assinado2.pdf")
        assert ok, msg
        forcar_devolucao_iniciada(item2["id"], "teste")
        ok, msg = inv_manager.confirm_return(item2["id"], "teste", "termos/devolucao2.pdf")
        assert ok, msg

        # item3: apenas emprestado, ainda não confirmado (Pendente)
        item3 = criar_item(revenda=nome)
        ok, msg = inv_manager.issue(
            item3["id"], "Ciclano", "33344455566", "101 - Puxada", "Analista", "TI",
            nome, DATA_HOJE, "teste",
        )
        assert ok, msg

        # item4: recém-cadastrado, nunca emprestado (Disponível)
        criar_item(revenda=nome)

        # item5: emprestado + confirmado + devolução iniciada mas NÃO confirmada
        # (Pendente Devolução)
        item5 = criar_item(revenda=nome)
        ok, msg = inv_manager.issue(
            item5["id"], "Sicrano", "44455566677", "101 - Puxada", "Analista", "TI",
            nome, DATA_HOJE, "teste",
        )
        assert ok, msg
        ok, msg = inv_manager.confirm_loan(item5["id"], "teste", "termos/assinado5.pdf")
        assert ok, msg
        forcar_devolucao_iniciada(item5["id"], "teste")

        resp = client_gestor.get(f"/api/unidades/{unidade['id']}/indicadores")
        assert resp.status_code == 200
        body = resp.json()

        assert body["termos_emitidos"] == 4  # item1, item2, item3, item5
        assert body["termos_confirmados"] == 3  # item1, item2, item5 (item3 não confirmado)
        assert body["devolucoes_concluidas"] == 1  # apenas item2
        assert body["itens_total"] == 5  # todos os 5 itens ativos
        assert body["itens_por_status"] == {
            "Disponível": 2,  # item2 (devolvido) + item4
            "Indisponível": 1,  # item1
            "Pendente": 1,  # item3
            "Pendente Devolução": 1,  # item5
        }
        assert body["emprestimos_ativos"] == 1  # == Indisponível
        assert body["perifericos_vinculados"] == 1  # apenas item1

    def test_indicadores_unidade_sem_movimento_retorna_tudo_zerado(self, client_gestor, unidade):
        resp = client_gestor.get(f"/api/unidades/{unidade['id']}/indicadores")
        assert resp.status_code == 200
        body = resp.json()
        assert body["termos_emitidos"] == 0
        assert body["termos_confirmados"] == 0
        assert body["devolucoes_concluidas"] == 0
        assert body["itens_total"] == 0
        assert body["itens_por_status"] == {
            "Disponível": 0, "Indisponível": 0, "Pendente": 0, "Pendente Devolução": 0,
        }
        assert body["emprestimos_ativos"] == 0
        assert body["perifericos_vinculados"] == 0

    def test_indicadores_unidade_inexistente_retorna_404(self, client_gestor):
        resp = client_gestor.get("/api/unidades/999999/indicadores")
        assert resp.status_code == 404

    def test_rbac_get_indicadores_gestor_e_tecnico_200_aprendiz_403_anonimo_401(
        self, client, client_gestor, client_tecnico, client_aprendiz, unidade
    ):
        rota = f"/api/unidades/{unidade['id']}/indicadores"
        assert client.get(rota).status_code == 401
        assert client_gestor.get(rota).status_code == 200
        assert client_tecnico.get(rota).status_code == 200
        resp_aprendiz = client_aprendiz.get(rota)
        assert resp_aprendiz.status_code == 403
        assert resp_aprendiz.json()["detail"] == MSG_GESTOR_OU_TECNICO


# ────────────────────────────────────────────────────────────────────────────
# /api/constants passa a servir revendas do banco
# ────────────────────────────────────────────────────────────────────────────


class TestConstantsRefleteUnidades:
    def test_revendas_reflete_unidades_ativas_ordenadas(self, client_gestor, criar_unidade):
        criar_unidade(nome="Zeta Unidade de Teste")
        criar_unidade(nome="Alfa Unidade de Teste")

        resp = client_gestor.get("/api/constants")
        assert resp.status_code == 200
        revendas = resp.json()["revendas"]
        assert "Zeta Unidade de Teste" in revendas
        assert "Alfa Unidade de Teste" in revendas
        assert revendas.index("Alfa Unidade de Teste") < revendas.index("Zeta Unidade de Teste")

    def test_revendas_nao_inclui_unidade_inativa(self, client_gestor, unidade, unidade_db):
        ok, msg = unidade_db.deactivate_unidade(unidade["id"])
        assert ok, msg
        resp = client_gestor.get("/api/constants")
        assert unidade["nome"] not in resp.json()["revendas"]

    def test_revendas_com_banco_vazio_retorna_lista_vazia(self, client_gestor):
        resp = client_gestor.get("/api/constants")
        assert resp.status_code == 200
        assert resp.json()["revendas"] == []
