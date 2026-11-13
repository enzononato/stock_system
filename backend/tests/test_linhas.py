"""Cadastro e consulta de linhas telefônicas (F5)."""
import pytest

from tests.fixtures_linhas import *  # noqa: F401,F403

pytestmark = pytest.mark.integration


def test_criar_linha_normaliza_numero_e_grava_ativacao(
    client_gestor, usuario_gestor, unidade, dados_linha, eventos_no_banco
):
    resp = client_gestor.post("/api/linhas", json=dados_linha(numero="+55 (74) 99999-0001"))

    assert resp.status_code == 201, resp.text
    linha = resp.json()
    assert linha["numero"] == "74999990001"
    assert linha["status"] == "ativa"
    assert linha["custo_mensal"] == "49.90"
    assert linha["unidade_nome"] == unidade["nome"]
    eventos = eventos_no_banco(linha["id"])
    assert [e["tipo"] for e in eventos] == ["ativacao"]
    assert eventos[0]["por_usuario"] == usuario_gestor["username"]


def test_criar_grava_auditoria(nova_linha, auditoria):
    linha = nova_linha(numero="(74) 99999-0001")

    registros = auditoria("linha", "criar")
    assert len(registros) == 1
    assert registros[0]["entidade_id"] == str(linha["id"])
    assert registros[0]["depois"]["numero"] == "74999990001"


def test_numero_duplicado_recusado(client_gestor, nova_linha, dados_linha):
    nova_linha(numero="74999990001")

    resp = client_gestor.post("/api/linhas", json=dados_linha(numero="(74) 99999-0001"))

    assert resp.status_code == 409
    assert resp.json() == {"detail": "Já existe uma linha com este número."}


def test_numero_de_linha_cancelada_sugere_reativar(client_gestor, nova_linha, dados_linha, executar_sql):
    linha = nova_linha(numero="74999990001")
    executar_sql("UPDATE linhas SET status='cancelada' WHERE id=%s", (linha["id"],))

    resp = client_gestor.post("/api/linhas", json=dados_linha(numero="74999990001"))

    assert resp.status_code == 409
    assert resp.json()["detail"] == (
        "Já existe uma linha cancelada com este número: reative-a em vez de cadastrar de novo."
    )


def test_numero_invalido_422(client_gestor, dados_linha):
    resp = client_gestor.post("/api/linhas", json=dados_linha(numero="7499999"))
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Número: informe DDD + 8 ou 9 dígitos."

    resp = client_gestor.post("/api/linhas", json=dados_linha(numero="20999990001"))
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Número: DDD inexistente."


def test_regras_do_aparelho(client_gestor, dados_linha, criar_aparelho, criar_item, executar_sql):
    def tentar(**extra):
        return client_gestor.post("/api/linhas", json=dados_linha(**extra))

    notebook = criar_item(tipo="Notebook")
    resp = tentar(item_id=notebook["id"])
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Só Celular ou Tablet podem receber uma linha."

    baixado = criar_aparelho()
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (baixado["id"],))
    resp = tentar(item_id=baixado["id"])
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Aparelho baixado não pode receber linha."

    resp = tentar(item_id=999999)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Aparelho não encontrado."

    tablet = criar_aparelho(tipo="Tablet", brand="Apple", model="iPad")
    resp = tentar(item_id=tablet["id"])
    assert resp.status_code == 201
    assert resp.json()["item_tipo"] == "Tablet"

    celular = criar_aparelho()
    assert tentar(numero="74999990001", item_id=celular["id"]).status_code == 201
    resp = tentar(numero="74999990002", item_id=celular["id"])
    assert resp.status_code == 409
    assert resp.json()["detail"] == "O aparelho já está vinculado à linha (74) 99999-0001."


def test_regras_do_colaborador_e_unidade(
    client_gestor, dados_linha, criar_colaborador, executar_sql, criar_unidade
):
    def tentar(**extra):
        return client_gestor.post("/api/linhas", json=dados_linha(**extra))

    desligado = criar_colaborador("João Lima", situacao="desligado")
    resp = tentar(colaborador_id=desligado)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Colaborador desligado não pode receber linha."

    afastado = criar_colaborador("Ana Lima", situacao="afastado")
    assert tentar(colaborador_id=afastado).status_code == 201

    resp = tentar(colaborador_id=999999)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Colaborador não encontrado."

    inativa = criar_unidade()
    executar_sql("UPDATE unidades SET is_active=0 WHERE id=%s", (inativa["id"],))
    resp = tentar(unidade_id=inativa["id"])
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Unidade não encontrada ou inativa."


def test_aprendiz_le_mas_nao_escreve(client_aprendiz, client_tecnico, nova_linha, dados_linha):
    linha = nova_linha()

    assert client_aprendiz.get("/api/linhas").status_code == 200
    assert client_aprendiz.get(f"/api/linhas/{linha['id']}").status_code == 200
    resp = client_aprendiz.post("/api/linhas", json=dados_linha())
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
    assert client_aprendiz.get("/api/linhas/aparelhos").status_code == 403

    assert client_tecnico.post("/api/linhas", json=dados_linha()).status_code == 201


def test_listar_filtros_e_paginacao(
    client_gestor, nova_linha, criar_aparelho, criar_colaborador, executar_sql
):
    celular = criar_aparelho()
    maria = criar_colaborador("Maria Souza")
    l1 = nova_linha(numero="74999990001", item_id=celular["id"], colaborador_id=maria)
    nova_linha(numero="74999990002")
    l3 = nova_linha(numero="74999990003", operadora="Claro")
    executar_sql("UPDATE linhas SET status='suspensa' WHERE id=%s", (l3["id"],))

    def listar(**params):
        resp = client_gestor.get("/api/linhas", params=params)
        assert resp.status_code == 200, resp.text
        return resp.json()

    assert listar(status="suspensa")["total"] == 1
    assert listar(operadora="Claro")["total"] == 1
    assert listar(item_id=celular["id"])["total"] == 1
    assert listar(colaborador_id=maria)["total"] == 1
    assert listar(busca="99999-0002")["total"] == 1
    assert listar(busca="maria")["total"] == 1
    pagina = listar(limit=1)
    assert pagina["total"] == 3
    assert len(pagina["items"]) == 1
    todos = listar()
    assert todos["items"][-1]["id"] == l3["id"]  # a suspensa vem por último
    assert todos["items"][0]["id"] == l1["id"]
    assert listar(limit=1, offset=2)["items"][0]["id"] == l3["id"]
    assert client_gestor.get("/api/linhas", params={"limit": 101}).status_code == 422


def test_linha_continua_legivel_com_aparelho_baixado_e_colaborador_desligado(
    client_gestor, nova_linha, criar_aparelho, criar_colaborador, executar_sql
):
    celular = criar_aparelho()
    colaborador = criar_colaborador()
    linha = nova_linha(item_id=celular["id"], colaborador_id=colaborador)
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (celular["id"],))
    executar_sql("UPDATE colaboradores SET situacao='desligado' WHERE id=%s", (colaborador,))

    resp = client_gestor.get(f"/api/linhas/{linha['id']}")
    assert resp.status_code == 200
    corpo = resp.json()
    assert corpo["item_ativo"] is False
    assert corpo["colaborador_situacao"] == "desligado"
    assert corpo["item_id"] == celular["id"]
    assert client_gestor.get("/api/linhas").json()["total"] == 1

    vazia = nova_linha(plano=None)
    resp = client_gestor.get(f"/api/linhas/{vazia['id']}")
    assert resp.status_code == 200
    corpo = resp.json()
    assert corpo["item_id"] is None
    assert corpo["item_ativo"] is None
    assert corpo["colaborador_id"] is None
    assert corpo["colaborador_situacao"] is None
    assert corpo["plano"] is None

    assert client_gestor.get("/api/linhas/999999").status_code == 404
    assert client_gestor.get("/api/linhas/999999").json()["detail"] == "Linha não encontrada."


def test_aparelhos_so_celular_e_tablet_livres(
    client_gestor, nova_linha, criar_aparelho, criar_item, criar_colaborador, executar_sql
):
    celular_livre = criar_aparelho(model="Livre")
    tablet_livre = criar_aparelho(tipo="Tablet", brand="Apple", model="iPadLivre")
    criar_item(tipo="Notebook")
    baixado = criar_aparelho(model="Baixado")
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (baixado["id"],))
    com_linha = criar_aparelho(model="Ocupado")
    linha = nova_linha(item_id=com_linha["id"])
    colaborador = criar_colaborador("Maria Souza")
    executar_sql("UPDATE items SET colaborador_id=%s WHERE id=%s", (colaborador, celular_livre["id"]))

    def ids(**params):
        resp = client_gestor.get("/api/linhas/aparelhos", params=params)
        assert resp.status_code == 200, resp.text
        return {a["id"] for a in resp.json()}

    assert ids() == {celular_livre["id"], tablet_livre["id"]}
    assert ids(linha_id=linha["id"]) == {celular_livre["id"], tablet_livre["id"], com_linha["id"]}
    assert ids(busca="iPadLivre") == {tablet_livre["id"]}

    por_id = {a["id"]: a for a in client_gestor.get("/api/linhas/aparelhos").json()}
    assert por_id[celular_livre["id"]]["responsavel_nome"] == "Maria Souza"
    assert por_id[celular_livre["id"]]["tipo"] == "Celular"
    assert ids(busca="maria") == {celular_livre["id"]}


def test_busca_por_texto_nao_casa_com_os_digitos_do_numero(client_gestor, nova_linha, criar_aparelho):
    celular = criar_aparelho(identificador="CEL-07")
    nova_linha(numero="74999990001", plano="Controle 20GB", item_id=celular["id"])
    nova_linha(numero="74999990004", plano="Controle 4GB")

    def total(busca):
        return client_gestor.get("/api/linhas", params={"busca": busca}).json()["total"]

    assert total("Controle 4GB") == 1  # não traz a linha de "Controle 20GB" pelos dígitos
    assert total("CEL-07") == 1
    assert total("CEL-99") == 0
    # termo que parece telefone continua achando pelo número, com ou sem máscara
    assert total("(74) 9999") == 2
    assert total("99999-0004") == 1
    assert total("999") == 0  # menos de 4 dígitos não busca por número


def test_put_para_numero_de_linha_cancelada_nao_sugere_cadastrar_de_novo(
    client_gestor, nova_linha, executar_sql
):
    cancelada = nova_linha(numero="74999990001")
    executar_sql("UPDATE linhas SET status='cancelada' WHERE id=%s", (cancelada["id"],))
    outra = nova_linha(numero="74999990002")

    resp = client_gestor.put(f"/api/linhas/{outra['id']}", json={"numero": "74999990001"})

    assert resp.status_code == 409
    assert resp.json()["detail"] == "Esse número pertence a outra linha (cancelada)."


def test_numero_fixo_antigo_de_celular_e_recusado(client_gestor, dados_linha):
    resp = client_gestor.post("/api/linhas", json=dados_linha(numero="7499999999"))

    assert resp.status_code == 422
    assert resp.json()["detail"] == "Número: Celular deve ter 11 dígitos (DDD + 9 + número)."
