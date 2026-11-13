"""Locais por unidade, vínculos de colaborador/local no item e filtro por colaborador."""
import pytest

from tests.fixtures_colaboradores import CPF_ATIVO, CPF_DESLIGADO

pytestmark = pytest.mark.integration

MSG_DUPLICADO = "Já existe um local com esse nome nesta unidade."


def _auditoria(entidade, acao):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT * FROM audit_log WHERE entidade = %s AND acao = %s ORDER BY id",
            (entidade, acao),
        )
        return cur.fetchall()


def _criar_local(client, unidade_id, nome="Almoxarifado", **extra):
    resp = client.post("/api/locais", json={"unidade_id": unidade_id, "nome": nome, **extra})
    assert resp.status_code == 201, resp.text
    return resp.json()


def test_criar_e_listar_locais_por_unidade(client_gestor, criar_unidade):
    u1 = criar_unidade()
    u2 = criar_unidade()
    criado = _criar_local(client_gestor, u1["id"], "Sala de TI", descricao="1o andar")
    _criar_local(client_gestor, u1["id"], "Almoxarifado")
    _criar_local(client_gestor, u2["id"], "Sala de TI")

    assert criado["nome"] == "Sala de TI"
    assert criado["descricao"] == "1o andar"
    assert criado["ativo"] is True
    assert criado["unidade_id"] == u1["id"]
    assert criado["unidade_nome"] == u1["nome"]

    da_u1 = client_gestor.get("/api/locais", params={"unidade_id": u1["id"]}).json()
    assert [loc["nome"] for loc in da_u1] == ["Almoxarifado", "Sala de TI"]  # ordenado por nome
    assert all(loc["unidade_id"] == u1["id"] for loc in da_u1)
    assert len(client_gestor.get("/api/locais").json()) == 3

    obtido = client_gestor.get(f"/api/locais/{criado['id']}")
    assert obtido.status_code == 200
    assert obtido.json()["nome"] == "Sala de TI"
    assert client_gestor.get("/api/locais/99999").status_code == 404

    linha = _auditoria("local", "criar")[0]
    assert linha["entidade_id"] == str(criado["id"])
    assert linha["username"] == "gestor.teste"


def test_local_duplicado_na_mesma_unidade_400(client_gestor, criar_unidade):
    u1 = criar_unidade()
    u2 = criar_unidade()
    _criar_local(client_gestor, u1["id"], "Sala de TI")

    resp = client_gestor.post("/api/locais", json={"unidade_id": u1["id"], "nome": "Sala de TI"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_DUPLICADO

    # mesma unidade, só difere em espaços nas pontas: também é duplicado
    resp = client_gestor.post("/api/locais", json={"unidade_id": u1["id"], "nome": "  Sala de TI "})
    assert resp.status_code == 400

    # outra unidade pode ter o mesmo nome
    assert client_gestor.post("/api/locais", json={"unidade_id": u2["id"], "nome": "Sala de TI"}).status_code == 201

    # renomear para um nome já usado na unidade
    outro = _criar_local(client_gestor, u1["id"], "Almoxarifado")
    resp = client_gestor.put(f"/api/locais/{outro['id']}", json={"nome": "Sala de TI"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_DUPLICADO


def test_criar_local_valida_entrada(client_gestor, criar_unidade):
    u = criar_unidade()
    assert client_gestor.post("/api/locais", json={"unidade_id": 99999, "nome": "X"}).status_code == 404
    assert client_gestor.post("/api/locais", json={"unidade_id": u["id"], "nome": "   "}).status_code == 422
    assert client_gestor.post("/api/locais", json={"unidade_id": u["id"]}).status_code == 422
    assert client_gestor.post("/api/locais", json={"unidade_id": u["id"], "nome": "a" * 151}).status_code == 422


def test_local_exige_gestor_ou_tecnico(client, client_tecnico, client_aprendiz, criar_unidade):
    u = criar_unidade()
    corpo = {"unidade_id": u["id"], "nome": "Sala"}
    assert client.get("/api/locais").status_code == 401
    assert client.post("/api/locais", json=corpo).status_code == 401

    assert client_aprendiz.post("/api/locais", json=corpo).status_code == 403
    assert client_aprendiz.get("/api/locais").status_code == 200  # leitura: qualquer autenticado

    criado = client_tecnico.post("/api/locais", json=corpo)
    assert criado.status_code == 201
    local_id = criado.json()["id"]
    assert client_aprendiz.put(f"/api/locais/{local_id}", json={"nome": "Outra"}).status_code == 403
    assert client_tecnico.put(f"/api/locais/{local_id}", json={"nome": "Outra"}).status_code == 200


def test_atualizar_local_parcial_e_auditado(client_gestor, criar_unidade):
    u = criar_unidade()
    local = _criar_local(client_gestor, u["id"], "Sala", descricao="Antiga")
    url = f"/api/locais/{local['id']}"

    # só o nome: descrição e ativo ficam
    assert client_gestor.put(url, json={"nome": "Sala 2"}).status_code == 200
    atual = client_gestor.get(url).json()
    assert (atual["nome"], atual["descricao"], atual["ativo"]) == ("Sala 2", "Antiga", True)

    # só ativo
    assert client_gestor.put(url, json={"ativo": False}).status_code == 200
    atual = client_gestor.get(url).json()
    assert (atual["nome"], atual["ativo"]) == ("Sala 2", False)

    # descrição pode ser limpa explicitamente
    assert client_gestor.put(url, json={"descricao": None}).status_code == 200
    assert client_gestor.get(url).json()["descricao"] is None

    # corpo sem nenhum campo, local inexistente, nome em branco
    assert client_gestor.put(url, json={}).status_code == 400
    assert client_gestor.put("/api/locais/99999", json={"nome": "X"}).status_code == 404
    assert client_gestor.put(url, json={"nome": "  "}).status_code == 422

    assert len(_auditoria("local", "editar")) == 3


def test_listar_locais_inativos_so_com_include_inactive(client_gestor, criar_unidade):
    u = criar_unidade()
    ativo = _criar_local(client_gestor, u["id"], "Ativo")
    inativo = _criar_local(client_gestor, u["id"], "Inativo")
    client_gestor.put(f"/api/locais/{inativo['id']}", json={"ativo": False})

    padrao = client_gestor.get("/api/locais", params={"unidade_id": u["id"]}).json()
    assert [loc["id"] for loc in padrao] == [ativo["id"]]

    todos = client_gestor.get(
        "/api/locais", params={"unidade_id": u["id"], "include_inactive": "true"}
    ).json()
    assert {loc["id"] for loc in todos} == {ativo["id"], inativo["id"]}


def test_locais_de_unidade_inativa_continuam_listados(client_gestor, criar_unidade, executar_sql):
    u = criar_unidade()
    local = _criar_local(client_gestor, u["id"], "Sala")
    executar_sql("UPDATE unidades SET is_active = 0 WHERE id = %s", (u["id"],))
    lista = client_gestor.get("/api/locais", params={"unidade_id": u["id"]}).json()
    assert [loc["id"] for loc in lista] == [local["id"]]


def _colaborador(client, cpf, nome, **extra):
    resp = client.post("/api/colaboradores", json={"cpf": cpf, "nome": nome, **extra})
    assert resp.status_code == 201, resp.text
    return resp.json()


def test_item_traz_colaborador_setor_cargo_superior_e_local(
    client_gestor, criar_unidade, criar_item, executar_sql
):
    u = criar_unidade()
    local = _criar_local(client_gestor, u["id"], "Sala de TI")
    chefe = _colaborador(client_gestor, CPF_DESLIGADO, "Carlos Souza")
    colab = _colaborador(
        client_gestor, CPF_ATIVO, "Maria da Silva",
        cargo="Analista", setor="TI", superior_id=chefe["id"],
    )
    item = criar_item()
    sem_vinculo = criar_item()

    antes = client_gestor.get(f"/api/items/{item['id']}").json()
    for campo in ("colaborador_id", "local_id", "local_nome", "colaborador_nome",
                  "colaborador_setor", "colaborador_cargo", "superior_nome"):
        assert antes[campo] is None

    executar_sql(
        "UPDATE items SET colaborador_id = %s, local_id = %s WHERE id = %s",
        (colab["id"], local["id"], item["id"]),
    )

    corpo = client_gestor.get(f"/api/items/{item['id']}").json()
    assert corpo["id"] == item["id"]
    assert corpo["colaborador_id"] == colab["id"]
    assert corpo["colaborador_nome"] == "Maria da Silva"
    assert corpo["colaborador_setor"] == "TI"
    assert corpo["colaborador_cargo"] == "Analista"
    assert corpo["superior_nome"] == "Carlos Souza"
    assert corpo["local_id"] == local["id"]
    assert corpo["local_nome"] == "Sala de TI"

    # a listagem traz o mesmo, sem duplicar nem reordenar linhas
    lista = client_gestor.get("/api/items").json()
    assert lista["total"] == 2
    assert [i["id"] for i in lista["items"]] == [sem_vinculo["id"], item["id"]]
    do_item = lista["items"][1]
    assert do_item["colaborador_nome"] == "Maria da Silva"
    assert do_item["superior_nome"] == "Carlos Souza"
    assert do_item["local_nome"] == "Sala de TI"
    assert lista["items"][0]["colaborador_nome"] is None


def test_filtro_por_colaborador_id(client_gestor, criar_item, executar_sql):
    a = _colaborador(client_gestor, CPF_ATIVO, "Maria da Silva")
    b = _colaborador(client_gestor, CPF_DESLIGADO, "Carlos Souza")
    i1, i2, i3 = criar_item(tipo="Notebook"), criar_item(tipo="Monitor"), criar_item(tipo="Notebook")
    executar_sql("UPDATE items SET colaborador_id = %s WHERE id IN (%s, %s)", (a["id"], i1["id"], i2["id"]))
    executar_sql("UPDATE items SET colaborador_id = %s WHERE id = %s", (b["id"], i3["id"]))

    resp = client_gestor.get("/api/items", params={"colaborador_id": a["id"]}).json()
    assert resp["total"] == 2
    assert {i["id"] for i in resp["items"]} == {i1["id"], i2["id"]}

    # combinável com os outros filtros
    resp = client_gestor.get("/api/items", params={"colaborador_id": a["id"], "tipo": "Notebook"}).json()
    assert [i["id"] for i in resp["items"]] == [i1["id"]]

    assert client_gestor.get("/api/items", params={"colaborador_id": 99999}).json()["total"] == 0
    assert client_gestor.get("/api/items").json()["total"] == 3


def test_apagar_colaborador_ou_local_solta_o_vinculo(
    client_gestor, criar_unidade, criar_item, executar_sql
):
    u = criar_unidade()
    local = _criar_local(client_gestor, u["id"], "Sala")
    colab = _colaborador(client_gestor, CPF_ATIVO, "Maria da Silva")
    item = criar_item()
    executar_sql(
        "UPDATE items SET colaborador_id = %s, local_id = %s WHERE id = %s",
        (colab["id"], local["id"], item["id"]),
    )
    executar_sql("DELETE FROM colaboradores WHERE id = %s", (colab["id"],))
    executar_sql("DELETE FROM locais WHERE id = %s", (local["id"],))
    corpo = client_gestor.get(f"/api/items/{item['id']}").json()
    assert corpo["colaborador_id"] is None and corpo["local_id"] is None
