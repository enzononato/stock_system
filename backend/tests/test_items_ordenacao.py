import pytest
from app.db.inventory_manager_db import montar_order_by_items


def _ids(client, **params):
    resp = client.get("/api/items", params=params)
    assert resp.status_code == 200, resp.text
    return [i["id"] for i in resp.json()["items"]]


def _vincular(executar_sql, item_id, *, cpf, nome, setor, cargo):
    executar_sql(
        "INSERT INTO colaboradores (cpf, nome, setor, cargo, situacao, origem, criado_em, atualizado_em) "
        "VALUES (%s, %s, %s, %s, 'ativo', 'manual', NOW(), NOW())", (cpf, nome, setor, cargo))
    executar_sql("UPDATE items SET colaborador_id = (SELECT id FROM colaboradores WHERE cpf = %s) WHERE id = %s",
                 (cpf, item_id))


@pytest.mark.unit
@pytest.mark.parametrize("sort,direcao,esperado", [
    ("id", "desc", "i.id DESC"),
    ("id", "asc", "i.id ASC"),
    ("tipo", "asc", "NULLIF(i.tipo, '') IS NULL, i.tipo ASC, i.id DESC"),
    ("equipamento", "desc", "NULLIF(i.brand, '') IS NULL, i.brand DESC, i.model DESC, i.id DESC"),
    ("status", "asc", "NULLIF(CAST(i.status AS CHAR), '') IS NULL, CAST(i.status AS CHAR) ASC, i.id DESC"),
    ("cadastro", "desc", "i.date_registered DESC, i.id DESC"),
])
def test_montar_order_by_items(sort, direcao, esperado):
    assert montar_order_by_items(sort, direcao) == esperado


@pytest.mark.unit
@pytest.mark.parametrize("sort,direcao", [("id; DROP TABLE items", "asc"), ("TIPO", "asc"), ("tipo", "up")])
def test_montar_order_by_items_recusa_fora_da_lista(sort, direcao):
    with pytest.raises(ValueError):
        montar_order_by_items(sort, direcao)


@pytest.mark.integration
class TestOrdenacaoDeItems:
    def test_setor_do_colaborador_com_sem_colaborador_no_fim(self, client_gestor, criar_item, executar_sql):
        ti, com, sem = criar_item(), criar_item(), criar_item()
        _vincular(executar_sql, ti["id"], cpf="52998224725", nome="ANA", setor="TI", cargo="Analista")
        _vincular(executar_sql, com["id"], cpf="11144477735", nome="BRUNO", setor="Comercial", cargo="Vendedor")
        assert _ids(client_gestor, sort="setor", dir="asc") == [com["id"], ti["id"], sem["id"]]
        assert _ids(client_gestor, sort="setor", dir="desc") == [ti["id"], com["id"], sem["id"]]
        assert _ids(client_gestor, sort="cargo", dir="asc") == [ti["id"], com["id"], sem["id"]]

    def test_sort_fora_da_lista_branca_responde_422(self, client_gestor, item_disponivel):
        resp = client_gestor.get("/api/items", params={"sort": "id; DROP TABLE items"})
        assert resp.status_code == 422
        assert resp.json()["detail"] == "Ordenação inválida."
        assert client_gestor.get("/api/items").json()["total"] == 1

    def test_padrao_continua_id_desc(self, client_gestor, criar_item):
        ids = [criar_item()["id"] for _ in range(3)]
        assert _ids(client_gestor) == sorted(ids, reverse=True)

    def test_tipo_asc_e_desc(self, client_gestor, criar_item):
        notebook = criar_item(tipo="Notebook")["id"]
        desktop = criar_item(tipo="Desktop")["id"]
        monitor = criar_item(tipo="Monitor")["id"]
        assert _ids(client_gestor, sort="tipo", dir="asc") == [desktop, monitor, notebook]
        assert _ids(client_gestor, sort="tipo", dir="desc") == [notebook, monitor, desktop]

    def test_status_em_ordem_alfabetica(self, client_gestor, criar_item, executar_sql):
        por_status = {}
        for status in ("Pendente Devolução", "Disponível", "Pendente", "Indisponível"):
            item = criar_item()
            executar_sql("UPDATE items SET status=%s WHERE id=%s", (status, item["id"]))
            por_status[status] = item["id"]
        assert _ids(client_gestor, sort="status", dir="asc") == [
            por_status["Disponível"], por_status["Indisponível"],
            por_status["Pendente"], por_status["Pendente Devolução"],
        ]

    def test_colaborador_usa_assigned_to_e_vazios_no_fim(self, client_gestor, criar_item, executar_sql):
        bruno, ana, vazio = criar_item(), criar_item(), criar_item()
        executar_sql("UPDATE items SET assigned_to=%s WHERE id=%s", ("Bruno", bruno["id"]))
        executar_sql("UPDATE items SET assigned_to=%s WHERE id=%s", ("Ana", ana["id"]))
        executar_sql("UPDATE items SET assigned_to=%s WHERE id=%s", ("", vazio["id"]))
        assert _ids(client_gestor, sort="colaborador", dir="asc") == [ana["id"], bruno["id"], vazio["id"]]
        assert _ids(client_gestor, sort="colaborador", dir="desc") == [bruno["id"], ana["id"], vazio["id"]]

    def test_colaborador_vinculado_ordena_pelo_nome_do_colaborador(self, client_gestor, criar_item, executar_sql):
        # A coluna mostra colaborador_nome (ou assigned_to): a ordenação segue a mesma regra.
        so_texto, vinculado, so_vinculo, vazio = criar_item(), criar_item(), criar_item(), criar_item()
        executar_sql("UPDATE items SET assigned_to=%s WHERE id=%s", ("Bruno", so_texto["id"]))
        executar_sql("UPDATE items SET assigned_to=%s WHERE id=%s", ("Zeca", vinculado["id"]))
        _vincular(executar_sql, vinculado["id"], cpf="52998224725", nome="ANA", setor="TI", cargo="Analista")
        _vincular(executar_sql, so_vinculo["id"], cpf="11144477735", nome="CARLA", setor="TI", cargo="Analista")
        assert _ids(client_gestor, sort="colaborador", dir="asc") == [
            vinculado["id"], so_texto["id"], so_vinculo["id"], vazio["id"],
        ]
        assert _ids(client_gestor, sort="colaborador", dir="desc") == [
            so_vinculo["id"], so_texto["id"], vinculado["id"], vazio["id"],
        ]

    def test_ordenacao_com_paginacao_nao_repete(self, client_gestor, criar_item):
        c = criar_item(tipo="C")["id"]
        a = criar_item(tipo="A")["id"]
        b = criar_item(tipo="B")["id"]
        p1 = client_gestor.get("/api/items", params={"sort": "tipo", "dir": "asc", "limit": 2, "offset": 0}).json()
        p2 = client_gestor.get("/api/items", params={"sort": "tipo", "dir": "asc", "limit": 2, "offset": 2}).json()
        assert [i["id"] for i in p1["items"]] == [a, b]
        assert [i["id"] for i in p2["items"]] == [c]
        assert p1["total"] == 3 and p2["total"] == 3

    def test_filtro_colaborador_combina_com_ordenacao(self, client_gestor, criar_item, executar_sql):
        notebook_a = criar_item(tipo="Notebook")["id"]
        desktop_a = criar_item(tipo="Desktop")["id"]
        item_b = criar_item(tipo="Monitor")["id"]
        _vincular(executar_sql, notebook_a, cpf="52998224725", nome="ANA", setor="TI", cargo="Analista")
        executar_sql("UPDATE items SET colaborador_id = (SELECT id FROM colaboradores WHERE cpf = %s) WHERE id = %s",
                     ("52998224725", desktop_a))
        _vincular(executar_sql, item_b, cpf="11144477735", nome="BRUNO", setor="Comercial", cargo="Vendedor")
        colab_a = client_gestor.get("/api/items/%d" % notebook_a).json()["colaborador_id"]
        resp = client_gestor.get("/api/items", params={"colaborador_id": colab_a, "sort": "tipo", "dir": "asc"}).json()
        assert [i["id"] for i in resp["items"]] == [desktop_a, notebook_a]
        assert resp["total"] == 2

    def test_dir_invalido_responde_422(self, client_gestor, item_disponivel):
        resp = client_gestor.get("/api/items", params={"dir": "up"})
        assert resp.status_code == 422
        assert resp.json()["detail"] == "Direção da ordenação inválida: use asc ou desc."
