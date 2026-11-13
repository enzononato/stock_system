"""Inventário físico: leituras (registrar, corrigir, desfazer, listar)."""
import json
from concurrent.futures import ThreadPoolExecutor

import pytest

pytestmark = pytest.mark.integration


def _criar_local(unidade_id, nome):
    from app.db.database_mysql import get_cursor

    with get_cursor(dict_cursor=False) as cur:
        cur.execute("INSERT INTO locais (unidade_id, nome) VALUES (%s, %s)", (unidade_id, nome))
        return cur.lastrowid


def _criar_colaborador(nome, cpf):
    from app.db.database_mysql import get_cursor

    with get_cursor(dict_cursor=False) as cur:
        cur.execute(
            "INSERT INTO colaboradores (cpf, nome, origem, situacao, criado_em, atualizado_em) "
            "VALUES (%s, %s, 'manual', 'ativo', NOW(), NOW())",
            (cpf, nome),
        )
        return cur.lastrowid


def _abrir(client, unidade_id):
    resp = client.post("/api/inventarios", json={"unidade_id": unidade_id})
    assert resp.status_code == 201, resp.text
    return resp.json()["id"]


def _ler(client, inv, codigo, **extra):
    return client.post(f"/api/inventarios/{inv}/leituras", json={"codigo": codigo, **extra})


def _auditoria(acao):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT * FROM audit_log WHERE entidade='inventario' AND acao=%s ORDER BY id", (acao,)
        )
        return cur.fetchall()


def _contar_leituras():
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT COUNT(*) AS n FROM inventario_leituras")
        return cur.fetchone()["n"]


def test_leitura_pelo_qr_confere_ok(client_tecnico, unidade, criar_item, executar_sql):
    local = _criar_local(unidade["id"], "Sala TI")
    item = criar_item(revenda=unidade["nome"])
    executar_sql("UPDATE items SET local_id=%s WHERE id=%s", (local, item["id"]))
    inv = _abrir(client_tecnico, unidade["id"])

    resp = _ler(client_tecnico, inv, f"http://localhost:5173/e/{item['id']}", local_id_encontrado=local)

    assert resp.status_code == 201, resp.text
    corpo = resp.json()
    assert corpo["resultado"] == "ok"
    assert corpo["item_id"] == item["id"]
    assert corpo["item_tipo"] == "Notebook"
    assert corpo["lido_por"] == "tecnico.teste"
    assert corpo["duplicada"] is False
    assert corpo["detalhe"] is None


def test_leitura_duplicada_nao_conta_duas_vezes(client_tecnico, unidade, criar_item):
    item = criar_item(revenda=unidade["nome"])
    inv = _abrir(client_tecnico, unidade["id"])

    primeira = _ler(client_tecnico, inv, f"http://localhost:5173/e/{item['id']}")
    segunda = _ler(client_tecnico, inv, f"#{item['id']}")

    assert primeira.status_code == 201 and segunda.status_code == 200
    assert segunda.json()["duplicada"] is True
    assert segunda.json()["id"] == primeira.json()["id"]
    assert client_tecnico.get(f"/api/inventarios/{inv}").json()["resumo"]["lidos"] == 1
    assert len(_auditoria("registrar_leitura")) == 1


def test_leitura_concorrente_conta_uma_vez(client_tecnico, unidade, criar_item):
    from app.db.inventario_db import InventarioDBManager

    item = criar_item(revenda=unidade["nome"])
    inv = _abrir(client_tecnico, unidade["id"])
    db = InventarioDBManager()

    def ler(_):
        _, duplicada = db.registrar_leitura(
            inv, codigo=f"#{item['id']}", local_id_encontrado=None,
            colaborador_id_encontrado=None, usuario_conferido=False, observacao=None, por="x",
        )
        return duplicada

    with ThreadPoolExecutor(2) as pool:
        flags = list(pool.map(ler, range(2)))

    assert sorted(flags) == [False, True]
    assert _contar_leituras() == 1


def test_local_errado_com_detalhe(client_tecnico, unidade, criar_item, executar_sql):
    sala = _criar_local(unidade["id"], "Sala TI")
    recepcao = _criar_local(unidade["id"], "Recepção")
    item = criar_item(revenda=unidade["nome"])
    executar_sql("UPDATE items SET local_id=%s WHERE id=%s", (sala, item["id"]))
    inv = _abrir(client_tecnico, unidade["id"])

    corpo = _ler(client_tecnico, inv, str(item["id"]), local_id_encontrado=recepcao).json()

    assert corpo["resultado"] == "local_errado"
    assert corpo["detalhe"] == "Local cadastrado: Sala TI · encontrado em: Recepção."
    assert corpo["local_esperado_nome"] == "Sala TI"
    assert corpo["local_encontrado_nome"] == "Recepção"


def test_item_de_outra_unidade(client_tecnico, unidade, criar_unidade, criar_item):
    outra = criar_unidade()
    item = criar_item(revenda=outra["nome"])
    inv = _abrir(client_tecnico, unidade["id"])

    corpo = _ler(client_tecnico, inv, f"#{item['id']}").json()

    assert corpo["resultado"] == "outra_unidade"
    assert corpo["detalhe"] == f"Cadastrado na unidade {outra['nome'].upper()}."
    assert corpo["item_revenda"] == outra["nome"]


def test_nao_cadastrado_conta_uma_vez(client_tecnico, unidade):
    inv = _abrir(client_tecnico, unidade["id"])

    primeira = _ler(client_tecnico, inv, "XYZ-999")
    segunda = _ler(client_tecnico, inv, "xyz-999")
    assert primeira.status_code == 201 and primeira.json()["item_id"] is None
    assert primeira.json()["resultado"] == "nao_cadastrado"
    assert segunda.status_code == 200 and segunda.json()["duplicada"] is True

    terceira = _ler(client_tecnico, inv, "#99999")
    quarta = _ler(client_tecnico, inv, "99999")
    assert terceira.status_code == 201 and quarta.status_code == 200
    assert quarta.json()["duplicada"] is True


def test_item_baixado_vira_nao_cadastrado(client_tecnico, unidade, criar_item, executar_sql):
    item = criar_item(revenda=unidade["nome"])
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (item["id"],))
    inv = _abrir(client_tecnico, unidade["id"])

    corpo = _ler(client_tecnico, inv, f"#{item['id']}").json()

    assert corpo["resultado"] == "nao_cadastrado"
    assert corpo["detalhe"] == "Equipamento baixado no sistema."
    assert corpo["item_id"] == item["id"]


def test_numero_de_serie_ambiguo_422(client_tecnico, unidade, criar_item):
    criar_item(revenda=unidade["nome"], identificador="SN-DUPLICADO")
    criar_item(revenda=unidade["nome"], identificador="SN-DUPLICADO")
    inv = _abrir(client_tecnico, unidade["id"])

    resp = _ler(client_tecnico, inv, "sn-duplicado")

    assert resp.status_code == 422
    assert resp.json()["detail"] == (
        "Há mais de um equipamento com esse número de série. "
        "Leia o QR code ou digite o número do patrimônio."
    )


def test_leitura_por_numero_de_serie(client_tecnico, unidade, criar_item):
    item = criar_item(revenda=unidade["nome"], identificador="SN-UNICO-1")
    inv = _abrir(client_tecnico, unidade["id"])

    corpo = _ler(client_tecnico, inv, " sn-unico-1 ").json()

    assert corpo["item_id"] == item["id"] and corpo["resultado"] == "ok"
    assert _ler(client_tecnico, inv, f"#{item['id']}").json()["duplicada"] is True


def test_codigo_vazio_422(client_tecnico, unidade):
    inv = _abrir(client_tecnico, unidade["id"])

    resp = _ler(client_tecnico, inv, "   ")

    assert resp.status_code == 422
    assert resp.json()["detail"] == "Código lido: Informe o código lido."


def test_local_de_outra_unidade_422(client_tecnico, unidade, criar_unidade):
    outra = criar_unidade()
    local_alheio = _criar_local(outra["id"], "Sala X")
    inv = _abrir(client_tecnico, unidade["id"])

    resp = _ler(client_tecnico, inv, "#1", local_id_encontrado=local_alheio)

    assert resp.status_code == 422
    assert resp.json()["detail"] == "O local informado não pertence à unidade do inventário."


def test_informar_usuario_por_patch(client_tecnico, unidade, criar_item, executar_sql):
    a = _criar_colaborador("Colab A", "11144477735")
    b = _criar_colaborador("Colab B", "52998224725")
    item = criar_item(revenda=unidade["nome"])
    executar_sql("UPDATE items SET colaborador_id=%s WHERE id=%s", (a, item["id"]))
    inv = _abrir(client_tecnico, unidade["id"])
    leitura = _ler(client_tecnico, inv, f"#{item['id']}").json()
    assert leitura["resultado"] == "ok"
    url = f"/api/inventarios/{inv}/leituras/{leitura['id']}"

    resp = client_tecnico.patch(url, json={"colaborador_id_encontrado": b})
    assert resp.status_code == 200, resp.text
    corpo = resp.json()
    assert corpo["resultado"] == "usuario_diferente"
    assert corpo["usuario_conferido"] is True
    assert corpo["colaborador_encontrado_nome"] == "Colab B"
    assert corpo["usuario_esperado"] == "Colab A"
    assert corpo["detalhe"] == "Usuário cadastrado: Colab A · encontrado com: Colab B."

    resp = client_tecnico.patch(url, json={"usuario_conferido": False})
    assert resp.json()["resultado"] == "ok"
    assert resp.json()["colaborador_id_encontrado"] is None
    assert resp.json()["detalhe"] is None

    resp = client_tecnico.patch(url, json={"colaborador_id_encontrado": a})
    assert resp.json()["resultado"] == "ok" and resp.json()["usuario_conferido"] is True
    assert len(_auditoria("editar_leitura")) == 3


def test_legado_assigned_to_sem_falso_usuario_diferente(client_tecnico, unidade, criar_item, executar_sql):
    fulano = _criar_colaborador("FULANO DE TAL", "11144477735")
    item = criar_item(revenda=unidade["nome"])
    executar_sql("UPDATE items SET assigned_to=%s, colaborador_id=NULL WHERE id=%s", ("Fulano de Tal", item["id"]))
    inv = _abrir(client_tecnico, unidade["id"])
    leitura = _ler(client_tecnico, inv, f"#{item['id']}").json()
    assert leitura["resultado"] == "ok" and leitura["usuario_esperado"] == "Fulano de Tal"
    url = f"/api/inventarios/{inv}/leituras/{leitura['id']}"

    assert client_tecnico.patch(url, json={"colaborador_id_encontrado": fulano}).json()["resultado"] == "ok"

    resp = client_tecnico.patch(url, json={"usuario_conferido": True, "colaborador_id_encontrado": None})
    assert resp.json()["resultado"] == "usuario_diferente"
    assert resp.json()["detalhe"] == "Usuário cadastrado: Fulano de Tal · encontrado com: ninguém."


def test_patch_de_local_reclassifica(client_tecnico, unidade, criar_item, executar_sql):
    sala = _criar_local(unidade["id"], "Sala TI")
    recepcao = _criar_local(unidade["id"], "Recepção")
    item = criar_item(revenda=unidade["nome"])
    executar_sql("UPDATE items SET local_id=%s WHERE id=%s", (sala, item["id"]))
    inv = _abrir(client_tecnico, unidade["id"])
    leitura = _ler(client_tecnico, inv, f"#{item['id']}", local_id_encontrado=recepcao).json()
    assert leitura["resultado"] == "local_errado"
    url = f"/api/inventarios/{inv}/leituras/{leitura['id']}"

    corpo = client_tecnico.patch(url, json={"local_id_encontrado": sala, "observacao": "Conferido de novo"}).json()

    assert corpo["resultado"] == "ok"
    assert corpo["observacao"] == "Conferido de novo"
    assert corpo["atualizado_em"] is not None
    assert client_tecnico.patch(f"/api/inventarios/{inv}/leituras/999999", json={}).status_code == 404


def test_desfazer_leitura_permite_ler_de_novo(client_tecnico, unidade, criar_item):
    item = criar_item(revenda=unidade["nome"])
    inv = _abrir(client_tecnico, unidade["id"])
    leitura = _ler(client_tecnico, inv, f"#{item['id']}").json()

    resp = client_tecnico.delete(f"/api/inventarios/{inv}/leituras/{leitura['id']}")

    assert resp.status_code == 200
    assert resp.json() == {"detail": "Leitura excluída."}
    de_novo = _ler(client_tecnico, inv, f"#{item['id']}")
    assert de_novo.status_code == 201 and de_novo.json()["duplicada"] is False
    (excluida,) = _auditoria("excluir_leitura")
    assert json.loads(excluida["antes"])["leitura_id"] == leitura["id"]
    assert client_tecnico.delete(f"/api/inventarios/{inv}/leituras/{leitura['id']}").status_code == 404


def test_listar_leituras_recentes_primeiro_e_por_resultado(client_tecnico, unidade, criar_item):
    a = criar_item(revenda=unidade["nome"])
    b = criar_item(revenda=unidade["nome"])
    inv = _abrir(client_tecnico, unidade["id"])
    _ler(client_tecnico, inv, f"#{a['id']}")
    _ler(client_tecnico, inv, "DESCONHECIDO")
    ultima = _ler(client_tecnico, inv, f"#{b['id']}").json()

    todas = client_tecnico.get(f"/api/inventarios/{inv}/leituras").json()
    assert todas["total"] == 3
    assert todas["items"][0]["id"] == ultima["id"]

    so_ok = client_tecnico.get(f"/api/inventarios/{inv}/leituras?resultado=ok&limit=1&offset=1").json()
    assert so_ok["total"] == 2 and len(so_ok["items"]) == 1

    nao_cad = client_tecnico.get(f"/api/inventarios/{inv}/leituras?resultado=nao_cadastrado").json()
    assert nao_cad["total"] == 1 and nao_cad["items"][0]["codigo_lido"] == "DESCONHECIDO"
    assert client_tecnico.get(f"/api/inventarios/{inv}/leituras?limit=101").status_code == 422
    assert client_tecnico.get("/api/inventarios/999999/leituras").status_code == 404


def test_aprendiz_403_nas_leituras(client_aprendiz, client_tecnico, unidade):
    inv = _abrir(client_tecnico, unidade["id"])
    for resp in (
        client_aprendiz.get(f"/api/inventarios/{inv}/leituras"),
        _ler(client_aprendiz, inv, "#1"),
        client_aprendiz.patch(f"/api/inventarios/{inv}/leituras/1", json={}),
        client_aprendiz.delete(f"/api/inventarios/{inv}/leituras/1"),
    ):
        assert resp.status_code == 403
        assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."


class _CursorQueEscondeAChave:
    """Faz a primeira consulta de duplicidade (sem FOR UPDATE) parecer vazia, para
    forçar o caminho defensivo do erro 1062 no INSERT."""

    def __init__(self, cur):
        self._cur = cur
        self._esconder = False
        self._feito = False

    def execute(self, sql, params=None):
        if (
            not self._feito
            and sql.startswith("SELECT id FROM inventario_leituras")
            and "FOR UPDATE" not in sql
        ):
            self._feito = True
            self._esconder = True
            return 0
        self._esconder = False
        return self._cur.execute(sql, params)

    def fetchone(self):
        return None if self._esconder else self._cur.fetchone()

    def __getattr__(self, nome):
        return getattr(self._cur, nome)


def _manager_com_cursor_que_esconde(monkeypatch):
    from contextlib import contextmanager

    from app.db import inventario_db

    real = inventario_db.get_cursor

    @contextmanager
    def falso(*args, **kwargs):
        with real(*args, **kwargs) as cur:
            yield _CursorQueEscondeAChave(cur)

    monkeypatch.setattr(inventario_db, "get_cursor", falso)
    return inventario_db.InventarioDBManager()


def _ler_direto(db, inv, codigo):
    return db.registrar_leitura(
        inv, codigo=codigo, local_id_encontrado=None, colaborador_id_encontrado=None,
        usuario_conferido=False, observacao=None, por="x",
    )


def test_1062_no_insert_rele_pela_chave_com_bloqueio(client_tecnico, unidade, criar_item, monkeypatch):
    item = criar_item(revenda=unidade["nome"])
    inv = _abrir(client_tecnico, unidade["id"])
    primeira = _ler(client_tecnico, inv, f"#{item['id']}").json()
    db = _manager_com_cursor_que_esconde(monkeypatch)

    leitura, duplicada = _ler_direto(db, inv, f"#{item['id']}")

    assert duplicada is True and leitura["id"] == primeira["id"]
    assert _contar_leituras() == 1


def test_1062_por_item_com_outra_chave_nao_e_confundido_com_duplicada(
    client_tecnico, unidade, criar_item, executar_sql
):
    """Linha antiga do mesmo item com chave diferente: o INSERT bate no UNIQUE do item e o
    erro sobe (nunca devolve uma leitura qualquer)."""
    import pymysql

    from app.db.inventario_db import InventarioDBManager

    item = criar_item(revenda=unidade["nome"])
    inv = _abrir(client_tecnico, unidade["id"])
    executar_sql(
        "INSERT INTO inventario_leituras (inventario_id, item_id, chave, codigo_lido, lido_por, "
        "lido_em, resultado) VALUES (%s, %s, 'codigo:LEGADO', 'x', 'x', NOW(), 'ok')",
        (inv, item["id"]),
    )

    with pytest.raises(pymysql.err.IntegrityError):
        _ler_direto(InventarioDBManager(), inv, f"#{item['id']}")
    assert _contar_leituras() == 1
