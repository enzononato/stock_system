"""Cadastro de colaboradores: sincronização com a base de RH (fixture `base_rh`),
complemento manual e API `/api/colaboradores`."""
import json

import pytest

from app.schemas.validators import normalizar_telefone, validar_email
from tests.fixtures_colaboradores import (
    CPF_AFASTADO,
    CPF_ATIVO,
    CPF_DESLIGADO,
    CPF_MASCARADO,
    CPF_MASCARADO_FORMATADO,
)

# CPFs válidos que não estão na base fictícia.
CPF_CARLOS = "39053344705"
CPF_OUTRO = "12345678909"

MSG_SO_GESTOR = "Acesso restrito ao Gestor."
MSG_INDISPONIVEL = "Base de colaboradores indisponível no momento."


# ── Unitários (sem banco) ───────────────────────────────────────────────────

@pytest.mark.unit
@pytest.mark.parametrize(
    "entrada,esperado",
    [
        (None, None),
        ("", None),
        ("   ", None),
        ("Maria@Revalle.com.br", "maria@revalle.com.br"),
        ("  joao@x.com  ", "joao@x.com"),
    ],
)
def test_validar_email_valido(entrada, esperado):
    assert validar_email(entrada) == esperado


@pytest.mark.unit
@pytest.mark.parametrize("entrada", ["sem-arroba", "a@b", "a b@c.com", "@x.com", "a@.com", "a@@b.com"])
def test_validar_email_invalido(entrada):
    with pytest.raises(ValueError, match="E-mail inválido."):
        validar_email(entrada)


@pytest.mark.unit
@pytest.mark.parametrize(
    "entrada,esperado",
    [
        (None, None),
        ("", None),
        ("+55 (74) 99999-0000", "74999990000"),
        ("(74) 99999-0000", "74999990000"),
        ("74 3611-0000", "7436110000"),
        ("5574999990000", "74999990000"),
        ("557436110000", "7436110000"),
        ("11 98888-7777", "11988887777"),
    ],
)
def test_normalizar_telefone_valido(entrada, esperado):
    assert normalizar_telefone(entrada) == esperado


@pytest.mark.unit
@pytest.mark.parametrize(
    "entrada",
    ["99999-0000", "(10) 99999-0000", "(74) 9999-000", "abc", "123456789012345"],
)
def test_normalizar_telefone_invalido(entrada):
    with pytest.raises(ValueError, match="Telefone inválido: informe DDD \\+ número."):
        normalizar_telefone(entrada)


@pytest.mark.unit
def test_validar_email_e_telefone():
    assert validar_email("A@B.CO") == "a@b.co"
    assert normalizar_telefone("+55 (74) 99999-0000") == "74999990000"
    with pytest.raises(ValueError):
        validar_email("x")
    with pytest.raises(ValueError):
        normalizar_telefone("123")


# ── Infra ───────────────────────────────────────────────────────────────────


@pytest.fixture
def colab_db(_test_schema, limpar_banco):
    from app.db.colaborador_db import ColaboradorDBManager

    return ColaboradorDBManager()


def _sincronizar(db):
    from app.db.colaborador_db import executar_sincronizacao
    from app.services.audit import AuditContext

    return executar_sincronizacao(db, AuditContext.sistema())


def _por_cpf(db, cpf):
    return db.obter_por_cpf(cpf)


def _auditoria(entidade, acao):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT * FROM audit_log WHERE entidade=%s AND acao=%s ORDER BY id", (entidade, acao)
        )
        linhas = list(cur.fetchall())
    for linha in linhas:
        linha["antes"] = json.loads(linha["antes"]) if linha["antes"] else None
        linha["depois"] = json.loads(linha["depois"]) if linha["depois"] else None
    return linhas


def _id_por_nome(db, nome):
    itens, _ = db.listar(busca=nome, limit=10, offset=0)
    assert len(itens) == 1, itens
    return itens[0]["id"]


# ── Sincronização ───────────────────────────────────────────────────────────

@pytest.mark.integration
def test_sincronizar_cria_e_atualiza(colab_db, base_rh):
    resumo = _sincronizar(colab_db)
    assert (resumo.total_base, resumo.criados, resumo.atualizados, resumo.desligados) == (4, 4, 0, 0)
    assert (resumo.ignorados, resumo.duplicados) == (0, 0)

    maria = _por_cpf(colab_db, CPF_ATIVO)
    assert maria["nome"] == "Maria da Silva"
    assert maria["cargo"] == "Analista"
    assert maria["setor"] == "TI"
    assert maria["unidade_rh"] == "Revalle Juazeiro"
    assert maria["matricula"] == "1001"
    assert maria["email"] == "maria@revalle.com.br"
    assert maria["telefone"] == "74999990001"
    assert maria["situacao"] == "ativo"
    assert maria["origem"] == "rh"
    assert maria["email_manual"] == 0
    assert maria["sincronizado_em"] is not None
    assert "data_nascimento" not in maria
    assert _por_cpf(colab_db, CPF_DESLIGADO)["situacao"] == "desligado"
    assert _por_cpf(colab_db, CPF_AFASTADO)["situacao"] == "afastado"
    assert _por_cpf(colab_db, CPF_MASCARADO)["nome"] == "Pedro Alves"

    base_rh('UPDATE "colaboradores-revalle" SET "Cargo" = %s WHERE "CPF" = %s', ("Analista Sênior", CPF_ATIVO))
    resumo = _sincronizar(colab_db)
    assert (resumo.criados, resumo.atualizados, resumo.desligados) == (0, 1, 0)
    assert _por_cpf(colab_db, CPF_ATIVO)["cargo"] == "Analista Sênior"

    resumo = _sincronizar(colab_db)
    assert (resumo.criados, resumo.atualizados, resumo.desligados) == (0, 0, 0)


@pytest.mark.integration
def test_sync_preserva_email_manual(colab_db, base_rh):
    _sincronizar(colab_db)
    maria_id = _por_cpf(colab_db, CPF_ATIVO)["id"]
    colab_db.atualizar(maria_id, {"email": "maria.nova@revalle.com.br"})

    base_rh('UPDATE "colaboradores-revalle" SET "E-mail" = %s, "Celular" = %s WHERE "CPF" = %s',
            ("outro@revalle.com.br", "(74) 98888-7777", CPF_ATIVO))
    _sincronizar(colab_db)

    maria = _por_cpf(colab_db, CPF_ATIVO)
    assert maria["email"] == "maria.nova@revalle.com.br"
    assert maria["email_manual"] == 1
    # Telefone não foi editado à mão: segue a base.
    assert maria["telefone"] == "74988887777"
    assert maria["telefone_manual"] == 0


@pytest.mark.integration
def test_ausente_da_base_vira_desligado(colab_db, base_rh):
    _sincronizar(colab_db)
    manual_id = colab_db.criar_manual({"cpf": CPF_CARLOS, "nome": "Carlos Souza"})
    base_rh('DELETE FROM "colaboradores-revalle" WHERE "CPF" = %s', (CPF_ATIVO,))

    resumo = _sincronizar(colab_db)

    assert resumo.desligados == 1  # João já estava desligado: não conta de novo
    assert _por_cpf(colab_db, CPF_ATIVO)["situacao"] == "desligado"
    assert colab_db.obter(manual_id)["situacao"] == "ativo"  # origem manual não é tocada


@pytest.mark.integration
def test_unidade_mapeada_por_alias(colab_db, base_rh, criar_unidade):
    # Alias com acento/caixa diferentes do RH; o nome da unidade nada tem a ver com o do RH.
    juazeiro = criar_unidade(nome="Matriz BA", alias_rh="RÉVALLE juazeiro")
    # Sem alias: casa por "nome contém o nome do RH".
    petrolina = criar_unidade(nome="Revalle Petrolina - PE")

    _sincronizar(colab_db)

    assert _por_cpf(colab_db, CPF_ATIVO)["unidade_id"] == juazeiro["id"]
    assert _por_cpf(colab_db, CPF_ATIVO)["unidade_nome"] == "Matriz BA"
    assert _por_cpf(colab_db, CPF_AFASTADO)["unidade_id"] == petrolina["id"]
    assert _por_cpf(colab_db, CPF_MASCARADO)["unidade_id"] == petrolina["id"]


@pytest.mark.integration
def test_unidade_ambigua_por_nome_nao_e_adivinhada(colab_db, base_rh, criar_unidade):
    criar_unidade(nome="Revalle Petrolina Centro")
    criar_unidade(nome="Revalle Petrolina Norte")

    _sincronizar(colab_db)

    assert _por_cpf(colab_db, CPF_AFASTADO)["unidade_id"] is None


@pytest.mark.integration
def test_superior_preenchido_pelo_gestor_da_base(colab_db, base_rh):
    _sincronizar(colab_db)
    ana = _por_cpf(colab_db, CPF_AFASTADO)
    pedro = _por_cpf(colab_db, CPF_MASCARADO)
    assert pedro["superior_id"] == ana["id"]  # gestor "Ana Lima"
    assert pedro["superior_nome"] == "Ana Lima"
    assert _por_cpf(colab_db, CPF_ATIVO)["superior_id"] is None  # "Carlos Souza" ainda não existe

    base_rh(
        'INSERT INTO "colaboradores-revalle" ("CPF", "Nome Completo", "Cargo") VALUES (%s, %s, %s)',
        (CPF_CARLOS, "Carlos Souza", "Diretor"),
    )
    _sincronizar(colab_db)

    carlos = _por_cpf(colab_db, CPF_CARLOS)
    assert _por_cpf(colab_db, CPF_ATIVO)["superior_id"] == carlos["id"]
    assert _por_cpf(colab_db, CPF_DESLIGADO)["superior_id"] == carlos["id"]


@pytest.mark.integration
def test_superior_ambiguo_nao_e_preenchido(colab_db, base_rh):
    base_rh(
        'INSERT INTO "colaboradores-revalle" ("CPF", "Nome Completo") VALUES (%s, %s), (%s, %s)',
        (CPF_CARLOS, "Carlos Souza", CPF_OUTRO, "CARLOS  souza"),
    )
    _sincronizar(colab_db)
    assert _por_cpf(colab_db, CPF_ATIVO)["superior_id"] is None


@pytest.mark.integration
def test_sync_ignora_cpf_invalido_e_deduplica(colab_db, base_rh):
    base_rh(
        'INSERT INTO "colaboradores-revalle" ("CPF", "Nome Completo") VALUES (%s, %s), (%s, %s)',
        ("11111111111", "Fulano CPF Falso", "12345678900", "Beltrano Dígito Errado"),
    )
    # Mesmo CPF de Maria com um nome que vem DEPOIS na ordem alfabética: vale a Maria.
    base_rh(
        'INSERT INTO "colaboradores-revalle" ("CPF", "Nome Completo", "Cargo") VALUES (%s, %s, %s)',
        (CPF_ATIVO, "Zuleica Repetida", "Cargo Duplicado"),
    )

    resumo = _sincronizar(colab_db)

    assert resumo.total_base == 7
    assert resumo.criados == 4
    assert resumo.ignorados == 2
    assert resumo.duplicados == 1
    assert _por_cpf(colab_db, CPF_ATIVO)["nome"] == "Maria da Silva"
    assert _por_cpf(colab_db, "11111111111") is None
    log = colab_db.ultima_sync()
    assert (log["ignorados"], log["duplicados"], log["status"]) == (2, 1, "ok")


@pytest.mark.integration
def test_executar_sincronizacao_registra_log_e_auditoria(colab_db, base_rh):
    _sincronizar(colab_db)

    log = colab_db.ultima_sync()
    assert log["status"] == "ok"
    assert (log["total_base"], log["criados"]) == (4, 4)
    assert log["concluido_em"] is not None and log["erro"] is None
    linha = _auditoria("colaborador", "sincronizar")[0]
    assert linha["username"] == "sistema"
    assert linha["depois"]["criados"] == 4


@pytest.mark.integration
def test_ultima_sync_sem_historico(colab_db):
    assert colab_db.ultima_sync() is None


# ── Endpoint de sincronização ───────────────────────────────────────────────

@pytest.mark.integration
def test_sincronizar_endpoint_exige_gestor(client, client_tecnico, client_gestor, base_rh):
    assert client.post("/api/colaboradores/sincronizar").status_code == 401
    resp = client_tecnico.post("/api/colaboradores/sincronizar")
    assert resp.status_code == 403
    assert resp.json()["detail"] == MSG_SO_GESTOR

    resp = client_gestor.post("/api/colaboradores/sincronizar")
    assert resp.status_code == 200
    assert resp.json() == {
        "total_base": 4, "criados": 4, "atualizados": 0, "desligados": 0, "ignorados": 0, "duplicados": 0,
    }
    status = client_gestor.get("/api/colaboradores/sincronizacao").json()
    assert status["configurado"] is True
    assert status["ultima"]["status"] == "ok"
    assert status["ultima"]["criados"] == 4
    assert _auditoria("colaborador", "sincronizar")[0]["username"] == "gestor.teste"


@pytest.mark.integration
def test_base_fora_do_ar_responde_503_e_registra_erro(client_gestor, monkeypatch):
    from app.core import config as _config

    monkeypatch.setattr(
        _config.settings, "COLABORADORES_DATABASE_URL", "postgresql://x:y@127.0.0.1:1/z"
    )

    resp = client_gestor.post("/api/colaboradores/sincronizar")

    assert resp.status_code == 503
    assert resp.json()["detail"] == MSG_INDISPONIVEL
    ultima = client_gestor.get("/api/colaboradores/sincronizacao").json()["ultima"]
    assert ultima["status"] == "erro"
    assert ultima["concluido_em"] is not None
    assert ultima["erro"] == MSG_INDISPONIVEL
    assert "x:y" not in json.dumps(ultima)


@pytest.mark.integration
def test_base_mal_configurada_responde_503_e_nao_500(client_gestor, base_rh, monkeypatch):
    from app.core import config as _config

    monkeypatch.setattr(_config.settings, "COLABORADORES_TABLE", "tabela-que-nao-existe")

    resp = client_gestor.post("/api/colaboradores/sincronizar")

    assert resp.status_code == 503
    assert resp.json()["detail"] == MSG_INDISPONIVEL
    assert client_gestor.get("/api/colaboradores/sincronizacao").json()["ultima"]["status"] == "erro"
    assert client_gestor.get(f"/api/colaboradores/cpf/{CPF_ATIVO}").status_code == 503


@pytest.mark.integration
def test_sem_url_configurada_status_nao_configurado(client_gestor, monkeypatch):
    from app.core import config as _config

    monkeypatch.setattr(_config.settings, "COLABORADORES_DATABASE_URL", "")
    status = client_gestor.get("/api/colaboradores/sincronizacao").json()
    assert status == {"ultima": None, "configurado": False}
    assert client_gestor.post("/api/colaboradores/sincronizar").status_code == 503


@pytest.mark.integration
def test_base_vazia_nao_desliga_ninguem(client_gestor, colab_db, base_rh):
    _sincronizar(colab_db)
    base_rh('DELETE FROM "colaboradores-revalle"')

    resp = client_gestor.post("/api/colaboradores/sincronizar")

    assert resp.status_code == 503
    assert _por_cpf(colab_db, CPF_ATIVO)["situacao"] == "ativo"
    assert colab_db.ultima_sync()["status"] == "erro"


# ── Consulta por CPF ────────────────────────────────────────────────────────

@pytest.mark.integration
def test_cpf_lookup_mensagens(client_gestor, base_rh):
    resp = client_gestor.get("/api/colaboradores/cpf/123")
    assert resp.status_code == 422
    assert resp.json()["detail"] == "CPF inválido."
    assert client_gestor.get("/api/colaboradores/cpf/11111111111").status_code == 422

    resp = client_gestor.get(f"/api/colaboradores/cpf/{CPF_CARLOS}")
    assert resp.status_code == 404
    assert resp.json()["detail"] == "CPF não encontrado na base de colaboradores."


@pytest.mark.integration
def test_cpf_lookup_base_nao_grava_e_local_tem_prioridade(client_gestor, base_rh, criar_unidade):
    unidade = criar_unidade(nome="Revalle Juazeiro")

    resp = client_gestor.get("/api/colaboradores/cpf/529.982.247-25")
    assert resp.status_code == 200
    corpo = resp.json()
    assert corpo["local"] is False
    assert corpo["id"] is None
    assert corpo["nome"] == "Maria da Silva"
    assert corpo["cpf"] == CPF_ATIVO
    assert corpo["telefone"] == "74999990001"
    assert corpo["unidade_id"] == unidade["id"]
    assert client_gestor.get("/api/colaboradores").json()["total"] == 0

    client_gestor.post("/api/colaboradores/sincronizar")
    corpo = client_gestor.get(f"/api/colaboradores/cpf/{CPF_ATIVO}").json()
    assert corpo["local"] is True
    assert corpo["id"] is not None


@pytest.mark.integration
def test_cpf_lookup_local_nao_precisa_da_base(client_gestor, monkeypatch):
    from app.core import config as _config

    criado = client_gestor.post("/api/colaboradores", json={"cpf": CPF_CARLOS, "nome": "Carlos Souza"})
    assert criado.status_code == 201
    monkeypatch.setattr(_config.settings, "COLABORADORES_DATABASE_URL", "")

    assert client_gestor.get(f"/api/colaboradores/cpf/{CPF_CARLOS}").status_code == 200
    assert client_gestor.get(f"/api/colaboradores/cpf/{CPF_OUTRO}").status_code == 503


# ── Listagem ────────────────────────────────────────────────────────────────

@pytest.mark.integration
def test_listar_busca_por_cpf_com_mascara(client_gestor, base_rh):
    client_gestor.post("/api/colaboradores/sincronizar")

    for busca in ("529.982.247-25", "52998224725", "99822"):
        resp = client_gestor.get("/api/colaboradores", params={"busca": busca})
        assert resp.status_code == 200
        assert [i["nome"] for i in resp.json()["items"]] == ["Maria da Silva"], busca
    assert client_gestor.get("/api/colaboradores", params={"busca": "935.411.347-80"}).json()["total"] == 1


@pytest.mark.integration
def test_listar_filtros_e_paginacao(client_gestor, base_rh, criar_unidade):
    unidade = criar_unidade(nome="Revalle Petrolina")
    client_gestor.post("/api/colaboradores/sincronizar")

    todos = client_gestor.get("/api/colaboradores").json()
    assert todos["total"] == 4
    assert [i["nome"] for i in todos["items"]] == ["Ana Lima", "João Pereira", "Maria da Silva", "Pedro Alves"]
    primeiro = todos["items"][0]
    assert "gestor_nome_rh" not in primeiro
    assert primeiro["unidade_nome"] == "Revalle Petrolina"
    assert primeiro["sincronizado_em"].endswith("Z")

    pagina = client_gestor.get("/api/colaboradores", params={"limit": 2, "offset": 2}).json()
    assert pagina["total"] == 4
    assert [i["nome"] for i in pagina["items"]] == ["Maria da Silva", "Pedro Alves"]

    por_matricula = client_gestor.get("/api/colaboradores", params={"busca": "1003"}).json()
    assert [i["nome"] for i in por_matricula["items"]] == ["Ana Lima"]
    por_nome = client_gestor.get("/api/colaboradores", params={"busca": "joao"}).json()
    assert por_nome["total"] == 1
    desligados = client_gestor.get("/api/colaboradores", params={"situacao": "desligado"}).json()
    assert [i["nome"] for i in desligados["items"]] == ["João Pereira"]
    da_unidade = client_gestor.get("/api/colaboradores", params={"unidade_id": unidade["id"]}).json()
    assert da_unidade["total"] == 2
    assert client_gestor.get("/api/colaboradores", params={"situacao": "xyz"}).status_code == 422
    assert client_gestor.get("/api/colaboradores", params={"busca": "100%"}).json()["total"] == 0


@pytest.mark.integration
def test_leitura_exige_autenticacao_e_aceita_qualquer_papel(client, client_aprendiz):
    assert client.get("/api/colaboradores").status_code == 401
    assert client.get("/api/colaboradores/cargos").status_code == 401
    assert client.get("/api/colaboradores/sincronizacao").status_code == 401
    assert client_aprendiz.get("/api/colaboradores").status_code == 200
    assert client_aprendiz.get("/api/colaboradores/cargos").status_code == 200
    assert client_aprendiz.get("/api/colaboradores/999999").status_code == 404


@pytest.mark.integration
def test_cargos_distintos_e_ordenados(client_gestor, base_rh):
    client_gestor.post("/api/colaboradores/sincronizar")
    client_gestor.post("/api/colaboradores", json={"cpf": CPF_CARLOS, "nome": "Carlos", "cargo": "Analista"})
    client_gestor.post("/api/colaboradores", json={"cpf": CPF_OUTRO, "nome": "Sem Cargo"})

    assert client_gestor.get("/api/colaboradores/cargos").json() == [
        "Analista", "Auxiliar", "Gerente", "Vendedor",
    ]


# ── Criação e edição ────────────────────────────────────────────────────────

@pytest.mark.integration
def test_criar_manual(client_tecnico, client_gestor, client_aprendiz, criar_unidade):
    unidade = criar_unidade()
    corpo = {
        "cpf": "390.533.447-05",
        "nome": "  Carlos Souza ",
        "cargo": "Diretor",
        "unidade_id": unidade["id"],
        "email": "Carlos@Revalle.com.br",
        "telefone": "+55 (74) 99999-0000",
    }
    assert client_aprendiz.post("/api/colaboradores", json=corpo).status_code == 403

    resp = client_tecnico.post("/api/colaboradores", json=corpo)

    assert resp.status_code == 201
    criado = resp.json()
    assert criado["cpf"] == CPF_CARLOS
    assert criado["nome"] == "Carlos Souza"
    assert criado["origem"] == "manual"
    assert criado["situacao"] == "ativo"
    assert criado["email"] == "carlos@revalle.com.br"
    assert criado["telefone"] == "74999990000"
    assert criado["unidade_nome"] == unidade["nome"]
    assert criado["sincronizado_em"] is None

    dup = client_gestor.post("/api/colaboradores", json={**corpo, "cpf": CPF_CARLOS})
    assert dup.status_code == 409
    assert dup.json()["detail"] == "Já existe um colaborador com este CPF."

    linha = _auditoria("colaborador", "criar")[0]
    assert linha["entidade_id"] == str(criado["id"])
    assert linha["depois"]["cpf"] == "***.533.447-**"
    assert linha["depois"]["nome"] == "Carlos Souza"


@pytest.mark.integration
def test_criar_manual_validacoes(client_gestor):
    base = {"cpf": CPF_CARLOS, "nome": "Carlos"}
    assert client_gestor.post("/api/colaboradores", json={**base, "cpf": "11111111111"}).status_code == 422
    resp = client_gestor.post("/api/colaboradores", json={**base, "email": "invalido"})
    assert resp.status_code == 422
    assert resp.json()["detail"] == "E-mail: E-mail inválido."
    resp = client_gestor.post("/api/colaboradores", json={**base, "telefone": "123"})
    assert resp.json()["detail"] == "Telefone: Telefone inválido: informe DDD + número."
    assert client_gestor.post("/api/colaboradores", json={**base, "nome": "  "}).status_code == 422
    assert client_gestor.post("/api/colaboradores", json={**base, "situacao": "x"}).status_code == 422
    assert client_gestor.post("/api/colaboradores", json={**base, "unidade_id": 9999}).status_code == 422
    assert client_gestor.post("/api/colaboradores", json={**base, "superior_id": 9999}).status_code == 422
    assert client_gestor.get("/api/colaboradores").json()["total"] == 0


@pytest.mark.integration
def test_editar_manual_e_semantica_de_chaves(client_tecnico, client_gestor):
    criado = client_gestor.post(
        "/api/colaboradores",
        json={"cpf": CPF_CARLOS, "nome": "Carlos", "email": "c@x.com", "telefone": "74999990000"},
    ).json()
    chefe = client_gestor.post("/api/colaboradores", json={"cpf": CPF_OUTRO, "nome": "Chefe"}).json()
    url = f"/api/colaboradores/{criado['id']}"

    # Campos de origem manual são livres; chave ausente não muda nada.
    resp = client_tecnico.put(url, json={"nome": "Carlos Souza", "cargo": "Diretor", "superior_id": chefe["id"]})
    assert resp.status_code == 200
    corpo = resp.json()
    assert (corpo["nome"], corpo["cargo"], corpo["superior_nome"]) == ("Carlos Souza", "Diretor", "Chefe")
    assert corpo["email"] == "c@x.com"
    assert corpo["telefone"] == "74999990000"

    # null limpa; "" também.
    resp = client_gestor.put(url, json={"email": None, "superior_id": None, "cargo": ""})
    corpo = resp.json()
    assert (corpo["email"], corpo["superior_id"], corpo["cargo"]) == (None, None, None)
    assert corpo["email_manual"] is True

    linha = _auditoria("colaborador", "editar")[-1]
    assert linha["antes"]["email"] == "c@x.com"
    assert linha["depois"]["email"] is None
    assert "cpf" not in (linha["antes"] or {})

    assert client_gestor.put(url, json={}).status_code == 400
    assert client_gestor.put(url, json={"nome": None}).status_code == 422
    assert client_gestor.put(url, json={"situacao": None}).status_code == 422
    assert client_gestor.put(url, json={"email": "ruim"}).status_code == 422
    assert client_gestor.put("/api/colaboradores/999999", json={"cargo": "x"}).status_code == 404


@pytest.mark.integration
def test_flag_manual_so_quando_o_valor_muda(client_gestor, base_rh):
    client_gestor.post("/api/colaboradores/sincronizar")
    maria = client_gestor.get(f"/api/colaboradores/cpf/{CPF_ATIVO}").json()
    url = f"/api/colaboradores/{maria['id']}"

    corpo = client_gestor.put(url, json={"email": "MARIA@revalle.com.br", "telefone": "(74) 99999-0001"}).json()
    assert corpo["email_manual"] is False
    assert corpo["telefone_manual"] is False

    corpo = client_gestor.put(url, json={"telefone": "(74) 91111-2222"}).json()
    assert corpo["telefone"] == "74911112222"
    assert corpo["telefone_manual"] is True
    assert corpo["email_manual"] is False


@pytest.mark.integration
def test_editar_campo_da_base_recusado(client_gestor, base_rh):
    client_gestor.post("/api/colaboradores/sincronizar")
    maria = client_gestor.get(f"/api/colaboradores/cpf/{CPF_ATIVO}").json()
    url = f"/api/colaboradores/{maria['id']}"

    for corpo in ({"nome": "Outro Nome"}, {"cargo": "X"}, {"setor": "X"}, {"unidade_id": None}, {"situacao": "afastado"}):
        resp = client_gestor.put(url, json=corpo)
        assert resp.status_code == 422, corpo
        assert resp.json()["detail"] == "Campo vem da base de RH e não pode ser editado aqui."

    # E-mail, telefone e superior continuam editáveis.
    assert client_gestor.put(url, json={"email": "novo@revalle.com.br"}).status_code == 200
    assert client_gestor.get(url).json()["nome"] == "Maria da Silva"


@pytest.mark.integration
def test_superior_de_si_mesmo_recusado(client_gestor, base_rh):
    client_gestor.post("/api/colaboradores/sincronizar")
    maria = client_gestor.get(f"/api/colaboradores/cpf/{CPF_ATIVO}").json()

    resp = client_gestor.put(f"/api/colaboradores/{maria['id']}", json={"superior_id": maria["id"]})

    assert resp.status_code == 422
    assert resp.json()["detail"] == "O colaborador não pode ser superior de si mesmo."
    assert client_gestor.get(f"/api/colaboradores/{maria['id']}").json()["superior_id"] is None


@pytest.mark.integration
def test_superior_inexistente_recusado(client_gestor):
    criado = client_gestor.post("/api/colaboradores", json={"cpf": CPF_CARLOS, "nome": "Carlos"}).json()
    resp = client_gestor.put(f"/api/colaboradores/{criado['id']}", json={"superior_id": 9999})
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Superior não encontrado."


@pytest.mark.integration
def test_escrita_exige_gestor_ou_tecnico(client_aprendiz, client_gestor):
    criado = client_gestor.post("/api/colaboradores", json={"cpf": CPF_CARLOS, "nome": "Carlos"}).json()
    assert client_aprendiz.put(f"/api/colaboradores/{criado['id']}", json={"cargo": "x"}).status_code == 403


@pytest.mark.integration
def test_manual_com_cpf_que_entra_na_base_vira_rh(colab_db, base_rh):
    manual_id = colab_db.criar_manual(
        {"cpf": CPF_ATIVO, "nome": "Maria (cadastro manual)", "email": "manual@x.com"}
    )

    resumo = _sincronizar(colab_db)

    assert resumo.criados == 3 and resumo.atualizados == 1
    maria = colab_db.obter(manual_id)
    assert maria["origem"] == "rh"
    assert maria["nome"] == "Maria da Silva"
    assert maria["email"] == "manual@x.com"  # preenchido à mão: preservado


# ── Unidades: alias_rh ──────────────────────────────────────────────────────

@pytest.mark.integration
def test_unidade_alias_rh_criar_editar_e_limpar(client_gestor):
    corpo = {
        "nome": "Unidade Alias",
        "razao_social": "REVENDA LTDA",
        "cnpj": "04.690.106/0001-15",
        "alias_rh": "  Revalle Juazeiro ",
    }
    resp = client_gestor.post("/api/unidades", json=corpo)
    assert resp.status_code == 200
    unidade_id = resp.json()["id"]
    assert client_gestor.get(f"/api/unidades/{unidade_id}").json()["alias_rh"] == "Revalle Juazeiro"

    # Outro campo: o alias não muda.
    assert client_gestor.put(f"/api/unidades/{unidade_id}", json={"cidade": "Juazeiro"}).status_code == 200
    assert client_gestor.get(f"/api/unidades/{unidade_id}").json()["alias_rh"] == "Revalle Juazeiro"

    resp = client_gestor.put(f"/api/unidades/{unidade_id}", json={"alias_rh": "Outro Nome"})
    assert resp.status_code == 200
    assert client_gestor.get(f"/api/unidades/{unidade_id}").json()["alias_rh"] == "Outro Nome"

    # "" limpa; null também; e só isso já é uma atualização válida.
    assert client_gestor.put(f"/api/unidades/{unidade_id}", json={"alias_rh": ""}).status_code == 200
    assert client_gestor.get(f"/api/unidades/{unidade_id}").json()["alias_rh"] is None

    client_gestor.put(f"/api/unidades/{unidade_id}", json={"alias_rh": "Voltou"})
    assert client_gestor.put(f"/api/unidades/{unidade_id}", json={"alias_rh": None}).status_code == 200
    assert client_gestor.get(f"/api/unidades/{unidade_id}").json()["alias_rh"] is None
    assert client_gestor.get("/api/unidades").json()[0]["alias_rh"] is None
