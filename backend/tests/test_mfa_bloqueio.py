"""F2 rodada 1: limite de falhas na 2ª etapa do login (por token e por usuário) e
transições atômicas do 2FA."""
import json
from datetime import datetime, timedelta, timezone

import pytest

from tests.fixtures_auth import (  # noqa: F401
    ativar_2fa_direto,
    chave_totp,
    codigo_totp,
    ida_e_volta_migracao,
    linhas_auditoria,
)

pytestmark = pytest.mark.integration

MSG_ETAPA_EXPIRADA = "Etapa de verificação expirada. Entre novamente."
MSG_TOKEN_ESGOTADO = "Muitas tentativas. Entre novamente."
MSG_BLOQUEADO = (
    "Verificação em duas etapas bloqueada por excesso de tentativas. Tente novamente em 15 minutos."
)


def _login(client, username="tecnico.teste", senha="SenhaForte#123"):
    return client.post("/api/auth/login", json={"username": username, "password": senha})


def _novo_token(client, username="tecnico.teste"):
    resp = _login(client, username)
    assert resp.json()["etapa"] == "mfa", resp.text
    return resp.json()["mfa_token"]


def _verificar(client, mfa_token, codigo):
    return client.post("/api/auth/mfa/verificar", json={"mfa_token": mfa_token, "codigo": codigo})


def _errado(segredo):
    validos = {codigo_totp(segredo, p) for p in (-1, 0, 1)}
    return next(c for c in ("000000", "111111", "222222") if c not in validos)


def _linhas(sql, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


@pytest.fixture
def tecnico_2fa(usuario_tecnico, ativar_2fa_direto):
    return usuario_tecnico, ativar_2fa_direto(usuario_tecnico["id"])


def test_migracao_0203_ida_e_volta(_test_schema):
    ida_e_volta_migracao(203)
    colunas = {l["nome"] for l in _linhas(
        "SELECT column_name AS nome FROM information_schema.columns "
        "WHERE table_schema = DATABASE() AND table_name = 'usuarios'"
    )}
    assert "mfa_bloqueado_ate" in colunas
    assert _linhas(
        "SELECT 1 FROM information_schema.tables WHERE table_schema = DATABASE() AND table_name = 'mfa_falhas'"
    )


def test_token_morre_apos_5_codigos_errados(client, tecnico_2fa):
    _, dados = tecnico_2fa
    errado = _errado(dados.segredo)
    mfa_token = _novo_token(client)

    for _ in range(4):
        resp = _verificar(client, mfa_token, errado)
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Código inválido."
    quinta = _verificar(client, mfa_token, errado)
    assert quinta.status_code == 400
    assert quinta.json()["detail"] == MSG_TOKEN_ESGOTADO

    # Mesmo o código certo não serve mais com esse token: é preciso entrar de novo.
    resp = _verificar(client, mfa_token, codigo_totp(dados.segredo))
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA
    assert _verificar(client, _novo_token(client), codigo_totp(dados.segredo)).status_code == 200


def test_cada_falha_e_auditada_sem_codigo_nem_token(client, tecnico_2fa):
    _, dados = tecnico_2fa
    errado = _errado(dados.segredo)
    mfa_token = _novo_token(client)
    for _ in range(3):
        _verificar(client, mfa_token, errado)

    linhas = linhas_auditoria("usuario", "falha_2fa")
    assert len(linhas) == 3
    assert [l["depois"]["falhas_no_token"] for l in linhas] == [1, 2, 3]
    assert all(l["username"] == "tecnico.teste" for l in linhas)
    texto = json.dumps(linhas, default=str)
    assert errado not in texto and mfa_token not in texto


def test_usuario_bloqueado_apos_10_falhas_entre_tokens(
    client, tecnico_2fa, usuario_aprendiz, executar_sql
):
    usuario, dados = tecnico_2fa
    errado = _errado(dados.segredo)

    primeiro = _novo_token(client)
    for _ in range(5):
        _verificar(client, primeiro, errado)

    segundo = _novo_token(client)
    for _ in range(4):
        assert _verificar(client, segundo, errado).json()["detail"] == "Código inválido."
    decima = _verificar(client, segundo, errado)
    assert decima.status_code == 429
    assert decima.json()["detail"] == MSG_BLOQUEADO

    # Bloqueado: nem o código certo nem um código de recuperação entram, com token novo.
    terceiro = _novo_token(client)
    certo = _verificar(client, terceiro, codigo_totp(dados.segredo))
    assert certo.status_code == 429
    assert certo.json()["detail"] == MSG_BLOQUEADO
    assert _verificar(client, terceiro, dados.codigos[0]).status_code == 429

    # Terminado o bloqueio, volta a funcionar.
    executar_sql(
        "UPDATE usuarios SET mfa_bloqueado_ate = %s WHERE id = %s",
        (datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(minutes=1), usuario["id"]),
    )
    liberado = _verificar(client, terceiro, codigo_totp(dados.segredo))
    assert liberado.status_code == 200, liberado.text
    assert _linhas("SELECT mfa_bloqueado_ate FROM usuarios WHERE id = %s", (usuario["id"],))[0][
        "mfa_bloqueado_ate"
    ] is None


def test_bloqueio_nao_afeta_outro_usuario(client, tecnico_2fa, ativar_2fa_direto, criar_usuario, senha_padrao_teste):
    _, dados = tecnico_2fa
    errado = _errado(dados.segredo)
    for _ in range(2):
        token = _novo_token(client)
        for _ in range(5):
            _verificar(client, token, errado)

    outro = criar_usuario("outro.tec", senha_padrao_teste, "Técnico")
    dados_outro = ativar_2fa_direto(outro["id"])
    token = _novo_token(client, "outro.tec")
    assert _verificar(client, token, codigo_totp(dados_outro.segredo)).status_code == 200


def test_sucesso_zera_contador_e_bloqueio(client, tecnico_2fa):
    usuario, dados = tecnico_2fa
    errado = _errado(dados.segredo)

    primeiro = _novo_token(client)
    for _ in range(3):
        _verificar(client, primeiro, errado)
    assert _verificar(client, primeiro, codigo_totp(dados.segredo)).status_code == 200

    assert _linhas("SELECT 1 FROM mfa_falhas WHERE usuario_id = %s", (usuario["id"],)) == []
    # Contador zerado: um token novo volta a ter as 4 tentativas antes de morrer.
    segundo = _novo_token(client)
    for _ in range(4):
        assert _verificar(client, segundo, errado).json()["detail"] == "Código inválido."


def test_iniciar_depois_da_ativacao_nao_desativa_o_2fa(tecnico_2fa):
    from app.db.mfa_db import MfaDBManager
    from app.services.mfa import MfaJaAtivo
    from app.services.segundo_fator import iniciar_cadastro

    usuario, _ = tecnico_2fa
    db = MfaDBManager()
    antes = db.estado(usuario["id"])

    assert db.salvar_segredo_pendente(usuario["id"], "outro-segredo") is False
    # Leitura velha (totp_ativo falso) de quem perdeu a corrida: a UPDATE condicional barra.
    with pytest.raises(MfaJaAtivo):
        iniciar_cadastro(db, {"id": usuario["id"], "username": usuario["username"], "totp_ativo": 0})

    depois = db.estado(usuario["id"])
    assert depois["totp_ativo"] == 1
    assert depois["totp_secret_cifrado"] == antes["totp_secret_cifrado"]


def test_ativar_so_ativa_o_segredo_verificado(client, chave_totp, usuario_tecnico):
    from app.db.mfa_db import MfaDBManager

    db = MfaDBManager()
    uid = usuario_tecnico["id"]
    assert db.salvar_segredo_pendente(uid, "segredo-b") is True

    # Verificou o segredo A, mas um "iniciar" concorrente já gravou o B.
    assert db.ativar(uid, ["h" * 64], passo=1, segredo_cifrado="segredo-a") is False
    assert db.estado(uid)["totp_ativo"] == 0
    assert _linhas("SELECT 1 FROM mfa_recovery_codes WHERE usuario_id = %s", (uid,)) == []

    assert db.ativar(uid, ["h" * 64], passo=1, segredo_cifrado="segredo-b") is True
    assert db.estado(uid)["totp_ativo"] == 1
    # Segunda ativação não repete nem troca códigos.
    assert db.ativar(uid, ["i" * 64], passo=2, segredo_cifrado="segredo-b") is False
    assert [l["code_hash"] for l in _linhas("SELECT code_hash FROM mfa_recovery_codes")] == ["h" * 64]


def test_iniciar_pelo_perfil_com_2fa_ativo_continua_ativo(
    _fastapi_app, client, tecnico_2fa, senha_padrao_teste
):
    from fastapi.testclient import TestClient

    usuario, dados = tecnico_2fa
    mfa_token = _novo_token(client)
    access = _verificar(client, mfa_token, codigo_totp(dados.segredo)).json()["access_token"]
    c = TestClient(_fastapi_app)
    c.headers.update({"Authorization": f"Bearer {access}"})

    resp = c.post("/api/perfil/2fa/iniciar", json={"senha": senha_padrao_teste})
    assert resp.status_code == 409
    assert _linhas("SELECT totp_ativo FROM usuarios WHERE id = %s", (usuario["id"],))[0]["totp_ativo"] == 1


# ── Concorrência (rodada 2) ─────────────────────────────────────────────────

def _em_paralelo(funcoes):
    """Roda as funções ao mesmo tempo (largada sincronizada) e devolve os resultados em ordem."""
    import threading
    from concurrent.futures import ThreadPoolExecutor

    largada = threading.Barrier(len(funcoes))

    def _rodar(f):
        largada.wait()
        return f()

    with ThreadPoolExecutor(max_workers=len(funcoes)) as ex:
        return list(ex.map(_rodar, funcoes))


def test_codigos_errados_simultaneos_respeitam_os_limites(_fastapi_app, client, tecnico_2fa, monkeypatch):
    """12 códigos errados ao mesmo tempo, 6 em cada um de 2 tokens. Invariantes (não dependem
    da ordem de chegada): nenhum 500; toda resposta é uma das mensagens previstas; no máximo
    5 códigos verificados por token; 10 falhas gravadas no total => bloqueio do usuário; tokens revogados."""
    from fastapi.testclient import TestClient

    import app.routers.auth as auth_router

    usuario, dados = tecnico_2fa
    errado = _errado(dados.segredo)
    tokens = [_novo_token(client), _novo_token(client)]

    verificados = []
    original = auth_router.verificar_segundo_fator

    def _espiao(db, usuario_, codigo, ctx):
        verificados.append(codigo)
        return original(db, usuario_, codigo, ctx)

    monkeypatch.setattr(auth_router, "verificar_segundo_fator", _espiao)

    def tentar(token):
        return lambda: TestClient(_fastapi_app).post(
            "/api/auth/mfa/verificar", json={"mfa_token": token, "codigo": errado}
        )

    respostas = _em_paralelo([tentar(t) for t in tokens for _ in range(6)])

    esperadas = {"Código inválido.", MSG_TOKEN_ESGOTADO, MSG_BLOQUEADO, MSG_ETAPA_EXPIRADA}
    assert all(r.status_code in (400, 429) for r in respostas), [r.status_code for r in respostas]
    assert {r.json()["detail"] for r in respostas} <= esperadas

    linhas = _linhas("SELECT jti FROM mfa_falhas WHERE usuario_id = %s", (usuario["id"],))
    por_token = {}
    for l in linhas:
        por_token[l["jti"]] = por_token.get(l["jti"], 0) + 1
    assert all(n <= 5 for n in por_token.values()), por_token  # limite estrito por token
    assert len(verificados) == len(linhas) <= 10  # só se verifica o que cabe no limite
    assert len(linhas) == 10  # 2 tokens x 5: o usuário chegou a 10 falhas gravadas
    assert len(linhas_auditoria("usuario", "falha_2fa")) == 10

    bloqueio = _linhas("SELECT mfa_bloqueado_ate FROM usuarios WHERE id = %s", (usuario["id"],))[0]
    assert bloqueio["mfa_bloqueado_ate"] is not None
    for token in tokens:  # revogados: nem o código certo entra
        assert _verificar(client, token, codigo_totp(dados.segredo)).json()["detail"] == MSG_ETAPA_EXPIRADA
    assert _verificar(client, _novo_token(client), codigo_totp(dados.segredo)).status_code == 429


def test_token_com_5_falhas_gravadas_nao_verifica_nem_o_codigo_certo(client, tecnico_2fa):
    """Limite estrito: sob a trava do usuário, o token que já tem 5 falhas é recusado sem
    conferir o código (cobre o pedido em voo que chegou antes da revogação)."""
    from jose import jwt as jose_jwt

    from app.db.mfa_db import MfaDBManager

    usuario, dados = tecnico_2fa
    mfa_token = _novo_token(client)
    jti = jose_jwt.get_unverified_claims(mfa_token)["jti"]
    db = MfaDBManager()
    agora = datetime.now(timezone.utc).replace(tzinfo=None)
    for _ in range(5):
        db.registrar_falha(usuario["id"], jti, agora, agora - timedelta(minutes=15))

    resp = _verificar(client, mfa_token, codigo_totp(dados.segredo))
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_TOKEN_ESGOTADO
    # ...e o token fica revogado.
    assert _verificar(client, mfa_token, codigo_totp(dados.segredo)).json()["detail"] == MSG_ETAPA_EXPIRADA
    assert len(_linhas("SELECT 1 FROM mfa_falhas WHERE usuario_id = %s", (usuario["id"],))) == 5


def test_substituir_codigos_e_consumo_simultaneos_nao_dao_deadlock(tecnico_2fa):
    """Todo caminho trava `usuarios` primeiro: trocar os códigos (perfil) ao mesmo tempo que
    o login consome um código (transação do usuário) não pode gerar espera cruzada pela FK."""
    from app.db.mfa_db import MfaDBManager
    from app.services.mfa import hash_codigo_recuperacao

    usuario, dados = tecnico_2fa
    uid = usuario["id"]
    db = MfaDBManager()
    codigo = hash_codigo_recuperacao(dados.codigos[0])

    def trocar():
        db.substituir_codigos(uid, [hash_codigo_recuperacao(c) for c in dados.codigos])
        return "trocou"

    def consumir():
        with db.transacao_do_usuario(uid):
            db.consumir_codigo(uid, codigo)
            db.contar_codigos_restantes(uid)
        return "consumiu"

    for _ in range(5):
        resultados = _em_paralelo([trocar, consumir, trocar, consumir, trocar, consumir])
        assert sorted(set(resultados)) == ["consumiu", "trocou"]


def test_falhas_simultaneas_de_usuarios_diferentes_sao_ambas_gravadas(
    _fastapi_app, client, tecnico_2fa, ativar_2fa_direto, criar_usuario, senha_padrao_teste
):
    from fastapi.testclient import TestClient

    _, dados_a = tecnico_2fa
    outro = criar_usuario("outro.tec", senha_padrao_teste, "Técnico")
    dados_b = ativar_2fa_direto(outro["id"])

    for _rodada in range(4):
        tokens = {"tecnico.teste": _novo_token(client), "outro.tec": _novo_token(client, "outro.tec")}
        erros = {"tecnico.teste": _errado(dados_a.segredo), "outro.tec": _errado(dados_b.segredo)}

        def tentar(nome):
            return lambda: TestClient(_fastapi_app).post(
                "/api/auth/mfa/verificar", json={"mfa_token": tokens[nome], "codigo": erros[nome]}
            )

        antes = len(_linhas("SELECT 1 FROM mfa_falhas"))
        respostas = _em_paralelo([tentar("tecnico.teste"), tentar("outro.tec")])
        assert [r.status_code for r in respostas] == [400, 400]
        assert len(_linhas("SELECT 1 FROM mfa_falhas")) == antes + 2


def test_registrar_falha_concorrente_nao_da_deadlock(tecnico_2fa, usuario_aprendiz):
    """No nível do banco: 10 falhas do mesmo usuário + 1 de outro, todas ao mesmo tempo."""
    from app.db.mfa_db import MfaDBManager

    usuario, _ = tecnico_2fa
    db = MfaDBManager()
    agora = datetime.now(timezone.utc).replace(tzinfo=None)
    janela = agora - timedelta(minutes=15)

    funcoes = [lambda: db.registrar_falha(usuario["id"], "jti-x", agora, janela) for _ in range(10)]
    funcoes.append(lambda: db.registrar_falha(usuario_aprendiz["id"], "jti-y", agora, janela))
    resultados = _em_paralelo(funcoes)

    do_usuario = sorted(r[1] for r in resultados[:10])
    assert do_usuario == list(range(1, 11))  # serializadas: cada uma viu as anteriores
    assert resultados[10] == (1, 1)
    assert len(_linhas("SELECT 1 FROM mfa_falhas")) == 11


def test_deadlock_na_verificacao_vira_409_nunca_500(client, tecnico_2fa, monkeypatch):
    import pymysql

    from app.db.mfa_db import MfaDBManager

    _, dados = tecnico_2fa
    mfa_token = _novo_token(client)

    def _estoura(self, *a, **k):
        raise pymysql.err.OperationalError(1213, "Deadlock found when trying to get lock")

    monkeypatch.setattr(MfaDBManager, "registrar_passo", _estoura)
    resp = _verificar(client, mfa_token, codigo_totp(dados.segredo))
    assert resp.status_code == 409
    assert resp.json()["detail"] == "Não foi possível verificar agora. Tente novamente."
