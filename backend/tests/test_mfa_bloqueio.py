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
