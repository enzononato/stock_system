"""F2 Task 6: cadastro obrigatório do autenticador para o Gestor e regra da chave ausente."""
import json
import logging
import time

import pytest
from cryptography.fernet import Fernet
from jose import jwt as jose_jwt

from tests.fixtures_auth import (  # noqa: F401
    ativar_2fa_direto,
    chave_totp,
    codigo_totp,
    linhas_auditoria,
)

pytestmark = pytest.mark.integration

MSG_ETAPA_EXPIRADA = "Etapa de verificação expirada. Entre novamente."
MSG_CADASTRO_OBRIGATORIO = "Cadastre a verificação em duas etapas para continuar. Entre novamente."


def _login(client, username="gestor.teste", senha="SenhaForte#123"):
    return client.post("/api/auth/login", json={"username": username, "password": senha})


def _iniciar(client, mfa_token):
    return client.post("/api/auth/mfa/cadastro/iniciar", json={"mfa_token": mfa_token})


def _confirmar(client, mfa_token, codigo):
    return client.post("/api/auth/mfa/cadastro/confirmar", json={"mfa_token": mfa_token, "codigo": codigo})


def _linhas(sql, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


def _chave_para(monkeypatch, valor):
    from app.core import config

    monkeypatch.setattr(config.settings, "TOTP_ENCRYPTION_KEY", valor)


def test_gestor_sem_2fa_recebe_etapa_cadastro(client, chave_totp, usuario_gestor):
    resp = _login(client)
    assert resp.status_code == 200
    corpo = resp.json()
    assert corpo["etapa"] == "cadastro_2fa"
    assert corpo["access_token"] is None
    assert "refresh_token" not in resp.cookies
    assert corpo["expira_em_segundos"] == 900
    claims = jose_jwt.get_unverified_claims(corpo["mfa_token"])
    assert claims["proposito"] == "cadastrar"
    assert claims["type"] == "mfa_pendente"
    assert time.time() + 895 <= claims["exp"] <= time.time() + 905


def test_token_de_cadastro_nao_acessa_a_api(client, chave_totp, usuario_gestor):
    mfa_token = _login(client).json()["mfa_token"]
    resp = client.get("/api/items", headers={"Authorization": f"Bearer {mfa_token}"})
    assert resp.status_code == 401


def test_cadastro_completo_emite_sessao_e_codigos(client, chave_totp, usuario_gestor):
    from app.core.security import sha256_hex
    from app.services.mfa import decifrar_segredo

    mfa_token = _login(client).json()["mfa_token"]
    ini = _iniciar(client, mfa_token)
    assert ini.status_code == 200, ini.text
    segredo = ini.json()["segredo"]
    uri = ini.json()["otpauth_uri"]
    assert uri.startswith("otpauth://totp/")
    assert f"secret={segredo}" in uri

    linha = _linhas("SELECT totp_secret_cifrado, totp_ativo FROM usuarios WHERE id = %s", (usuario_gestor["id"],))[0]
    assert linha["totp_secret_cifrado"] != segredo
    assert decifrar_segredo(linha["totp_secret_cifrado"]) == segredo
    assert linha["totp_ativo"] == 0

    conf = _confirmar(client, mfa_token, codigo_totp(segredo))
    assert conf.status_code == 200, conf.text
    corpo = conf.json()
    assert corpo["etapa"] == "ok"
    assert "refresh_token" in conf.cookies
    codigos = corpo["codigos_recuperacao"]
    assert len(codigos) == 10 and len(set(codigos)) == 10
    assert client.get("/api/items", headers={"Authorization": f"Bearer {corpo['access_token']}"}).status_code == 200

    hashes = [l["code_hash"] for l in _linhas("SELECT code_hash FROM mfa_recovery_codes")]
    assert len(hashes) == 10
    assert not set(hashes) & set(codigos)
    assert set(hashes) == {sha256_hex(c.replace("-", "")) for c in codigos}

    auditoria = linhas_auditoria("usuario", "ativar_2fa")
    assert len(auditoria) == 1
    assert auditoria[0]["antes"] == {"totp_ativo": False}
    assert auditoria[0]["depois"] == {"totp_ativo": True}
    texto = json.dumps(auditoria, default=str)
    assert segredo not in texto
    assert not any(c in texto for c in codigos)


def test_confirmar_com_codigo_errado(client, chave_totp, usuario_gestor):
    mfa_token = _login(client).json()["mfa_token"]
    segredo = _iniciar(client, mfa_token).json()["segredo"]
    certo = codigo_totp(segredo)
    errado = "000000" if certo != "000000" else "111111"
    resp = _confirmar(client, mfa_token, errado)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Código inválido."
    assert _linhas("SELECT totp_ativo FROM usuarios WHERE id = %s", (usuario_gestor["id"],))[0]["totp_ativo"] == 0


def test_confirmar_sem_iniciar(client, chave_totp, usuario_gestor):
    mfa_token = _login(client).json()["mfa_token"]
    resp = _confirmar(client, mfa_token, "123456")
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Inicie o cadastro do autenticador antes de confirmar."


def test_token_de_cadastro_reutilizado_apos_confirmar(client, chave_totp, usuario_gestor):
    mfa_token = _login(client).json()["mfa_token"]
    segredo = _iniciar(client, mfa_token).json()["segredo"]
    assert _confirmar(client, mfa_token, codigo_totp(segredo)).status_code == 200
    resp = _iniciar(client, mfa_token)
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA


def test_token_de_verificacao_nao_serve_para_cadastro(client, chave_totp, usuario_gestor):
    from app.core.security import create_mfa_token
    from app.services.sessao import dados_do_token

    token = create_mfa_token(dados_do_token(usuario_gestor), "verificar")
    resp = _iniciar(client, token)
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA


def test_gestor_sem_2fa_com_sessao_antiga_e_bloqueado(client, chave_totp, usuario_gestor):
    from app.core.security import create_access_token, create_refresh_token

    dados = {"sub": str(usuario_gestor["id"]), "username": "gestor.teste", "role": "Gestor"}
    resp = client.get("/api/items", headers={"Authorization": f"Bearer {create_access_token(dados)}"})
    assert resp.status_code == 401
    assert resp.json()["detail"] == MSG_CADASTRO_OBRIGATORIO

    client.cookies.set("refresh_token", create_refresh_token(dados))
    ref = client.post("/api/auth/refresh")
    assert ref.status_code == 401
    assert ref.json()["detail"] == MSG_CADASTRO_OBRIGATORIO


def test_tecnico_sem_2fa_nao_e_obrigado(client, chave_totp, usuario_tecnico, senha_padrao_teste):
    resp = _login(client, "tecnico.teste", senha_padrao_teste)
    assert resp.json()["etapa"] == "ok"


def test_chave_vazia_gestor_entra_sem_2fa(client, usuario_gestor):
    resp = _login(client)
    assert resp.status_code == 200
    assert resp.json()["etapa"] == "ok"
    token = resp.json()["access_token"]
    assert client.get("/api/items", headers={"Authorization": f"Bearer {token}"}).status_code == 200


def test_chave_vazia_endpoints_de_cadastro_respondem_503(client, chave_totp, usuario_gestor, monkeypatch):
    from app.services.mfa import MSG_MFA_INDISPONIVEL

    mfa_token = _login(client).json()["mfa_token"]
    _chave_para(monkeypatch, "")
    resp = _iniciar(client, mfa_token)
    assert resp.status_code == 503
    assert resp.json()["detail"] == MSG_MFA_INDISPONIVEL


def test_chave_vazia_usuario_com_2fa_entra_com_codigo_de_recuperacao(
    client, usuario_gestor, senha_padrao_teste, ativar_2fa_direto, monkeypatch
):
    from app.services.mfa import MSG_MFA_INDISPONIVEL

    dados = ativar_2fa_direto(usuario_gestor["id"])
    _chave_para(monkeypatch, "")

    login = _login(client)
    assert login.json()["etapa"] == "mfa"
    mfa_token = login.json()["mfa_token"]

    totp = client.post(
        "/api/auth/mfa/verificar", json={"mfa_token": mfa_token, "codigo": codigo_totp(dados.segredo)}
    )
    assert totp.status_code == 503
    assert totp.json()["detail"] == MSG_MFA_INDISPONIVEL

    rec = client.post("/api/auth/mfa/verificar", json={"mfa_token": mfa_token, "codigo": dados.codigos[0]})
    assert rec.status_code == 200, rec.text
    assert rec.json()["etapa"] == "ok"


def test_chave_trocada_responde_segredo_ilegivel(client, usuario_gestor, ativar_2fa_direto, monkeypatch):
    from app.services.mfa import MSG_SEGREDO_ILEGIVEL

    dados = ativar_2fa_direto(usuario_gestor["id"])
    _chave_para(monkeypatch, Fernet.generate_key().decode())

    mfa_token = _login(client).json()["mfa_token"]
    resp = client.post(
        "/api/auth/mfa/verificar", json={"mfa_token": mfa_token, "codigo": codigo_totp(dados.segredo)}
    )
    assert resp.status_code == 503
    assert resp.json()["detail"] == MSG_SEGREDO_ILEGIVEL


def test_aviso_de_chave_ausente_no_boot(caplog, monkeypatch):
    import app.main as main_module

    _chave_para(monkeypatch, "")
    with caplog.at_level(logging.WARNING):
        main_module.avisar_2fa_sem_chave()
    assert "TOTP_ENCRYPTION_KEY ausente ou inválida" in caplog.text

    caplog.clear()
    _chave_para(monkeypatch, Fernet.generate_key().decode())
    with caplog.at_level(logging.WARNING):
        main_module.avisar_2fa_sem_chave()
    assert "TOTP_ENCRYPTION_KEY" not in caplog.text
