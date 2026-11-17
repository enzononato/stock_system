"""F2 Task 5: login em duas etapas (senha -> código TOTP ou de recuperação)."""
import json
import time
from datetime import datetime, timedelta, timezone

import pytest
from jose import jwt as jose_jwt

from tests.fixtures_auth import (  # noqa: F401
    email_de_teste,
    ativar_2fa_direto,
    chave_totp,
    cliente_com_rate_limit,
    codigo_totp,
    ida_e_volta_migracao,
    linhas_auditoria,
)

pytestmark = pytest.mark.integration

MSG_ETAPA_EXPIRADA = "Etapa de verificação expirada. Entre novamente."


def _login(client, username, senha):
    return client.post("/api/auth/login", json={"username": email_de_teste(username), "password": senha})


def _verificar(client, mfa_token, codigo):
    return client.post("/api/auth/mfa/verificar", json={"mfa_token": mfa_token, "codigo": codigo})


@pytest.fixture
def tecnico_2fa(usuario_tecnico, ativar_2fa_direto):
    return usuario_tecnico, ativar_2fa_direto(usuario_tecnico["id"])


def test_migracao_0202_ida_e_volta(_test_schema):
    from app.db.database_mysql import get_cursor

    ida_e_volta_migracao(202)
    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT COUNT(*) AS n FROM information_schema.tables "
            "WHERE table_schema = DATABASE() AND table_name = 'mfa_recovery_codes'"
        )
        assert cur.fetchone()["n"] == 1


def test_login_com_2fa_pede_codigo(client, tecnico_2fa, senha_padrao_teste):
    resp = _login(client, "tecnico.teste", senha_padrao_teste)
    assert resp.status_code == 200
    corpo = resp.json()
    assert corpo["etapa"] == "mfa"
    assert corpo["access_token"] is None
    assert "refresh_token" not in resp.cookies
    assert corpo["expira_em_segundos"] == 300

    claims = jose_jwt.get_unverified_claims(corpo["mfa_token"])
    assert claims["type"] == "mfa_pendente"
    assert claims["proposito"] == "verificar"
    agora = time.time()
    assert agora + 295 <= claims["exp"] <= agora + 305


def test_mfa_token_nao_serve_como_access(client, tecnico_2fa, senha_padrao_teste):
    mfa_token = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    resp = client.get("/api/auth/me", headers={"Authorization": f"Bearer {mfa_token}"})
    assert resp.status_code == 401


def test_verificar_com_totp_emite_sessao(client, tecnico_2fa, senha_padrao_teste):
    _, dados = tecnico_2fa
    mfa_token = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    resp = _verificar(client, mfa_token, codigo_totp(dados.segredo))
    assert resp.status_code == 200, resp.text
    corpo = resp.json()
    assert corpo["etapa"] == "ok"
    assert corpo["codigos_recuperacao_restantes"] is None
    assert "refresh_token" in resp.cookies
    me = client.get("/api/auth/me", headers={"Authorization": f"Bearer {corpo['access_token']}"})
    assert me.status_code == 200
    assert me.json()["username"] == "tecnico.teste"


def test_verificar_codigo_errado(client, tecnico_2fa, senha_padrao_teste):
    _, dados = tecnico_2fa
    certo = codigo_totp(dados.segredo)
    errado = "000000" if certo != "000000" else "111111"
    mfa_token = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    resp = _verificar(client, mfa_token, errado)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Código inválido."


def test_verificar_codigo_reutilizado_recusado(client, tecnico_2fa, senha_padrao_teste):
    _, dados = tecnico_2fa
    c = codigo_totp(dados.segredo)
    primeiro = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    assert _verificar(client, primeiro, c).status_code == 200

    segundo = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    resp = _verificar(client, segundo, c)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Código inválido."


def test_verificar_aceita_codigo_com_espaco(client, tecnico_2fa, senha_padrao_teste):
    _, dados = tecnico_2fa
    c = codigo_totp(dados.segredo)
    mfa_token = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    assert _verificar(client, mfa_token, f"{c[:3]} {c[3:]}").status_code == 200


def test_verificar_com_codigo_de_recuperacao(client, tecnico_2fa, senha_padrao_teste):
    _, dados = tecnico_2fa
    codigo = dados.codigos[0]
    mfa_token = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    resp = _verificar(client, mfa_token, codigo.upper().replace("-", ""))
    assert resp.status_code == 200, resp.text
    assert resp.json()["etapa"] == "ok"
    assert resp.json()["codigos_recuperacao_restantes"] == 9

    linhas = linhas_auditoria("usuario", "usar_codigo_recuperacao")
    assert len(linhas) == 1
    assert linhas[0]["depois"] == {"codigos_restantes": 9}
    texto = json.dumps(linhas, default=str)
    assert codigo not in texto
    assert codigo.replace("-", "") not in texto

    de_novo = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    assert _verificar(client, de_novo, codigo).status_code == 400


def test_mfa_token_usado_nao_vale_de_novo(client, tecnico_2fa, senha_padrao_teste):
    _, dados = tecnico_2fa
    mfa_token = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    assert _verificar(client, mfa_token, codigo_totp(dados.segredo)).status_code == 200
    resp = _verificar(client, mfa_token, codigo_totp(dados.segredo, passos=1))
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA


def test_mfa_token_expirado(client, tecnico_2fa):
    from app.core import config

    usuario, _ = tecnico_2fa
    expirado = jose_jwt.encode(
        {
            "sub": str(usuario["id"]), "username": usuario["username"], "role": usuario["role"], "sv": 0,
            "type": "mfa_pendente", "proposito": "verificar", "jti": "expirado-1",
            "exp": datetime.now(timezone.utc) - timedelta(minutes=1),
        },
        config.settings.JWT_SECRET,
        algorithm=config.settings.JWT_ALGORITHM,
    )
    resp = _verificar(client, expirado, "123456")
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA


def test_access_token_nao_serve_como_mfa_token(client, usuario_tecnico, senha_padrao_teste):
    access = _login(client, "tecnico.teste", senha_padrao_teste).json()["access_token"]
    resp = _verificar(client, access, "123456")
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA


def test_sessoes_revogadas_entre_as_etapas_invalidam_o_mfa_token(client, tecnico_2fa, senha_padrao_teste):
    from app.db.user_manager_db import UserDBManager

    usuario, dados = tecnico_2fa
    mfa_token = _login(client, "tecnico.teste", senha_padrao_teste).json()["mfa_token"]
    UserDBManager().revogar_sessoes(usuario["id"])
    resp = _verificar(client, mfa_token, codigo_totp(dados.segredo))
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA


def test_reset_de_senha_nao_desativa_2fa(client, tecnico_2fa):
    from app.core.security import sha256_hex
    from app.db.password_reset_db import PasswordResetDBManager

    usuario, _ = tecnico_2fa
    agora = datetime.now(timezone.utc).replace(tzinfo=None)
    PasswordResetDBManager().criar(
        usuario["id"], sha256_hex("tok-teste"), agora, agora + timedelta(minutes=30), "127.0.0.1"
    )
    resp = client.post("/api/auth/reset", json={"token": "tok-teste", "nova_senha": "NovaSenha#9"})
    assert resp.status_code == 200

    login = _login(client, "tecnico.teste", "NovaSenha#9")
    assert login.status_code == 200
    assert login.json()["etapa"] == "mfa"


def test_verificar_tem_rate_limit(chave_totp):
    with cliente_com_rate_limit("3/minute") as c:
        for _ in range(3):
            assert _verificar(c, "x", "000000").status_code == 400
        assert _verificar(c, "x", "000000").status_code == 429


def test_usuario_sem_2fa_continua_com_login_direto(client, usuario_tecnico, senha_padrao_teste):
    resp = _login(client, "tecnico.teste", senha_padrao_teste)
    assert resp.status_code == 200
    corpo = resp.json()
    assert corpo["etapa"] == "ok"
    assert corpo["access_token"]
    assert "refresh_token" in resp.cookies


def test_login_de_usuario_inexistente_tambem_confere_um_hash_bcrypt(client, usuario_tecnico, monkeypatch):
    """O tempo de resposta não pode revelar se o usuário existe: sempre 1 verificação de bcrypt."""
    import app.routers.auth as auth_router

    chamadas = []
    original = auth_router.verify_password

    def _espiao(plain, hashed):
        chamadas.append(hashed)
        return original(plain, hashed)

    monkeypatch.setattr(auth_router, "verify_password", _espiao)

    resp = _login(client, "ninguem.assim", "qualquer-senha")
    assert resp.status_code == 401
    assert resp.json()["detail"] == "Usuário ou senha inválidos."
    assert len(chamadas) == 1 and chamadas[0].startswith("$2b$")
