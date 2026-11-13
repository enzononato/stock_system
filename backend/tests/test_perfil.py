"""F2 Task 7: perfil e 2FA opcional para os demais papéis."""
import json

import pytest
from fastapi.testclient import TestClient

from tests.fixtures_auth import (  # noqa: F401
    ativar_2fa_direto,
    chave_totp,
    codigo_totp,
    gestor_2fa,
    linhas_auditoria,
    usuario_com_email,
)

pytestmark = pytest.mark.integration

MSG_JA_ATIVO = "A verificação em duas etapas já está ativa."
MSG_NAO_ATIVO = "A verificação em duas etapas não está ativa."
SENHA = "SenhaForte#123"  # senha_padrao_teste


def _cliente_logado(app, client, username, senha) -> TestClient:
    resp = client.post("/api/auth/login", json={"username": username, "password": senha})
    assert resp.status_code == 200 and resp.json()["etapa"] == "ok", resp.text
    c = TestClient(app)
    c.headers.update({"Authorization": f"Bearer {resp.json()['access_token']}"})
    return c


def _ativar_pelo_perfil(c, usuario_segredo=None):
    segredo = c.post("/api/perfil/2fa/iniciar", json={"senha": SENHA}).json()["segredo"]
    resp = c.post("/api/perfil/2fa/confirmar", json={"codigo": codigo_totp(segredo)})
    assert resp.status_code == 200, resp.text
    return segredo, resp.json()["codigos_recuperacao"]


def _linhas(sql, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


def test_perfil_exige_login(client):
    assert client.get("/api/perfil").status_code == 401


def test_get_perfil(_fastapi_app, client, chave_totp, usuario_com_email, senha_padrao_teste):
    c = _cliente_logado(_fastapi_app, client, "tecnico.email", senha_padrao_teste)
    resp = c.get("/api/perfil")
    assert resp.status_code == 200
    assert resp.json() == {
        "id": usuario_com_email["id"],
        "username": "tecnico.email",
        "role": "Técnico",
        "email": "tecnico@revalle.com.br",
        "totp_ativo": False,
        "mfa_obrigatorio": False,
        "mfa_disponivel": True,
        "codigos_recuperacao_restantes": 0,
    }


def test_perfil_sem_chave_indica_indisponivel(client_tecnico):
    from app.services.mfa import MSG_MFA_INDISPONIVEL

    perfil = client_tecnico.get("/api/perfil").json()
    assert perfil["mfa_disponivel"] is False
    resp = client_tecnico.post("/api/perfil/2fa/iniciar", json={"senha": SENHA})
    assert resp.status_code == 503
    assert resp.json()["detail"] == MSG_MFA_INDISPONIVEL


def test_tecnico_ativa_2fa_no_perfil(_fastapi_app, client, chave_totp, usuario_tecnico, senha_padrao_teste):
    c = _cliente_logado(_fastapi_app, client, "tecnico.teste", senha_padrao_teste)
    ini = c.post("/api/perfil/2fa/iniciar", json={"senha": SENHA})
    assert ini.status_code == 200, ini.text
    segredo = ini.json()["segredo"]
    assert ini.json()["otpauth_uri"].startswith("otpauth://totp/")

    conf = c.post("/api/perfil/2fa/confirmar", json={"codigo": codigo_totp(segredo)})
    assert conf.status_code == 200, conf.text
    codigos = conf.json()["codigos_recuperacao"]
    assert len(codigos) == 10 and len(set(codigos)) == 10

    perfil = c.get("/api/perfil").json()
    assert perfil["totp_ativo"] is True
    assert perfil["codigos_recuperacao_restantes"] == 10

    assert client.post(
        "/api/auth/login", json={"username": "tecnico.teste", "password": senha_padrao_teste}
    ).json()["etapa"] == "mfa"

    auditoria = linhas_auditoria("usuario", "ativar_2fa")
    assert len(auditoria) == 1
    assert auditoria[0]["username"] == "tecnico.teste"
    texto = json.dumps(auditoria, default=str)
    assert segredo not in texto and not any(cod in texto for cod in codigos)


def test_iniciar_com_2fa_ativo_409(_fastapi_app, client, usuario_tecnico, senha_padrao_teste, ativar_2fa_direto):
    c = _cliente_logado(_fastapi_app, client, "tecnico.teste", senha_padrao_teste)
    ativar_2fa_direto(usuario_tecnico["id"])
    resp = c.post("/api/perfil/2fa/iniciar", json={"senha": SENHA})
    assert resp.status_code == 409
    assert resp.json()["detail"] == MSG_JA_ATIVO


def test_confirmar_sem_iniciar_400(_fastapi_app, client, chave_totp, usuario_tecnico, senha_padrao_teste):
    c = _cliente_logado(_fastapi_app, client, "tecnico.teste", senha_padrao_teste)
    resp = c.post("/api/perfil/2fa/confirmar", json={"codigo": "123456"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Inicie o cadastro do autenticador antes de confirmar."


def test_confirmar_codigo_errado_400(_fastapi_app, client, chave_totp, usuario_tecnico, senha_padrao_teste):
    c = _cliente_logado(_fastapi_app, client, "tecnico.teste", senha_padrao_teste)
    segredo = c.post("/api/perfil/2fa/iniciar", json={"senha": SENHA}).json()["segredo"]
    certo = codigo_totp(segredo)
    errado = "000000" if certo != "000000" else "111111"
    resp = c.post("/api/perfil/2fa/confirmar", json={"codigo": errado})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Código inválido."


def test_desativar_exige_senha(_fastapi_app, client, chave_totp, usuario_tecnico, senha_padrao_teste):
    c = _cliente_logado(_fastapi_app, client, "tecnico.teste", senha_padrao_teste)
    _ativar_pelo_perfil(c)

    errada = c.post("/api/perfil/2fa/desativar", json={"senha": "errada"})
    assert errada.status_code == 400
    assert errada.json()["detail"] == "Senha incorreta."
    assert c.get("/api/perfil").json()["totp_ativo"] is True

    certa = c.post("/api/perfil/2fa/desativar", json={"senha": senha_padrao_teste})
    assert certa.status_code == 200
    assert certa.json()["detail"] == "Verificação em duas etapas desativada."

    linha = _linhas("SELECT totp_ativo, totp_secret_cifrado FROM usuarios WHERE id = %s", (usuario_tecnico["id"],))[0]
    assert linha["totp_ativo"] == 0 and linha["totp_secret_cifrado"] is None
    assert _linhas("SELECT 1 FROM mfa_recovery_codes WHERE usuario_id = %s", (usuario_tecnico["id"],)) == []
    assert len(linhas_auditoria("usuario", "desativar_2fa")) == 1
    assert client.post(
        "/api/auth/login", json={"username": "tecnico.teste", "password": senha_padrao_teste}
    ).json()["etapa"] == "ok"


def test_gestor_nao_desativa(gestor_2fa, senha_padrao_teste):
    resp = gestor_2fa["client"].post("/api/perfil/2fa/desativar", json={"senha": senha_padrao_teste})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Para o Gestor a verificação em duas etapas é obrigatória."
    perfil = gestor_2fa["client"].get("/api/perfil").json()
    assert perfil["totp_ativo"] is True and perfil["mfa_obrigatorio"] is True


def test_regenerar_codigos(_fastapi_app, client, chave_totp, usuario_tecnico, senha_padrao_teste):
    c = _cliente_logado(_fastapi_app, client, "tecnico.teste", senha_padrao_teste)
    segredo, antigos = _ativar_pelo_perfil(c)

    resp = c.post("/api/perfil/2fa/codigos", json={"senha": senha_padrao_teste})
    assert resp.status_code == 200, resp.text
    novos = resp.json()["codigos_recuperacao"]
    assert len(novos) == 10 and not set(novos) & set(antigos)

    def _entrar(codigo):
        mfa_token = client.post(
            "/api/auth/login", json={"username": "tecnico.teste", "password": senha_padrao_teste}
        ).json()["mfa_token"]
        return client.post("/api/auth/mfa/verificar", json={"mfa_token": mfa_token, "codigo": codigo})

    assert _entrar(antigos[0]).status_code == 400
    assert _entrar(novos[0]).status_code == 200

    auditoria = linhas_auditoria("usuario", "regenerar_codigos_2fa")
    assert len(auditoria) == 1
    texto = json.dumps(auditoria, default=str)
    assert not any(cod in texto for cod in novos + antigos)


def test_regenerar_senha_errada_400(_fastapi_app, client, chave_totp, usuario_tecnico, senha_padrao_teste):
    c = _cliente_logado(_fastapi_app, client, "tecnico.teste", senha_padrao_teste)
    _ativar_pelo_perfil(c)
    resp = c.post("/api/perfil/2fa/codigos", json={"senha": "errada"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Senha incorreta."


def test_regenerar_sem_2fa_400(_fastapi_app, client, usuario_tecnico, senha_padrao_teste):
    c = _cliente_logado(_fastapi_app, client, "tecnico.teste", senha_padrao_teste)
    resp = c.post("/api/perfil/2fa/codigos", json={"senha": senha_padrao_teste})
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_NAO_ATIVO


def test_iniciar_exige_a_senha_atual(_fastapi_app, client, chave_totp, usuario_tecnico, senha_padrao_teste):
    c = _cliente_logado(_fastapi_app, client, "tecnico.teste", senha_padrao_teste)

    errada = c.post("/api/perfil/2fa/iniciar", json={"senha": "errada"})
    assert errada.status_code == 400
    assert errada.json()["detail"] == "Senha incorreta."
    assert c.post("/api/perfil/2fa/iniciar", json={}).status_code == 422
    # Nada foi gravado pela tentativa com senha errada.
    assert _linhas("SELECT totp_secret_cifrado FROM usuarios WHERE id = %s", (usuario_tecnico["id"],))[0][
        "totp_secret_cifrado"
    ] is None

    certa = c.post("/api/perfil/2fa/iniciar", json={"senha": senha_padrao_teste})
    assert certa.status_code == 200
