"""F2 Task 2: versão da sessão (encerrar todas as sessões de um usuário)."""
import pytest
from jose import jwt as jose_jwt

pytestmark = pytest.mark.integration

MSG = "Sessão encerrada. Faça login novamente."


def _login(client, username, senha):
    resp = client.post("/api/auth/login", json={"username": username, "password": senha})
    assert resp.status_code == 200, resp.text
    return resp


def test_tokens_levam_sv(client, usuario_tecnico, senha_padrao_teste):
    resp = _login(client, "tecnico.teste", senha_padrao_teste)
    assert jose_jwt.get_unverified_claims(resp.json()["access_token"])["sv"] == 0
    assert jose_jwt.get_unverified_claims(resp.cookies["refresh_token"])["sv"] == 0


def test_revogar_sessoes_derruba_access_e_refresh(client, usuario_tecnico, senha_padrao_teste):
    from app.db.user_manager_db import UserDBManager

    resp = _login(client, "tecnico.teste", senha_padrao_teste)
    access = resp.json()["access_token"]
    cookie = resp.cookies["refresh_token"]

    UserDBManager().revogar_sessoes(usuario_tecnico["id"])

    me = client.get("/api/auth/me", headers={"Authorization": f"Bearer {access}"})
    assert me.status_code == 401
    assert me.json()["detail"] == MSG

    client.cookies.clear()
    client.cookies.set("refresh_token", cookie)
    ref = client.post("/api/auth/refresh")
    assert ref.status_code == 401
    assert ref.json()["detail"] == MSG

    client.cookies.clear()
    novo = _login(client, "tecnico.teste", senha_padrao_teste)
    novo_access = novo.json()["access_token"]
    assert jose_jwt.get_unverified_claims(novo_access)["sv"] == 1
    assert client.get("/api/auth/me", headers={"Authorization": f"Bearer {novo_access}"}).status_code == 200


def test_token_legado_sem_sv_continua_valido(client, usuario_tecnico):
    from app.core.security import create_access_token

    token = create_access_token(
        {"sub": str(usuario_tecnico["id"]), "username": "tecnico.teste", "role": "Técnico"}
    )
    resp = client.get("/api/auth/me", headers={"Authorization": f"Bearer {token}"})
    assert resp.status_code == 200


def test_token_de_usuario_removido_e_recusado(client, client_gestor, usuario_tecnico, senha_padrao_teste):
    access = _login(client, "tecnico.teste", senha_padrao_teste).json()["access_token"]
    assert client_gestor.delete(f"/api/users/{usuario_tecnico['id']}").status_code == 200

    resp = client.get("/api/auth/me", headers={"Authorization": f"Bearer {access}"})
    assert resp.status_code == 401
    assert resp.json()["detail"] == MSG


def test_refresh_mantem_sv(client, usuario_tecnico, senha_padrao_teste):
    from app.db.user_manager_db import UserDBManager

    UserDBManager().revogar_sessoes(usuario_tecnico["id"])
    resp = _login(client, "tecnico.teste", senha_padrao_teste)
    assert jose_jwt.get_unverified_claims(resp.json()["access_token"])["sv"] == 1

    ref = client.post("/api/auth/refresh")
    assert ref.status_code == 200, ref.text
    novo = ref.json()["access_token"]
    assert jose_jwt.get_unverified_claims(novo)["sv"] == 1
    assert client.get("/api/auth/me", headers={"Authorization": f"Bearer {novo}"}).status_code == 200
