"""F2 Task 1: colunas novas de `usuarios` e e-mail no cadastro de usuários."""
import json
from pathlib import Path

import pytest

from tests.fixtures_auth import ida_e_volta_migracao, linhas_auditoria  # noqa: F401

pytestmark = pytest.mark.integration

COLUNAS = ["email", "totp_secret_cifrado", "totp_ativo", "totp_ultimo_passo", "sessao_versao"]


def _colunas_de_usuarios() -> dict:
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT column_name AS nome, data_type AS tipo, character_maximum_length AS tamanho, "
            "is_nullable AS anulavel, column_default AS padrao "
            "FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'usuarios'"
        )
        return {linha["nome"]: linha for linha in cur.fetchall()}


def _caminho_migracao_0200() -> Path:
    from app.db import migrations

    return next(
        p for p in sorted(Path(migrations.MIGRATIONS_DIR).glob("0200_*.sql")) if not p.name.endswith(".down.sql")
    )


def test_migracao_0200_cria_colunas_e_indice(_test_schema):
    from app.db.database_mysql import get_cursor

    colunas = _colunas_de_usuarios()
    assert colunas["email"]["tipo"] == "varchar"
    assert colunas["email"]["tamanho"] == 200
    assert colunas["email"]["anulavel"] == "YES"
    assert "totp_secret_cifrado" in colunas
    assert colunas["totp_ativo"]["padrao"] == "0"
    assert "totp_ultimo_passo" in colunas
    assert colunas["sessao_versao"]["padrao"] == "0"

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT non_unique AS nao_unico FROM information_schema.statistics "
            "WHERE table_schema = DATABASE() AND table_name = 'usuarios' AND index_name = 'uq_usuarios_email'"
        )
        linhas = cur.fetchall()
    assert linhas and all(l["nao_unico"] == 0 for l in linhas)


def test_migracao_0200_idempotente_e_reversivel(_test_schema):
    from app.db import migrations
    from app.db.database_mysql import get_cursor

    texto = _caminho_migracao_0200().read_text(encoding="utf-8")
    for _ in range(2):
        with get_cursor() as cur:
            migrations.executar_sql(cur, texto)

    ida_e_volta_migracao(200)

    colunas = _colunas_de_usuarios()
    assert all(c in colunas for c in COLUNAS)


def test_usuarios_legados_sem_email_e_sem_2fa(client, client_gestor, executar_sql):
    from app.core.security import hash_password

    executar_sql(
        "INSERT INTO usuarios (username, password, role) VALUES ('legado', %s, 'Técnico')",
        (hash_password("Senha#1"),),
    )

    lista = client_gestor.get("/api/users").json()
    legado = next(u for u in lista if u["username"] == "legado")
    assert legado["email"] is None
    assert legado["totp_ativo"] is False

    resp = client.post("/api/auth/login", json={"username": "legado", "password": "Senha#1"})
    assert resp.status_code == 200
    assert resp.json()["access_token"]


def test_criar_usuario_com_email_normaliza(client_gestor):
    resp = client_gestor.post(
        "/api/users",
        json={"username": "joao", "password": "Abc#1234", "role": "Técnico", "email": "Joao.Silva@Revalle.com.br"},
    )
    assert resp.status_code == 200, resp.text

    lista = client_gestor.get("/api/users").json()
    assert next(u for u in lista if u["username"] == "joao")["email"] == "joao.silva@revalle.com.br"

    linha = linhas_auditoria("usuario", "criar")[-1]
    assert linha["depois"]["email"] == "joao.silva@revalle.com.br"
    assert "$2b$" not in json.dumps(linha, default=str)


def test_criar_usuario_email_vazio_vira_null(client_gestor):
    resp = client_gestor.post(
        "/api/users", json={"username": "maria", "password": "Abc#1234", "role": "Técnico", "email": "   "}
    )
    assert resp.status_code == 200
    lista = client_gestor.get("/api/users").json()
    assert next(u for u in lista if u["username"] == "maria")["email"] is None


def test_criar_usuario_email_invalido_422(client_gestor):
    resp = client_gestor.post(
        "/api/users", json={"username": "maria", "password": "Abc#1234", "role": "Técnico", "email": "joao@"}
    )
    assert resp.status_code == 422
    assert "E-mail inválido." in resp.json()["detail"]


def test_criar_usuario_email_duplicado_400(client_gestor):
    base = {"password": "Abc#1234", "role": "Técnico", "email": "dup@revalle.com.br"}
    assert client_gestor.post("/api/users", json={**base, "username": "u1"}).status_code == 200
    resp = client_gestor.post("/api/users", json={**base, "username": "u2", "email": "DUP@revalle.com.br"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Este e-mail já está em uso por outro usuário."


def test_editar_email(client_gestor, usuario_tecnico):
    resp = client_gestor.put(f"/api/users/{usuario_tecnico['id']}", json={"email": "tec@revalle.com.br"})
    assert resp.status_code == 200, resp.text
    assert resp.json()["email"] == "tec@revalle.com.br"

    linha = linhas_auditoria("usuario", "editar")[-1]
    assert linha["entidade_id"] == str(usuario_tecnico["id"])
    assert linha["antes"] == {"email": None}
    assert linha["depois"] == {"email": "tec@revalle.com.br"}


def test_editar_email_null_limpa(client_gestor, usuario_tecnico):
    uid = usuario_tecnico["id"]
    client_gestor.put(f"/api/users/{uid}", json={"email": "tec@revalle.com.br"})
    resp = client_gestor.put(f"/api/users/{uid}", json={"email": None})
    assert resp.status_code == 200
    assert resp.json()["email"] is None


def test_editar_email_duplicado_400(client_gestor, usuario_tecnico, usuario_aprendiz):
    client_gestor.put(f"/api/users/{usuario_tecnico['id']}", json={"email": "x@revalle.com.br"})
    resp = client_gestor.put(f"/api/users/{usuario_aprendiz['id']}", json={"email": "x@revalle.com.br"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Este e-mail já está em uso por outro usuário."


def test_editar_mesmo_email_nao_e_erro(client_gestor, usuario_tecnico):
    uid = usuario_tecnico["id"]
    client_gestor.put(f"/api/users/{uid}", json={"email": "x@revalle.com.br"})
    assert client_gestor.put(f"/api/users/{uid}", json={"email": "x@revalle.com.br"}).status_code == 200


def test_editar_sem_chave_email_422(client_gestor, usuario_tecnico):
    resp = client_gestor.put(f"/api/users/{usuario_tecnico['id']}", json={})
    assert resp.status_code == 422


def test_editar_usuario_inexistente_404(client_gestor):
    resp = client_gestor.put("/api/users/99999", json={"email": "a@revalle.com.br"})
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Usuário não encontrado."


def test_editar_email_exige_gestor(client, client_tecnico, usuario_tecnico):
    url = f"/api/users/{usuario_tecnico['id']}"
    resp = client_tecnico.put(url, json={"email": "a@revalle.com.br"})
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito ao Gestor."
    assert client.put(url, json={"email": "a@revalle.com.br"}).status_code == 401


def test_listar_traz_email_e_totp_ativo(client_gestor):
    lista = client_gestor.get("/api/users").json()
    assert lista
    for item in lista:
        assert "email" in item and "totp_ativo" in item
        assert "password" not in item
