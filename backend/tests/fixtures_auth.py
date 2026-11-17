"""Apoio dos testes da Fase 2 (login, 2FA, redefinição de senha, perfil).

Os testes da F2 importam o que usam daqui (`# noqa: F401`); estes fixtures não
entram no `conftest.py`.
"""
import json
from pathlib import Path


def linhas_auditoria(entidade: str, acao: str | None = None) -> list[dict]:
    """Linhas de `audit_log` da entidade (e da ação, se dada), em ordem de id,
    com `antes`/`depois` já decodificados."""
    from app.db.database_mysql import get_cursor

    sql = "SELECT * FROM audit_log WHERE entidade = %s"
    params: list = [entidade]
    if acao is not None:
        sql += " AND acao = %s"
        params.append(acao)
    sql += " ORDER BY id"
    with get_cursor(commit=False) as cur:
        cur.execute(sql, tuple(params))
        linhas = list(cur.fetchall())
    for linha in linhas:
        for chave in ("antes", "depois"):
            valor = linha.get(chave)
            if isinstance(valor, (str, bytes)):
                linha[chave] = json.loads(valor)
    return linhas


def ida_e_volta_migracao(versao: int) -> None:
    """Roda o `.down.sql` e depois o `.sql` da migração `versao` (o `.sql` roda
    num `finally`, para nunca deixar o banco de teste sem a migração)."""
    from app.db import migrations
    from app.db.database_mysql import get_cursor

    pasta = Path(migrations.MIGRATIONS_DIR)
    up = next(p for p in sorted(pasta.glob(f"{versao:04d}_*.sql")) if not p.name.endswith(".down.sql"))
    down = up.with_name(f"{up.stem}.down.sql")
    try:
        with get_cursor() as cur:
            migrations.executar_sql(cur, down.read_text(encoding="utf-8"))
    finally:
        with get_cursor() as cur:
            migrations.executar_sql(cur, up.read_text(encoding="utf-8"))


# ── Redefinição de senha (Task 3) ───────────────────────────────────────────
import importlib
import os
from contextlib import contextmanager

import pytest

_MODULOS_DO_RATE_LIMIT = (
    "app.core.config",
    "app.dependencies",
    "app.routers.auth",
    "app.routers.auth_senha",
    "app.routers.perfil",
    "app.main",
)


def _recarregar_modulos() -> None:
    for nome in _MODULOS_DO_RATE_LIMIT:
        try:
            modulo = importlib.import_module(nome)
        except ModuleNotFoundError:
            continue  # módulo de uma tarefa posterior ainda não existe
        importlib.reload(modulo)


@contextmanager
def cliente_com_rate_limit(valor: str):
    """TestClient de um app recarregado com `LOGIN_RATE_LIMIT=valor`.

    O decorator `@limiter.limit(settings.LOGIN_RATE_LIMIT)` congela o valor na
    importação; por isso recarregamos a cadeia config -> dependencies ->
    routers -> main com a variável já baixa e desfazemos tudo no fim (o
    `_fastapi_app` da sessão não pode herdar um limite baixo). Mesma técnica de
    `_cliente_com_login_rate_limit` em tests/test_auth.py.
    """
    from fastapi.testclient import TestClient

    original = os.environ.get("LOGIN_RATE_LIMIT")
    os.environ["LOGIN_RATE_LIMIT"] = valor
    try:
        _recarregar_modulos()
        import app.main as main_module

        with TestClient(main_module.app) as c:
            yield c
    finally:
        if original is None:
            os.environ.pop("LOGIN_RATE_LIMIT", None)
        else:
            os.environ["LOGIN_RATE_LIMIT"] = original
        _recarregar_modulos()


@pytest.fixture
def usuario_com_email(criar_usuario, senha_padrao_teste, executar_sql) -> dict:
    """Técnico `tecnico.email` com e-mail `tecnico@revalle.com.br`."""
    from app.db.user_manager_db import UserDBManager

    criar_usuario("tecnico.email", senha_padrao_teste, "Técnico")
    executar_sql(
        "UPDATE usuarios SET email = %s WHERE username = %s",
        ("tecnico@revalle.com.br", "tecnico.email"),
    )
    return UserDBManager().get_user_by_username("tecnico.email")


# ── 2FA (Task 5) ────────────────────────────────────────────────────────────
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone


@pytest.fixture
def chave_totp(monkeypatch) -> str:
    """Liga o 2FA no servidor com uma chave Fernet nova (o runner não define a chave)."""
    from cryptography.fernet import Fernet

    from app.core import config

    chave = Fernet.generate_key().decode()
    monkeypatch.setattr(config.settings, "TOTP_ENCRYPTION_KEY", chave)
    return chave


@dataclass
class Dados2fa:
    segredo: str
    codigos: list[str]


@pytest.fixture
def ativar_2fa_direto(chave_totp) -> Callable[[int], Dados2fa]:
    """Ativa o 2FA de um usuário direto no banco (sem passar pela API)."""

    def _ativar(usuario_id: int) -> Dados2fa:
        from app.db.mfa_db import MfaDBManager
        from app.services import mfa

        db = MfaDBManager()
        segredo = mfa.gerar_segredo()
        db.salvar_segredo_pendente(usuario_id, mfa.cifrar_segredo(segredo))
        codigos = mfa.gerar_codigos_recuperacao()
        db.ativar(usuario_id, [mfa.hash_codigo_recuperacao(c) for c in codigos], passo=0)
        return Dados2fa(segredo=segredo, codigos=codigos)

    return _ativar


def codigo_totp(segredo: str, passos: int = 0) -> str:
    """Código TOTP de agora (+30 s por passo). Use `passos=1` num segundo login do
    mesmo usuário no mesmo teste, por causa do anti-replay."""
    import pyotp

    return pyotp.TOTP(segredo).at(datetime.now(timezone.utc) + timedelta(seconds=30 * passos))


def login_com_2fa(client, identificador: str, senha: str, codigo: str) -> str:
    """Login completo em duas etapas; devolve o access token."""
    resp = client.post("/api/auth/login", json={"username": identificador, "password": senha})
    assert resp.status_code == 200, resp.text
    assert resp.json()["etapa"] == "mfa", resp.text
    ver = client.post(
        "/api/auth/mfa/verificar", json={"mfa_token": resp.json()["mfa_token"], "codigo": codigo}
    )
    assert ver.status_code == 200, ver.text
    return ver.json()["access_token"]


@pytest.fixture
def gestor_2fa(_fastapi_app, client, usuario_gestor, senha_padrao_teste, ativar_2fa_direto) -> dict:
    """Gestor com 2FA ativo e um TestClient autenticado. Com `chave_totp` ligada, os
    testes não usam `client_gestor` (o login do Gestor sem 2FA não devolve access token)."""
    from fastapi.testclient import TestClient

    dados = ativar_2fa_direto(usuario_gestor["id"])
    access = login_com_2fa(client, usuario_gestor["email"], senha_padrao_teste, codigo_totp(dados.segredo))
    c = TestClient(_fastapi_app)
    c.headers.update({"Authorization": f"Bearer {access}"})
    return {"usuario": usuario_gestor, "segredo": dados.segredo, "codigos": dados.codigos, "client": c}


# ── Login por e-mail ────────────────────────────────────────────────────────
def email_de_teste(identificador: str) -> str:
    """Contas de teste entram com `<username>@revalle.com.br` (ver `criar_usuario`)."""
    return identificador if "@" in identificador else f"{identificador}@revalle.com.br"
