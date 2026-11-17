"""Revisão: emitir um link de senha (convite ou envio do Gestor) é atômico."""
import pytest

pytestmark = pytest.mark.integration


def _abertos(usuario_id: int) -> int:
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT COUNT(*) AS n FROM password_reset_tokens WHERE usuario_id = %s AND usado_em IS NULL",
            (usuario_id,),
        )
        return int(cur.fetchone()["n"])


def test_substituir_deixa_um_link_aberto(usuario_tecnico):
    from datetime import datetime, timedelta

    from app.db.password_reset_db import PasswordResetDBManager

    db = PasswordResetDBManager()
    agora = datetime(2026, 11, 17, 12, 0, 0)
    assert db.substituir(usuario_tecnico["id"], "a" * 64, agora, agora + timedelta(hours=72)) is True
    assert db.substituir(usuario_tecnico["id"], "b" * 64, agora, agora + timedelta(hours=72)) is True
    assert _abertos(usuario_tecnico["id"]) == 1
    assert db.substituir(999999, "c" * 64, agora, agora + timedelta(hours=72)) is False


def test_emitir_link_usa_substituir(usuario_tecnico, monkeypatch):
    from app.db.password_reset_db import PasswordResetDBManager
    from app.routers import auth_senha

    chamadas = []
    original = PasswordResetDBManager.substituir

    def _espiao(self, *args, **kwargs):
        chamadas.append(args)
        return original(self, *args, **kwargs)

    monkeypatch.setattr(PasswordResetDBManager, "substituir", _espiao)
    db = PasswordResetDBManager()
    auth_senha.emitir_link(usuario_tecnico["id"], True, db)
    auth_senha.emitir_link(usuario_tecnico["id"], False, db)
    assert len(chamadas) == 2
    assert _abertos(usuario_tecnico["id"]) == 1
