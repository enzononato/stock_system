"""Links de redefinição de senha (`password_reset_tokens`).

Só o SHA-256 do token é guardado. Datas em UTC, sempre passadas pelo Python.
"""
from datetime import datetime

from app.db.database_mysql import get_cursor


class PasswordResetDBManager:
    def criar(
        self,
        usuario_id: int,
        token_hash: str,
        criado_em: datetime,
        expira_em: datetime,
        criado_ip: str | None,
    ) -> None:
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "INSERT INTO password_reset_tokens (usuario_id, token_hash, criado_em, expira_em, criado_ip) "
                "VALUES (%s, %s, %s, %s, %s)",
                (usuario_id, token_hash, criado_em, expira_em, criado_ip),
            )

    def contar_criados_desde(self, usuario_id: int, desde: datetime) -> int:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT COUNT(*) AS total FROM password_reset_tokens WHERE usuario_id = %s AND criado_em >= %s",
                (usuario_id, desde),
            )
            return cur.fetchone()["total"]

    def consumir(self, token_hash: str, agora: datetime) -> int | None:
        """Marca o link como usado e devolve o `usuario_id`, ou None se o link não
        existe, já foi usado ou expirou. Um único UPDATE condicional: atômico
        contra uso duplo simultâneo."""
        with get_cursor() as cur:
            cur.execute(
                "UPDATE password_reset_tokens SET usado_em = %s "
                "WHERE token_hash = %s AND usado_em IS NULL AND expira_em > %s",
                (agora, token_hash, agora),
            )
            if cur.rowcount != 1:
                return None
            cur.execute("SELECT usuario_id FROM password_reset_tokens WHERE token_hash = %s", (token_hash,))
            return cur.fetchone()["usuario_id"]

    def invalidar_abertos(self, usuario_id: int, agora: datetime) -> None:
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "UPDATE password_reset_tokens SET usado_em = %s WHERE usuario_id = %s AND usado_em IS NULL",
                (agora, usuario_id),
            )
