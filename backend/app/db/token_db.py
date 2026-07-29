"""
Camada de acesso à tabela de tokens revogados.

Usada para permitir revogar um refresh token antes da sua expiração natural
(ex.: logout explícito, troca de senha) — sem isso, um refresh token JWT
continuaria válido até expirar mesmo depois do usuário "deslogar", já que o
JWT em si não é invalidado no servidor. A tabela funciona como uma blacklist:
guarda o `jti` (identificador único do token) e sua expiração original, só
para saber até quando vale a pena mantê-lo na lista.
"""
from datetime import datetime

from app.db.database_mysql import get_cursor


class TokenDBManager:
    def __init__(self):
        self._create_tables()

    def _create_tables(self):
        with get_cursor(dict_cursor=False) as cur:
            cur.execute("""
            CREATE TABLE IF NOT EXISTS revoked_tokens (
                jti VARCHAR(64) PRIMARY KEY,
                expires_at DATETIME NOT NULL,
                INDEX(expires_at)
            )
            """)

    def revoke(self, jti: str, expires_at: datetime) -> None:
        """Marca um token (pelo jti) como revogado até sua expiração natural.
        Idempotente: revogar o mesmo jti duas vezes apenas atualiza expires_at."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                """INSERT INTO revoked_tokens (jti, expires_at) VALUES (%s, %s)
                ON DUPLICATE KEY UPDATE expires_at = VALUES(expires_at)""",
                (jti, expires_at),
            )

    def is_revoked(self, jti: str) -> bool:
        """Verifica se um token (pelo jti) foi revogado."""
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT 1 FROM revoked_tokens WHERE jti = %s", (jti,))
            return cur.fetchone() is not None

    def purge_expired(self) -> int:
        """Remove da blacklist os tokens cuja expiração natural já passou (não
        precisam mais ser checados: o JWT já seria rejeitado por expiração).
        Retorna quantas linhas foram removidas."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute("DELETE FROM revoked_tokens WHERE expires_at < %s", (datetime.now(),))
            return cur.rowcount
