"""Códigos de confirmação do e-mail de login (`email_confirmacoes`).

Só o HMAC do código é guardado. Datas em UTC, sempre passadas pelo Python.
"""
from datetime import datetime, timezone

from app.db.database_mysql import get_cursor


def _agora_utc() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


class EmailConfirmacaoDB:
    def criar(self, usuario_id: int, email: str, codigo_hash: str, expira_em: datetime) -> int:
        """Invalida os códigos pendentes do usuário e grava o novo, numa transação com
        a linha do usuário travada (pedidos simultâneos não deixam dois pendentes)."""
        with get_cursor() as cur:
            cur.execute("SELECT id FROM usuarios WHERE id = %s FOR UPDATE", (usuario_id,))
            cur.execute(
                "UPDATE email_confirmacoes SET usado = 1 WHERE usuario_id = %s AND usado = 0",
                (usuario_id,),
            )
            cur.execute(
                "INSERT INTO email_confirmacoes (usuario_id, email, codigo_hash, expira_em, criado_em) "
                "VALUES (%s, %s, %s, %s, %s)",
                (usuario_id, email, codigo_hash, expira_em, _agora_utc()),
            )
            return cur.lastrowid

    def apagar(self, confirmacao_id: int) -> None:
        """Remove um código que não chegou a ser enviado (falha no e-mail)."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute("DELETE FROM email_confirmacoes WHERE id = %s", (confirmacao_id,))

    def pendente(self, usuario_id: int) -> dict | None:
        """O código mais recente ainda não usado do usuário."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT id, usuario_id, email, codigo_hash, expira_em, tentativas, criado_em "
                "FROM email_confirmacoes WHERE usuario_id = %s AND usado = 0 "
                "ORDER BY id DESC LIMIT 1",
                (usuario_id,),
            )
            return cur.fetchone()

    def registrar_falha(self, confirmacao_id: int) -> int:
        """Soma uma tentativa errada (de forma atômica) e devolve o total."""
        with get_cursor() as cur:
            cur.execute(
                "UPDATE email_confirmacoes SET tentativas = tentativas + 1 WHERE id = %s",
                (confirmacao_id,),
            )
            cur.execute("SELECT tentativas FROM email_confirmacoes WHERE id = %s", (confirmacao_id,))
            linha = cur.fetchone()
            return int(linha["tentativas"]) if linha else 0

    def marcar_usado(self, confirmacao_id: int) -> bool:
        """Gasta o código. Devolve False se ele já estava usado (pedido simultâneo)."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "UPDATE email_confirmacoes SET usado = 1 WHERE id = %s AND usado = 0",
                (confirmacao_id,),
            )
            return cur.rowcount == 1

    def segundos_desde_ultimo(self, usuario_id: int) -> int | None:
        """Segundos desde o último código pedido pelo usuário (usado ou não), ou None."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT MAX(criado_em) AS ultimo FROM email_confirmacoes WHERE usuario_id = %s",
                (usuario_id,),
            )
            linha = cur.fetchone()
        if not linha or linha["ultimo"] is None:
            return None
        return int((_agora_utc() - linha["ultimo"]).total_seconds())
