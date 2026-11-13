"""Estado do 2FA (TOTP) dos usuários e seus códigos de recuperação.

Guarda o segredo já cifrado e só o SHA-256 dos códigos. Datas em UTC, passadas pelo Python.
"""
from datetime import datetime, timezone

from app.db.database_mysql import get_cursor


def _agora_utc() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


class MfaDBManager:
    def estado(self, usuario_id: int) -> dict | None:
        """`totp_secret_cifrado`, `totp_ativo` e `totp_ultimo_passo` do usuário."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT totp_secret_cifrado, totp_ativo, totp_ultimo_passo FROM usuarios WHERE id = %s",
                (usuario_id,),
            )
            return cur.fetchone()

    def salvar_segredo_pendente(self, usuario_id: int, cifrado: str) -> None:
        """Grava o segredo ainda não confirmado (`totp_ativo=0`), sobrescrevendo um anterior."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "UPDATE usuarios SET totp_secret_cifrado = %s, totp_ativo = 0, totp_ultimo_passo = NULL "
                "WHERE id = %s",
                (cifrado, usuario_id),
            )

    def _trocar_codigos(self, cur, usuario_id: int, hashes: list[str]) -> None:
        cur.execute("DELETE FROM mfa_recovery_codes WHERE usuario_id = %s", (usuario_id,))
        if hashes:
            cur.executemany(
                "INSERT INTO mfa_recovery_codes (usuario_id, code_hash) VALUES (%s, %s)",
                [(usuario_id, h) for h in hashes],
            )

    def ativar(self, usuario_id: int, hashes: list[str], passo: int) -> None:
        """Ativa o 2FA e troca os códigos de recuperação, numa transação."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "UPDATE usuarios SET totp_ativo = 1, totp_ultimo_passo = %s WHERE id = %s",
                (passo, usuario_id),
            )
            self._trocar_codigos(cur, usuario_id, hashes)

    def substituir_codigos(self, usuario_id: int, hashes: list[str]) -> None:
        """Troca todos os códigos de recuperação (os antigos deixam de valer)."""
        with get_cursor(dict_cursor=False) as cur:
            self._trocar_codigos(cur, usuario_id, hashes)

    def desativar(self, usuario_id: int) -> None:
        """Desliga o 2FA: apaga o segredo, o último passo e os códigos. Não mexe em `sessao_versao`."""
        with get_cursor(dict_cursor=False) as cur:
            self._desligar(cur, usuario_id)

    def _desligar(self, cur, usuario_id: int) -> None:
        cur.execute(
            "UPDATE usuarios SET totp_ativo = 0, totp_secret_cifrado = NULL, totp_ultimo_passo = NULL "
            "WHERE id = %s",
            (usuario_id,),
        )
        cur.execute("DELETE FROM mfa_recovery_codes WHERE usuario_id = %s", (usuario_id,))

    def zerar(self, usuario_id: int) -> None:
        """Como `desativar`, mais `sessao_versao + 1` (encerra as sessões do usuário), numa transação."""
        with get_cursor(dict_cursor=False) as cur:
            self._desligar(cur, usuario_id)
            cur.execute("UPDATE usuarios SET sessao_versao = sessao_versao + 1 WHERE id = %s", (usuario_id,))

    def registrar_passo(self, usuario_id: int, passo: int) -> bool:
        """Anti-replay: grava o passo só se for maior que o último aceito. False se já usado."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "UPDATE usuarios SET totp_ultimo_passo = %s "
                "WHERE id = %s AND (totp_ultimo_passo IS NULL OR totp_ultimo_passo < %s)",
                (passo, usuario_id, passo),
            )
            return cur.rowcount == 1

    def consumir_codigo(self, usuario_id: int, code_hash: str) -> bool:
        """Marca o código como usado; False se não existe ou já foi usado (atômico)."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "UPDATE mfa_recovery_codes SET usado_em = %s "
                "WHERE usuario_id = %s AND code_hash = %s AND usado_em IS NULL",
                (_agora_utc(), usuario_id, code_hash),
            )
            return cur.rowcount == 1

    def contar_codigos_restantes(self, usuario_id: int) -> int:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT COUNT(*) AS total FROM mfa_recovery_codes WHERE usuario_id = %s AND usado_em IS NULL",
                (usuario_id,),
            )
            return cur.fetchone()["total"]
