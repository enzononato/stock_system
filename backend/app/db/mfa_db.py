"""Estado do 2FA (TOTP) dos usuários e seus códigos de recuperação.

Guarda o segredo já cifrado e só o SHA-256 dos códigos. Datas em UTC, passadas pelo Python.
"""
from datetime import datetime, timedelta, timezone

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

    def salvar_segredo_pendente(self, usuario_id: int, cifrado: str) -> bool:
        """Grava o segredo ainda não confirmado, sobrescrevendo um pendente anterior.

        Só vale enquanto o 2FA NÃO está ativo (`WHERE totp_ativo = 0`, na própria
        UPDATE): um "iniciar" que chega depois da ativação não derruba o 2FA ativo.
        Devolve False se nada foi gravado (já ativo ou usuário inexistente)."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "UPDATE usuarios SET totp_secret_cifrado = %s, totp_ultimo_passo = NULL "
                "WHERE id = %s AND totp_ativo = 0",
                (cifrado, usuario_id),
            )
            return cur.rowcount == 1

    def _trocar_codigos(self, cur, usuario_id: int, hashes: list[str]) -> None:
        cur.execute("DELETE FROM mfa_recovery_codes WHERE usuario_id = %s", (usuario_id,))
        if hashes:
            cur.executemany(
                "INSERT INTO mfa_recovery_codes (usuario_id, code_hash) VALUES (%s, %s)",
                [(usuario_id, h) for h in hashes],
            )

    def ativar(
        self, usuario_id: int, hashes: list[str], passo: int, segredo_cifrado: str | None = None
    ) -> bool:
        """Ativa o 2FA e troca os códigos de recuperação, numa transação.

        A UPDATE só vale com o 2FA ainda inativo e, quando `segredo_cifrado` é dado, só
        ativa EXATAMENTE o segredo cujo código foi verificado (um "iniciar" concorrente
        que trocou o segredo no meio faz a ativação falhar). Devolve False se não ativou;
        nesse caso os códigos não são tocados."""
        with get_cursor(dict_cursor=False) as cur:
            sql = "UPDATE usuarios SET totp_ativo = 1, totp_ultimo_passo = %s WHERE id = %s AND totp_ativo = 0"
            params: list = [passo, usuario_id]
            if segredo_cifrado is not None:
                sql += " AND totp_secret_cifrado = %s"
                params.append(segredo_cifrado)
            cur.execute(sql, tuple(params))
            if cur.rowcount != 1:
                return False
            self._trocar_codigos(cur, usuario_id, hashes)
            return True

    # ── Falhas da 2ª etapa do login (limite por token e por usuário) ────────────

    def bloqueado_ate(self, usuario_id: int) -> datetime | None:
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT mfa_bloqueado_ate FROM usuarios WHERE id = %s", (usuario_id,))
            linha = cur.fetchone()
        return linha["mfa_bloqueado_ate"] if linha else None

    def registrar_falha(self, usuario_id: int, jti: str, agora: datetime, janela_inicio: datetime) -> tuple[int, int]:
        """Grava uma falha e devolve (falhas deste token, falhas do usuário desde `janela_inicio`).
        Também descarta falhas com mais de um dia."""
        with get_cursor() as cur:
            cur.execute(
                "INSERT INTO mfa_falhas (usuario_id, jti, criado_em) VALUES (%s, %s, %s)",
                (usuario_id, jti, agora),
            )
            cur.execute("SELECT COUNT(*) AS n FROM mfa_falhas WHERE usuario_id = %s AND jti = %s", (usuario_id, jti))
            do_token = cur.fetchone()["n"]
            cur.execute(
                "SELECT COUNT(*) AS n FROM mfa_falhas WHERE usuario_id = %s AND criado_em >= %s",
                (usuario_id, janela_inicio),
            )
            do_usuario = cur.fetchone()["n"]
            cur.execute("DELETE FROM mfa_falhas WHERE criado_em < %s", (agora - timedelta(days=1),))
        return do_token, do_usuario

    def bloquear(self, usuario_id: int, ate: datetime) -> None:
        with get_cursor(dict_cursor=False) as cur:
            cur.execute("UPDATE usuarios SET mfa_bloqueado_ate = %s WHERE id = %s", (ate, usuario_id))

    def limpar_falhas(self, usuario_id: int) -> None:
        """Sucesso na 2ª etapa: zera o contador e o bloqueio do usuário."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute("DELETE FROM mfa_falhas WHERE usuario_id = %s", (usuario_id,))
            cur.execute("UPDATE usuarios SET mfa_bloqueado_ate = NULL WHERE id = %s", (usuario_id,))

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
            "UPDATE usuarios SET totp_ativo = 0, totp_secret_cifrado = NULL, totp_ultimo_passo = NULL, "
            "mfa_bloqueado_ate = NULL WHERE id = %s",
            (usuario_id,),
        )
        cur.execute("DELETE FROM mfa_recovery_codes WHERE usuario_id = %s", (usuario_id,))
        cur.execute("DELETE FROM mfa_falhas WHERE usuario_id = %s", (usuario_id,))

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
