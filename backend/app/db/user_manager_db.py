import logging

import bcrypt
import pymysql
from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)

MSG_EMAIL_DUPLICADO = "Este e-mail já está em uso por outro usuário."

_CAMPOS_PUBLICOS = "id, username, role, email, email_confirmado, totp_ativo, sessao_versao"
_CAMPOS_LOGIN = "id, username, password, role, email, email_confirmado, totp_ativo, sessao_versao"


def _erro_de_email_duplicado(exc: pymysql.MySQLError) -> bool:
    """1062 (chave duplicada) cuja mensagem cita o índice do e-mail."""
    # "Duplicate entry '<valor>' for key 'usuarios.uq_usuarios_email'": o nome do
    # índice vem no fim, então um username parecido não confunde a checagem.
    mensagem = str(exc.args[1]) if len(exc.args) > 1 else ""
    return exc.args[0] == 1062 and mensagem.rstrip().endswith("uq_usuarios_email'")


class EmailEmUso(Exception):
    """O e-mail já pertence a outra conta (índice único `uq_usuarios_email`)."""


class UserDBManager:
    def get_all_users(self):
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT id, username, role, email, email_confirmado, totp_ativo FROM usuarios")
            return cur.fetchall()

    def get_user_by_username(self, username: str):
        """Retorna o usuário completo (incluindo hash da senha) pelo username."""
        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT {_CAMPOS_LOGIN} FROM usuarios WHERE username = %s", (username,))
            return cur.fetchone()

    def get_user_para_login(self, identificador: str) -> dict | None:
        """Conta do login (com o hash da senha). Com `@`, pelo e-mail sem diferenciar
        caixa (confirmado ou não); sem `@`, pelo `username`. Quem decide se a conta
        pode entrar por esse caminho é o router, depois de conferir a senha."""
        identificador = (identificador or "").strip()
        if not identificador:
            return None
        with get_cursor(commit=False) as cur:
            if "@" in identificador:
                cur.execute(
                    f"SELECT {_CAMPOS_LOGIN} FROM usuarios WHERE LOWER(email) = %s",
                    (identificador.lower(),),
                )
            else:
                cur.execute(f"SELECT {_CAMPOS_LOGIN} FROM usuarios WHERE username = %s", (identificador,))
            return cur.fetchone()

    def email_em_uso_por_outro(self, email: str, user_id: int) -> bool:
        """Outra conta (confirmada ou não) já usa este e-mail (sem diferenciar caixa)?"""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT 1 FROM usuarios WHERE LOWER(email) = LOWER(%s) AND id <> %s LIMIT 1",
                (email, user_id),
            )
            return cur.fetchone() is not None

    def confirmar_email(self, user_id: int, email: str) -> None:
        """Grava o e-mail já confirmado pelo código. `EmailEmUso` se outra conta o usa."""
        try:
            with get_cursor(dict_cursor=False) as cur:
                cur.execute(
                    "UPDATE usuarios SET email = %s, email_confirmado = 1 WHERE id = %s",
                    (email, user_id),
                )
        except pymysql.MySQLError as e:
            if _erro_de_email_duplicado(e):
                raise EmailEmUso(MSG_EMAIL_DUPLICADO) from e
            raise

    def gerar_username(self, email: str) -> str:
        """`username` (só interno, para o histórico) a partir da parte antes do `@`;
        se já existir, ganha um número no fim: `joao.silva`, `joao.silva2`, `joao.silva3`..."""
        base = (email or "").strip().lower().split("@", 1)[0][:90] or "usuario"
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT LOWER(username) AS username FROM usuarios WHERE username LIKE %s",
                (base.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%",),
            )
            existentes = {linha["username"] for linha in cur.fetchall()}
        candidato, n = base, 2
        while candidato in existentes:
            candidato, n = f"{base}{n}", n + 1
        return candidato

    def get_user_by_id(self, user_id: int) -> dict | None:
        """Retorna o usuário (sem o hash da senha) pelo id, ou None se não existir."""
        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT {_CAMPOS_PUBLICOS} FROM usuarios WHERE id = %s", (user_id,))
            return cur.fetchone()

    def get_user_by_email(self, email: str) -> dict | None:
        """Usuário pelo e-mail, sem diferenciar caixa (sem o hash da senha)."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                f"SELECT {_CAMPOS_PUBLICOS} FROM usuarios WHERE LOWER(email) = LOWER(%s)",
                (email,),
            )
            return cur.fetchone()

    def get_estado_sessao(self, user_id: int) -> dict | None:
        """`role`, `sessao_versao` e `totp_ativo` pela PK (1 SELECT por requisição autenticada)."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT role, sessao_versao, totp_ativo FROM usuarios WHERE id = %s", (user_id,)
            )
            return cur.fetchone()

    def revogar_sessoes(self, user_id: int) -> None:
        """Encerra todas as sessões abertas do usuário (tokens antigos deixam de valer)."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute("UPDATE usuarios SET sessao_versao = sessao_versao + 1 WHERE id = %s", (user_id,))

    def count_gestores(self) -> int:
        """Quantidade de usuários ativos com role='Gestor'. Útil para impedir
        remover/rebaixar o último Gestor do sistema."""
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT COUNT(*) as total FROM usuarios WHERE role = 'Gestor'")
            return cur.fetchone()["total"]

    def add_user(self, username: str, password: str, role: str, email: str | None = None):
        if not username or not password or not role:
            return False, "Todos os campos são obrigatórios."
        try:
            hashed = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            with get_cursor(dict_cursor=False) as cur:
                cur.execute(
                    "INSERT INTO usuarios (username, password, role, email) VALUES (%s, %s, %s, %s)",
                    (username, hashed, role, email),
                )
            return True, f"Usuário '{username}' cadastrado com sucesso."
        except pymysql.MySQLError as e:
            if _erro_de_email_duplicado(e):
                return False, MSG_EMAIL_DUPLICADO
            if e.args[0] == 1062:
                return False, f"O nome de usuário '{username}' já existe."
            logger.exception("Erro ao cadastrar usuário")
            return False, "Erro ao cadastrar usuário."

    def update_email(self, user_id: int, email: str | None):
        """Troca o e-mail. Um e-mail diferente do atual volta a `email_confirmado = 0`
        (a pessoa confirma no próximo login); salvar o mesmo e-mail não mexe na
        confirmação. Não usa `rowcount` para saber se o usuário existe: o MySQL
        devolve 0 quando o valor não muda."""
        try:
            with get_cursor(dict_cursor=False) as cur:
                # O MySQL avalia o SET da esquerda para a direita: a confirmação é
                # decidida comparando com o e-mail ANTES da troca.
                cur.execute(
                    "UPDATE usuarios SET "
                    "email_confirmado = IF(LOWER(email) <=> LOWER(%s), email_confirmado, 0), "
                    "email = %s WHERE id = %s",
                    (email, email, user_id),
                )
            return True, "E-mail atualizado com sucesso."
        except pymysql.MySQLError as e:
            if _erro_de_email_duplicado(e):
                return False, MSG_EMAIL_DUPLICADO
            logger.exception("Erro ao atualizar o e-mail do usuário")
            return False, "Erro ao atualizar o e-mail."

    def remove_user(self, user_id: int):
        # `rowcount` é o que distingue "removeu" de "não havia nada para remover":
        # sem essa checagem, a camada reportava sucesso para um id inexistente e a
        # interface exibiria "Usuário removido com sucesso" sem ter removido nada.
        try:
            with get_cursor(dict_cursor=False) as cur:
                cur.execute("DELETE FROM usuarios WHERE id = %s", (user_id,))
                if cur.rowcount == 0:
                    return False, "Usuário não encontrado."
            return True, "Usuário removido com sucesso."
        except pymysql.MySQLError:
            logger.exception("Erro ao remover usuário")
            return False, "Erro ao remover usuário."

    def update_password(self, user_id: int, new_password: str):
        if not new_password:
            return False, "A nova senha não pode estar em branco."
        try:
            hashed = bcrypt.hashpw(new_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            with get_cursor(dict_cursor=False) as cur:
                cur.execute("UPDATE usuarios SET password = %s WHERE id = %s", (hashed, user_id))
                # Mesmo motivo do remove_user: um id inexistente afeta 0 linhas e
                # não deve ser reportado como troca de senha bem-sucedida.
                if cur.rowcount == 0:
                    return False, "Usuário não encontrado."
            return True, "Senha alterada com sucesso."
        except pymysql.MySQLError:
            logger.exception("Erro ao alterar a senha")
            return False, "Erro ao alterar a senha."
