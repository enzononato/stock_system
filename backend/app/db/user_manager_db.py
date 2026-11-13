import logging

import bcrypt
import pymysql
from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)

MSG_EMAIL_DUPLICADO = "Este e-mail já está em uso por outro usuário."

_CAMPOS_PUBLICOS = "id, username, role, email, totp_ativo, sessao_versao"


def _erro_de_email_duplicado(exc: pymysql.MySQLError) -> bool:
    """1062 (chave duplicada) cuja mensagem cita o índice do e-mail."""
    # "Duplicate entry '<valor>' for key 'usuarios.uq_usuarios_email'": o nome do
    # índice vem no fim, então um username parecido não confunde a checagem.
    mensagem = str(exc.args[1]) if len(exc.args) > 1 else ""
    return exc.args[0] == 1062 and mensagem.rstrip().endswith("uq_usuarios_email'")


class UserDBManager:
    def get_all_users(self):
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT id, username, role, email, totp_ativo FROM usuarios")
            return cur.fetchall()

    def get_user_by_username(self, username: str):
        """Retorna o usuário completo (incluindo hash da senha) pelo username."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT id, username, password, role, email, totp_ativo, sessao_versao "
                "FROM usuarios WHERE username = %s",
                (username,),
            )
            return cur.fetchone()

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
        """Troca o e-mail. Não usa `rowcount` para saber se o usuário existe: o
        MySQL devolve 0 quando o valor não muda."""
        try:
            with get_cursor(dict_cursor=False) as cur:
                cur.execute("UPDATE usuarios SET email = %s WHERE id = %s", (email, user_id))
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
