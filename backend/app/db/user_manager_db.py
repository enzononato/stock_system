import bcrypt
import pymysql
from app.db.database_mysql import get_cursor


class UserDBManager:
    def get_all_users(self):
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT id, username, role FROM usuarios")
            return cur.fetchall()

    def get_user_by_username(self, username: str):
        """Retorna o usuário completo (incluindo hash da senha) pelo username."""
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT id, username, password, role FROM usuarios WHERE username = %s", (username,))
            return cur.fetchone()

    def get_user_by_id(self, user_id: int) -> dict | None:
        """Retorna o usuário (sem o hash da senha) pelo id, ou None se não existir."""
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT id, username, role FROM usuarios WHERE id = %s", (user_id,))
            return cur.fetchone()

    def count_gestores(self) -> int:
        """Quantidade de usuários ativos com role='Gestor'. Útil para impedir
        remover/rebaixar o último Gestor do sistema."""
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT COUNT(*) as total FROM usuarios WHERE role = 'Gestor'")
            return cur.fetchone()["total"]

    def add_user(self, username: str, password: str, role: str):
        if not username or not password or not role:
            return False, "Todos os campos são obrigatórios."
        try:
            hashed = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            with get_cursor(dict_cursor=False) as cur:
                cur.execute(
                    "INSERT INTO usuarios (username, password, role) VALUES (%s, %s, %s)",
                    (username, hashed, role),
                )
            return True, f"Usuário '{username}' cadastrado com sucesso."
        except pymysql.MySQLError as e:
            if e.args[0] == 1062:
                return False, f"O nome de usuário '{username}' já existe."
            return False, f"Erro ao cadastrar usuário: {e}"

    def remove_user(self, user_id: int):
        try:
            with get_cursor(dict_cursor=False) as cur:
                cur.execute("DELETE FROM usuarios WHERE id = %s", (user_id,))
            return True, "Usuário removido com sucesso."
        except pymysql.MySQLError as e:
            return False, f"Erro ao remover usuário: {e}"

    def update_password(self, user_id: int, new_password: str):
        if not new_password:
            return False, "A nova senha não pode estar em branco."
        try:
            hashed = bcrypt.hashpw(new_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            with get_cursor(dict_cursor=False) as cur:
                cur.execute("UPDATE usuarios SET password = %s WHERE id = %s", (hashed, user_id))
            return True, "Senha alterada com sucesso."
        except pymysql.MySQLError as e:
            return False, f"Erro ao alterar a senha: {e}"
