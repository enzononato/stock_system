import bcrypt
import pymysql
from app.db.database_mysql import get_connection


class UserDBManager:
    def get_all_users(self):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        cur.execute("SELECT id, username, role FROM usuarios")
        users = cur.fetchall()
        cur.close()
        conn.close()
        return users

    def get_user_by_username(self, username: str):
        """Retorna o usuário completo (incluindo hash da senha) pelo username."""
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        cur.execute("SELECT id, username, password, role FROM usuarios WHERE username = %s", (username,))
        user = cur.fetchone()
        cur.close()
        conn.close()
        return user

    def add_user(self, username: str, password: str, role: str):
        if not username or not password or not role:
            return False, "Todos os campos são obrigatórios."
        try:
            hashed = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            conn = get_connection()
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO usuarios (username, password, role) VALUES (%s, %s, %s)",
                (username, hashed, role),
            )
            conn.commit()
            cur.close()
            conn.close()
            return True, f"Usuário '{username}' cadastrado com sucesso."
        except pymysql.MySQLError as e:
            if e.errno == 1062:
                return False, f"O nome de usuário '{username}' já existe."
            return False, f"Erro ao cadastrar usuário: {e}"

    def remove_user(self, user_id: int):
        try:
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("DELETE FROM usuarios WHERE id = %s", (user_id,))
            conn.commit()
            cur.close()
            conn.close()
            return True, "Usuário removido com sucesso."
        except pymysql.MySQLError as e:
            return False, f"Erro ao remover usuário: {e}"

    def update_password(self, user_id: int, new_password: str):
        if not new_password:
            return False, "A nova senha não pode estar em branco."
        try:
            hashed = bcrypt.hashpw(new_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("UPDATE usuarios SET password = %s WHERE id = %s", (hashed, user_id))
            conn.commit()
            cur.close()
            conn.close()
            return True, "Senha alterada com sucesso."
        except pymysql.MySQLError as e:
            return False, f"Erro ao alterar a senha: {e}"
