import bcrypt
from database_mysql import get_connection
from mysql.connector import Error

class UserDBManager:
    """
    Gerencia as operações de CRUD para a tabela de usuários.
    """

    def get_all_users(self):
        """Retorna uma lista de todos os usuários (sem o hash da senha)."""
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        # Nunca retorne o campo 'password' para a interface
        cur.execute("SELECT id, username, role FROM usuarios")
        users = cur.fetchall()
        cur.close()
        conn.close()
        return users

    def add_user(self, username, password, role):
        """Adiciona um novo usuário, com a senha já convertida para hash."""
        if not username or not password or not role:
            return False, "Todos os campos são obrigatórios."
            
        try:
            # É crucial que a nova senha também seja convertida para hash
            hashed_password = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            
            conn = get_connection()
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO usuarios (username, password, role) VALUES (%s, %s, %s)",
                (username, hashed_password, role)
            )
            conn.commit()
            cur.close()
            conn.close()
            return True, f"Usuário '{username}' cadastrado com sucesso."
        except Error as e:
            # Código 1062 é para entrada duplicada (username UNIQUE)
            if e.errno == 1062:
                return False, f"O nome de usuário '{username}' já existe."
            return False, f"Erro ao cadastrar usuário: {e}"

    def remove_user(self, user_id):
        """Remove um usuário pelo ID."""
        try:
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("DELETE FROM usuarios WHERE id = %s", (user_id,))
            conn.commit()
            cur.close()
            conn.close()
            return True, "Usuário removido com sucesso."
        except Error as e:
            return False, f"Erro ao remover usuário: {e}"

    def update_password(self, user_id, new_password):
        """Atualiza a senha de um usuário, já aplicando o hash."""
        if not new_password:
            return False, "A nova senha não pode estar em branco."
            
        try:
            hashed_password = bcrypt.hashpw(new_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("UPDATE usuarios SET password = %s WHERE id = %s", (hashed_password, user_id))
            conn.commit()
            cur.close()
            conn.close()
            return True, "Senha alterada com sucesso."
        except Error as e:
            return False, f"Erro ao alterar a senha: {e}"