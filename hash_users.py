# ================================================================
# UTILIZAR APENAS SE TIVER ALGUM USUÁRIO NÃO TIVER A SENHA EM HASH
# ================================================================

import bcrypt
from database_mysql import get_connection

def hash_existing_passwords():
    """
    Este script busca por senhas que ainda não foram convertidas para hash
    e as atualiza no banco de dados.
    """
    conn = get_connection()
    cur = conn.cursor(dictionary=True)

    try:
        # Busca todos os usuários
        cur.execute("SELECT id, username, password FROM usuarios")
        users = cur.fetchall()

        print(f"Encontrados {len(users)} usuários. Verificando senhas...")

        for user in users:
            password = user["password"]

            # Um hash bcrypt começa com '$2b$'. Se a senha não começar assim,
            # significa que ela ainda está em texto plano.
            if not password.startswith('$2b$'):
                print(f"Senha do usuário '{user['username']}' (ID: {user['id']}) não está em hash. Atualizando...")
                
                # Gera o hash da senha
                hashed_password = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
                
                # Atualiza o registro no banco de dados com a senha hasheada
                cur.execute("UPDATE usuarios SET password = %s WHERE id = %s", (hashed_password, user["id"]))
                print(f"-> Usuário '{user['username']}' atualizado com sucesso.")
            else:
                print(f"Senha do usuário '{user['username']}' (ID: {user['id']}) já está em hash. Ignorando.")

        conn.commit()
        print("\nProcesso de hashing concluído.")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        conn.rollback()
    finally:
        cur.close()
        conn.close()

if __name__ == "__main__":
    hash_existing_passwords()