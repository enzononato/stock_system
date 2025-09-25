# O nome do arquivo permanece inicializar_db.py

import bcrypt
from database_mysql import get_connection
import pymysql

# --- Defina aqui os seus usuários padrão ---
USUARIOS_PADRAO = [
    {'username': 'thiago', 'password': '123', 'role': 'Gestor'},
]

def setup_database():
    """
    Garante que a tabela 'usuarios' exista e, se estiver vazia,
    adiciona os usuários padrão com senhas hasheadas.
    Esta função deve ser chamada na inicialização do programa.
    """
    print("Iniciando verificação e configuração do banco de dados...")
    
    try:
        conn = get_connection()
        cur = conn.cursor()

        # 1. Garante que a tabela de usuários exista
        print("Verificando a existência da tabela 'usuarios'...")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INT AUTO_INCREMENT PRIMARY KEY,
                username VARCHAR(100) UNIQUE NOT NULL,
                password VARCHAR(255) NOT NULL,
                role ENUM('Gestor', 'Técnico', 'Jovem Aprendiz') NOT NULL
            )
        """)
        print("-> Tabela 'usuarios' verificada/criada com sucesso.")

        # 2. Verifica se a tabela já tem algum usuário
        cur.execute("SELECT COUNT(id) AS total FROM usuarios")
        row = cur.fetchone()
        numero_de_usuarios = int(row["total"])

        # 3. Se já existirem usuários, não faz nada
        if numero_de_usuarios > 0:
            print(f"O banco de dados já possui {numero_de_usuarios} usuário(s). Nenhuma ação de inserção necessária.")
            return

        # 4. Se a tabela estiver vazia, insere os usuários padrão
        print("Tabela de usuários vazia. Criando usuários padrão...")
        
        for usuario in USUARIOS_PADRAO:
            username = usuario['username']
            password_texto_puro = usuario['password']
            role = usuario['role']

            # É FUNDAMENTAL converter a senha para hash antes de salvar!
            hashed_password = bcrypt.hashpw(
                password_texto_puro.encode("utf-8"), bcrypt.gensalt()
            ).decode("utf-8")

            # Insere o usuário no banco
            cur.execute(
                "INSERT INTO usuarios (username, password, role) VALUES (%s, %s, %s)",
                (username, hashed_password, role)
            )
            print(f"- Usuário '{username}' criado.")
        
        conn.commit()
        print("\nUsuários padrão criados com sucesso!")

    except pymysql.MySQLError as e:
        print(f"Ocorreu um erro de banco de dados durante a configuração: {e}")
    finally:
        if 'conn' in locals() and conn:
            cur.close()
            conn.close()
            print("Conexão com o banco de dados fechada.")

# --- Mantemos esta parte para que o script ainda possa ser rodado manualmente se necessário ---
if __name__ == "__main__":
    setup_database()