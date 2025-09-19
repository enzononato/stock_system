import bcrypt
from database_mysql import get_connection
from mysql.connector import Error

# --- Defina aqui os seus usuários padrão ---
USUARIOS_PADRAO = [
    {'username': 'thiago', 'password': '123', 'role': 'Gestor'},
    {'username': 'tecnico', 'password': '123', 'role': 'Técnico'},
    {'username': 'aprendiz', 'password': '123', 'role': 'Jovem Aprendiz'}
]

def inicializar_usuarios():
    """
    Verifica se a tabela de usuários está vazia e, se estiver,
    adiciona os usuários padrão com senhas hasheadas.
    """
    print("Iniciando verificação do banco de dados...")
    
    try:
        conn = get_connection()
        cur = conn.cursor()

        # 1. Verifica se a tabela já tem algum usuário
        cur.execute("SELECT COUNT(id) FROM usuarios")
        (numero_de_usuarios,) = cur.fetchone()

        # 2. Se já existirem usuários, não faz nada
        if numero_de_usuarios > 0:
            print(f"O banco de dados já possui {numero_de_usuarios} usuário(s). Nenhuma ação necessária.")
            return

        # 3. Se a tabela estiver vazia, insere os usuários padrão
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

    except Error as e:
        print(f"Ocorreu um erro de banco de dados: {e}")
    finally:
        if conn and conn.is_connected():
            cur.close()
            conn.close()
            print("Conexão com o banco de dados fechada.")

# --- Executa a função ao rodar o script ---
if __name__ == "__main__":
    inicializar_usuarios()