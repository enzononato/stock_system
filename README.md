# Instalação e Configuração

Para começar, você precisa instalar as dependências do projeto e configurar o acesso ao seu banco de dados.

---

### 1. Instalar as Dependências

Primeiro, instale todas as bibliotecas necessárias, que estão listadas no arquivo **`requirements.txt`**. Abra o terminal na pasta do projeto e execute o comando:

```bash
pip install -r requirements.txt
```


2. Configurar o Banco de Dados
Agora, configure a conexão com o seu banco de dados MySQL.

Crie um arquivo chamado database_mysql.py na raiz do projeto.

Copie o código a seguir para dentro desse novo arquivo e substitua as informações de host, user, password e database pelos dados do seu banco.

Python

import mysql.connector

DB_CONFIG = {
    "host": "localhost",
    "user": "seu_usuario",       # Substitua pelo seu usuário do MySQL
    "password": "sua_senha",       # Substitua pela sua senha
    "database": "nome_do_banco"   # Substitua pelo nome do seu banco de dados
}

def get_connection():
    return mysql.connector.connect(**DB_CONFIG)
3. Inicializar o Projeto
Com a conexão configurada, você pode inicializar o banco de dados e criar o usuário padrão.

No terminal, execute o seguinte comando:

Bash

python inicializar_db.py
Usuário Padrão:

Usuário: thiago

Senha: 123