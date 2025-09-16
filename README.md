crie um arquivo no diretorio principal chamado "database_mysql.py" e adicione essas informações
do seu banco MYSQL:

import mysql.connector

DB_CONFIG = {
    "host": "localhost",
    "user": "user",          <<< usuário
    "password": "senha",     <<< coloque a senha aqui
    "database": "db_name"    <<< nome do banco
}

def get_connection():
    return mysql.connector.connect(**DB_CONFIG)
