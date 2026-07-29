"""
Script IDEMPOTENTE de bootstrap do banco de dados.

Problema que resolve: a tabela `usuarios` não é criada por nenhum código do
backend FastAPI — antes só existia via `inicializar_db.py`, script do
antigo app desktop que está sendo removido nesta mesma leva de mudanças.
Sem ele, um deploy limpo (MySQL vazio) sobe a API sem nenhuma tabela
`usuarios` e, consequentemente, sem nenhum usuário capaz de logar.

O que este script faz, nesta ordem:
1. Cria a tabela `usuarios` se ela ainda não existir (CREATE TABLE IF NOT
   EXISTS — não mexe em nada se ela já existir, então rodar em cima do banco
   de produção atual é inofensivo).
2. Semeia o primeiro usuário com role='Gestor', lendo credenciais das
   variáveis de ambiente BOOTSTRAP_ADMIN_USER / BOOTSTRAP_ADMIN_PASSWORD, com
   a senha salva em hash bcrypt (nunca em texto puro). Só insere se esse
   username ainda não existir — não sobrescreve usuário existente.

Seguro rodar quantas vezes for preciso (idempotente nos dois passos).

Uso:
    # PowerShell
    $env:BOOTSTRAP_ADMIN_USER = "admin"
    $env:BOOTSTRAP_ADMIN_PASSWORD = "uma-senha-forte"
    python backend/scripts/bootstrap_db.py

Se as variáveis de ambiente não forem definidas, o script cria/confirma a
tabela normalmente e apenas pula a etapa de seed (com um aviso), sem erro.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import bcrypt  # noqa: E402

from app.db.database_mysql import get_cursor  # noqa: E402

# Schema confirmado no MySQL de produção via `DESCRIBE usuarios` antes de
# escrever este script (é seguro: apenas leitura). Os valores do ENUM `role`
# são os mesmos aceitos por app.schemas.users.UserCreate.role
# (Literal["Gestor", "Técnico", "Jovem Aprendiz"]). Isso diverge um pouco da
# referência legada (`role VARCHAR(50) NOT NULL`) do inicializar_db.py que
# está sendo removido — o schema real em produção já usa ENUM, então um
# deploy limpo deve replicar o schema real, não o texto do script antigo.
CREATE_USUARIOS_SQL = """
CREATE TABLE IF NOT EXISTS usuarios (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(100) NOT NULL UNIQUE,
    password VARCHAR(255) NOT NULL,
    role ENUM('Gestor','Técnico','Jovem Aprendiz') NOT NULL
)
"""


def create_usuarios_table() -> None:
    with get_cursor(dict_cursor=False) as cur:
        cur.execute(CREATE_USUARIOS_SQL)
    print("Tabela 'usuarios' OK (criada agora ou já existia).")


def seed_admin_gestor() -> None:
    username = os.environ.get("BOOTSTRAP_ADMIN_USER")
    password = os.environ.get("BOOTSTRAP_ADMIN_PASSWORD")

    if not username or not password:
        print(
            "BOOTSTRAP_ADMIN_USER / BOOTSTRAP_ADMIN_PASSWORD não definidos no "
            "ambiente — nenhum usuário será semeado. Isso é esperado se o "
            "banco já tem usuários ou se o primeiro Gestor será criado "
            "manualmente por outro meio."
        )
        return

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT id FROM usuarios WHERE username = %s", (username,))
        existing = cur.fetchone()

    if existing:
        print(f"Usuário '{username}' já existe (id={existing['id']}) — nada a fazer, não sobrescrevo.")
        return

    hashed = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    with get_cursor(dict_cursor=False) as cur:
        cur.execute(
            "INSERT INTO usuarios (username, password, role) VALUES (%s, %s, 'Gestor')",
            (username, hashed),
        )
    print(f"Usuário Gestor '{username}' criado com sucesso.")


def main():
    create_usuarios_table()
    seed_admin_gestor()


if __name__ == "__main__":
    main()
