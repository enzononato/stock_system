"""
Fixture da base fictícia de RH (Postgres) usada pelos testes da fonte de colaboradores.

Cria, no Postgres compartilhado de desenvolvimento, um database temporário
`colab_test_<sufixo>` com a tabela `"colaboradores-revalle"` (mesmos nomes de coluna
entre aspas da base real, com acento e maiúsculas), aponta `settings.COLABORADORES_DATABASE_URL`
para ele e o derruba (`DROP DATABASE ... WITH (FORCE)`) ao final. Sem
`COLAB_TEST_PG_ADMIN_URL` o teste é pulado -- nunca falha por falta de infraestrutura.
"""
import os
import re

import pytest

# CPFs com dígitos verificadores válidos.
CPF_ATIVO = "52998224725"
CPF_DESLIGADO = "11144477735"
CPF_AFASTADO = "16899535009"
CPF_MASCARADO = "93541134780"
CPF_MASCARADO_FORMATADO = "935.411.347-80"

_DDL = """
CREATE TABLE "colaboradores-revalle" (
    "CPF" text,
    "Nome Completo" text,
    "Data Nascimento" date,
    "Cargo" text,
    "Setor" text,
    "Unidade" text,
    "Matrícula" text,
    "E-mail" text,
    "Celular" text,
    "Situação" text,
    "Gestor" text
)
"""

_PESSOAS = [
    (CPF_ATIVO, "Maria da Silva", "1990-05-17", "Analista", "TI", "Revalle Juazeiro",
     "1001", "maria@revalle.com.br", "(74) 99999-0001", "Ativo", "Carlos Souza"),
    (CPF_DESLIGADO, "João Pereira", "1985-01-02", "Vendedor", "Comercial", "Revalle Juazeiro",
     "1002", "joao@revalle.com.br", "(74) 99999-0002", "Desligado", "Carlos Souza"),
    (CPF_AFASTADO, "Ana Lima", "1992-11-30", "Auxiliar", "Logística", "Revalle Petrolina",
     "1003", "ana@revalle.com.br", "(87) 99999-0003", "Afastado (licença)", None),
    # CPF com máscara na origem, sem situação preenchida
    (CPF_MASCARADO_FORMATADO, "Pedro Alves", "1978-07-08", "Gerente", "Frota", "Revalle Petrolina",
     "1004", "pedro@revalle.com.br", "(87) 99999-0004", None, "Ana Lima"),
]


@pytest.fixture
def base_rh(monkeypatch):
    admin_url = (os.environ.get("COLAB_TEST_PG_ADMIN_URL") or "").strip()
    if not admin_url:
        pytest.skip("COLAB_TEST_PG_ADMIN_URL não definido: sem Postgres para a base fictícia de RH.")

    import psycopg
    from psycopg.conninfo import make_conninfo

    sufixo = re.sub(r"[^a-z0-9_]", "_", (os.environ.get("COLAB_TEST_DB_SUFFIX") or "local").lower())
    nome_db = f"colab_test_{sufixo}"

    try:
        admin = psycopg.connect(admin_url, autocommit=True, connect_timeout=5)
    except Exception as e:  # noqa: BLE001
        # Variável definida mas Postgres fora do ar: é falha de ambiente que não pode passar despercebida.
        pytest.fail(
            "COLAB_TEST_PG_ADMIN_URL está definido, mas o Postgres de teste está inacessível "
            f"({type(e).__name__}). Suba o dev-postgres (localhost:5433) ou remova a variável.",
            pytrace=False,
        )

    try:
        admin.execute(f'DROP DATABASE IF EXISTS "{nome_db}" WITH (FORCE)')
        admin.execute(f'CREATE DATABASE "{nome_db}"')
    except Exception:
        admin.close()
        raise

    url_db = make_conninfo(admin_url, dbname=nome_db)

    def executar(sql, params=None):
        with psycopg.connect(url_db, autocommit=True) as conn:
            conn.execute(sql, params)

    try:
        executar(_DDL)
        for p in _PESSOAS:
            executar(
                'INSERT INTO "colaboradores-revalle" ("CPF", "Nome Completo", "Data Nascimento", '
                '"Cargo", "Setor", "Unidade", "Matrícula", "E-mail", "Celular", "Situação", "Gestor") '
                "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                p,
            )

        from app.core import config as _config

        settings = _config.settings  # lido na hora: test_auth recarrega app.core.config

        monkeypatch.setattr(settings, "COLABORADORES_DATABASE_URL", url_db)
        monkeypatch.setattr(settings, "COLABORADORES_TABLE", "colaboradores-revalle")
        monkeypatch.setattr(settings, "COLABORADORES_SCHEMA", "public")
        monkeypatch.setattr(settings, "COLABORADORES_COLUMNS", "")
        yield executar
    finally:
        try:
            admin.execute(f'DROP DATABASE IF EXISTS "{nome_db}" WITH (FORCE)')
        finally:
            admin.close()
