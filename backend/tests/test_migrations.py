"""
Testes do executor de migrações versionadas (app/db/migrations.py).

Os testes de integração usam o MySQL de teste (fixture `_test_schema`) e, quando
precisam de arquivos de migração, criam-nos num diretório temporário apontado via
monkeypatch de `MIGRATIONS_DIR` -- nunca dependem dos arquivos reais de
backend/migrations/, que crescem a cada fase.
"""
import pytest

from app.core.config import Settings



def _escrever(pasta, nome, texto=""):
    caminho = pasta / nome
    caminho.write_text(texto, encoding="utf-8")
    return caminho


def _versoes_aplicadas():
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT versao FROM schema_migrations ORDER BY versao")
        return [r["versao"] for r in cur.fetchall()]


@pytest.mark.integration
def test_listar_migracoes_ordena_e_pareia_down(tmp_path, monkeypatch):
    from app.db import migrations

    _escrever(tmp_path, "0002_b.sql")
    _escrever(tmp_path, "0001_a.sql")
    _escrever(tmp_path, "0001_a.down.sql")
    _escrever(tmp_path, "README.md")  # arquivos fora da convenção são ignorados
    monkeypatch.setattr(migrations, "MIGRATIONS_DIR", tmp_path)

    lista = migrations.listar_migracoes()

    assert [m.versao for m in lista] == [1, 2]
    assert [m.nome for m in lista] == ["a", "b"]
    assert lista[0].down_path == tmp_path / "0001_a.down.sql"
    assert lista[1].down_path is None


@pytest.mark.integration
def test_aplicar_pendentes_e_idempotente(_test_schema):
    from app.db import migrations

    migrations.garantir_baseline()
    migrations.aplicar_pendentes()
    assert migrations.aplicar_pendentes() == []

    esperadas = [m.versao for m in migrations.listar_migracoes()]
    assert set(esperadas) <= set(_versoes_aplicadas())


@pytest.mark.integration
def test_aplicar_pendentes_aplica_versao_faltante_mesmo_com_maior_ja_aplicada(
    _test_schema, tmp_path, monkeypatch
):
    from app.db import migrations
    from app.db.database_mysql import get_cursor

    _escrever(tmp_path, "9001_a.sql", "CREATE TABLE tmp_mig_a (id INT PRIMARY KEY);")
    _escrever(tmp_path, "9001_a.down.sql", "DROP TABLE IF EXISTS tmp_mig_a;")
    _escrever(tmp_path, "9002_b.sql", "CREATE TABLE tmp_mig_b (id INT PRIMARY KEY);")
    _escrever(tmp_path, "9002_b.down.sql", "DROP TABLE IF EXISTS tmp_mig_b;")
    monkeypatch.setattr(migrations, "MIGRATIONS_DIR", tmp_path)
    migrations.aplicar_pendentes()  # garante schema_migrations; aplica 9001 e 9002
    try:
        with get_cursor() as cur:
            cur.execute("DELETE FROM schema_migrations WHERE versao=9001")
            cur.execute("DROP TABLE tmp_mig_a")

        assert migrations.aplicar_pendentes() == [9001]
    finally:
        with get_cursor() as cur:
            cur.execute("DROP TABLE IF EXISTS tmp_mig_a")
            cur.execute("DROP TABLE IF EXISTS tmp_mig_b")
            cur.execute("DELETE FROM schema_migrations WHERE versao IN (9001, 9002)")


@pytest.mark.integration
def test_reverter_ultima_e_reaplicar(_test_schema, tmp_path, monkeypatch):
    from app.db import migrations
    from app.db.database_mysql import get_cursor

    # Migração descartável com número alto: é sempre a "última" aplicada e não
    # depende de quais migrações reais existem em backend/migrations/.
    _escrever(tmp_path, "9001_descartavel.sql", "CREATE TABLE tmp_mig_descartavel (id INT PRIMARY KEY);")
    _escrever(tmp_path, "9001_descartavel.down.sql", "DROP TABLE IF EXISTS tmp_mig_descartavel;")
    monkeypatch.setattr(migrations, "MIGRATIONS_DIR", tmp_path)
    try:
        assert migrations.aplicar_pendentes() == [9001]

        v = migrations.reverter_ultima()

        assert v == 9001
        assert v not in _versoes_aplicadas()
        assert migrations.aplicar_pendentes() == [v]
    finally:
        with get_cursor() as cur:
            cur.execute("DROP TABLE IF EXISTS tmp_mig_descartavel")
            cur.execute("DELETE FROM schema_migrations WHERE versao=9001")


@pytest.mark.integration
def test_reverter_ultima_sem_down_levanta_erro(_test_schema, tmp_path, monkeypatch):
    from app.db import migrations
    from app.db.database_mysql import get_cursor

    _escrever(tmp_path, "9001_sem_down.sql", "CREATE TABLE tmp_mig_sem_down (id INT PRIMARY KEY);")
    monkeypatch.setattr(migrations, "MIGRATIONS_DIR", tmp_path)
    try:
        migrations.aplicar_pendentes()
        with pytest.raises(RuntimeError, match="9001"):
            migrations.reverter_ultima()
        assert 9001 in _versoes_aplicadas()
    finally:
        with get_cursor() as cur:
            cur.execute("DROP TABLE IF EXISTS tmp_mig_sem_down")
            cur.execute("DELETE FROM schema_migrations WHERE versao=9001")


@pytest.mark.integration
def test_executar_sql_divide_comandos(_test_schema):
    from app.db.database_mysql import get_cursor
    from app.db.migrations import executar_sql

    texto = """
    -- comentário inicial; com ponto e vírgula dentro
    CREATE TEMPORARY TABLE tmp_exec_a (
        id INT PRIMARY KEY,
        nome VARCHAR(10)
    );

    -- outro comentário
    CREATE TEMPORARY TABLE tmp_exec_b (id INT PRIMARY KEY);
    """
    with get_cursor() as cur:
        executar_sql(cur, texto)
        cur.execute("INSERT INTO tmp_exec_a (id, nome) VALUES (1, 'x')")
        cur.execute("INSERT INTO tmp_exec_b (id) VALUES (1)")
        cur.execute("SELECT COUNT(*) AS n FROM tmp_exec_a")
        assert cur.fetchone()["n"] == 1
        cur.execute("SELECT COUNT(*) AS n FROM tmp_exec_b")
        assert cur.fetchone()["n"] == 1
        cur.execute("DROP TEMPORARY TABLE tmp_exec_a")
        cur.execute("DROP TEMPORARY TABLE tmp_exec_b")


@pytest.mark.unit
def test_settings_padroes_novos():
    s = Settings(
        _env_file=None,
        DB_HOST="h",
        DB_USER="u",
        DB_PASSWORD="p",
        DB_NAME="n",
        JWT_SECRET="x" * 32,
        SMTP_USER="naoresponda@exemplo.com",
        SMTP_FROM="",
    )

    assert s.ALERTA_GARANTIA_DIAS == 30
    assert s.smtp_from_effective == "naoresponda@exemplo.com"
    s2 = Settings(
        _env_file=None, DB_HOST="h", DB_USER="u", DB_PASSWORD="p", DB_NAME="n",
        JWT_SECRET="x" * 32, SMTP_USER="a@x.com", SMTP_FROM="Patrimônio <b@x.com>",
    )
    assert s2.smtp_from_effective == "Patrimônio <b@x.com>"
