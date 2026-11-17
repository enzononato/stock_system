"""Migração 0600: tabelas do inventário físico."""
from pathlib import Path

import pytest

pytestmark = pytest.mark.integration

MIGRACOES = Path(__file__).resolve().parent.parent / "migrations"

COLUNAS = {
    "inventarios": {
        "id", "unidade_id", "status", "iniciado_por", "iniciado_em", "encerrado_por",
        "encerrado_em", "observacao", "motivo_cancelamento", "unidade_aberta",
    },
    "inventario_leituras": {
        "id", "inventario_id", "item_id", "chave", "codigo_lido", "lido_por", "lido_em",
        "local_id_encontrado", "usuario_conferido", "colaborador_id_encontrado",
        "revenda_item", "local_id_esperado", "colaborador_id_esperado", "usuario_esperado",
        "resultado", "detalhe", "observacao", "atualizado_em",
    },
    "inventario_nao_encontrados": {
        "inventario_id", "item_id", "status_item", "local_id", "colaborador_id", "usuario",
    },
}


def _inserir_inventario(executar_sql, unidade_id, status="aberto"):
    executar_sql(
        "INSERT INTO inventarios (unidade_id, status, iniciado_por, iniciado_em) "
        "VALUES (%s, %s, 'teste', NOW())",
        (unidade_id, status),
    )


def test_tabelas_e_colunas_existem(_test_schema):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT table_name AS t, column_name AS c FROM information_schema.columns "
            "WHERE table_schema = DATABASE() AND table_name LIKE 'inventario%%'"
        )
        achadas: dict[str, set] = {}
        for linha in cur.fetchall():
            achadas.setdefault(linha["t"], set()).add(linha["c"])
    for tabela, colunas in COLUNAS.items():
        assert colunas <= achadas.get(tabela, set()), tabela


def test_um_inventario_aberto_por_unidade(unidade, executar_sql):
    import pymysql

    _inserir_inventario(executar_sql, unidade["id"])
    with pytest.raises(pymysql.err.IntegrityError):
        _inserir_inventario(executar_sql, unidade["id"])
    executar_sql("UPDATE inventarios SET status='encerrado' WHERE unidade_id=%s", (unidade["id"],))
    _inserir_inventario(executar_sql, unidade["id"])  # agora passa


def test_leitura_unica_por_chave_e_por_item(unidade, executar_sql):
    import pymysql

    _inserir_inventario(executar_sql, unidade["id"])

    def ler(chave):
        executar_sql(
            "INSERT INTO inventario_leituras (inventario_id, item_id, chave, codigo_lido, "
            "lido_por, lido_em, resultado) "
            "SELECT id, NULL, %s, 'x', 'teste', NOW(), 'nao_cadastrado' FROM inventarios "
            "WHERE unidade_id=%s",
            (chave, unidade["id"]),
        )

    ler("codigo:A")
    with pytest.raises(pymysql.err.IntegrityError):
        ler("codigo:A")
    ler("codigo:B")  # item_id NULL repetido não conflita


def test_down_e_up(_test_schema):
    from app.db.database_mysql import get_cursor
    from app.db.migrations import executar_sql

    down = (MIGRACOES / "0600_inventarios.down.sql").read_text(encoding="utf-8")
    up = (MIGRACOES / "0600_inventarios.sql").read_text(encoding="utf-8")
    try:
        with get_cursor() as cur:
            executar_sql(cur, down)
    finally:
        with get_cursor() as cur:
            executar_sql(cur, up)
    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT COUNT(*) AS n FROM information_schema.tables "
            "WHERE table_schema = DATABASE() AND table_name IN "
            "('inventarios','inventario_leituras','inventario_nao_encontrados')"
        )
        assert cur.fetchone()["n"] == 3
