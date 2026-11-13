"""Migrações da F4 (0401-0405): colunas de compra, status Em Manutenção e as
tabelas manutencoes, transferencias e baixas_apagamento."""
from pathlib import Path

import pymysql
import pytest

# `_test_schema` garante o banco (e as migrações já aplicadas) mesmo quando o teste não pede fixture.
pytestmark = [pytest.mark.integration, pytest.mark.usefixtures("_test_schema")]

PASTA = Path(__file__).resolve().parent.parent / "migrations"
COLUNAS_COMPRA = ("valor_compra", "data_compra", "garantia_ate", "nf_anexo_key")
TABELAS = ("manutencoes", "transferencias", "baixas_apagamento")
ENUM_COM = "enum('Disponível','Indisponível','Pendente','Pendente Devolução','Em Manutenção')"
ENUM_SEM = "enum('Disponível','Indisponível','Pendente','Pendente Devolução')"


def _rodar(nome: str) -> None:
    from app.db import migrations as mig
    from app.db.database_mysql import get_cursor

    with get_cursor() as cur:
        mig.executar_sql(cur, (PASTA / nome).read_text(encoding="utf-8"))


def _colunas_de_items() -> dict:
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT column_name AS nome, column_type AS tipo FROM information_schema.columns "
            "WHERE table_schema = DATABASE() AND table_name = 'items'"
        )
        return {r["nome"]: r["tipo"].lower() if isinstance(r["tipo"], str) else r["tipo"].decode().lower()
                for r in cur.fetchall()}


def _tabelas() -> set:
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT table_name AS nome FROM information_schema.tables WHERE table_schema = DATABASE()"
        )
        return {r["nome"] for r in cur.fetchall()}


def _tipo_status() -> str:
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT column_type AS tipo FROM information_schema.columns "
            "WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'status'"
        )
        tipo = cur.fetchone()["tipo"]
        return tipo if isinstance(tipo, str) else tipo.decode()


def test_colunas_de_compra_existem():
    cols = _colunas_de_items()
    assert cols["valor_compra"] == "decimal(12,2)"
    assert cols["data_compra"] == "date"
    assert cols["garantia_ate"] == "date"
    assert cols["nf_anexo_key"] == "varchar(255)"


def test_status_aceita_em_manutencao():
    assert _tipo_status() == ENUM_COM


def test_down_do_status_recusa_com_item_em_manutencao(criar_item, executar_sql):
    item = criar_item()
    executar_sql("UPDATE items SET status='Em Manutenção' WHERE id=%s", (item["id"],))
    try:
        with pytest.raises(pymysql.MySQLError) as exc:
            _rodar("0402_items_status_manutencao.down.sql")
        assert exc.value.args[0] == 1242
        assert "'Em Manutenção'" in _tipo_status()
        from app.db.database_mysql import get_cursor

        with get_cursor(commit=False) as cur:
            cur.execute("SELECT status FROM items WHERE id=%s", (item["id"],))
            assert cur.fetchone()["status"] == "Em Manutenção"
    finally:
        executar_sql("UPDATE items SET status='Disponível' WHERE id=%s", (item["id"],))


def test_down_e_up_do_status_sem_item_em_manutencao():
    try:
        _rodar("0402_items_status_manutencao.down.sql")
        assert _tipo_status() == ENUM_SEM
    finally:
        _rodar("0402_items_status_manutencao.sql")
    assert "'Em Manutenção'" in _tipo_status()


def test_ida_e_volta_das_migracoes_f4(criar_item):
    item = criar_item()  # schema atual com dados
    try:
        for nome in ("0405_baixas_apagamento", "0404_transferencias", "0403_manutencoes", "0401_items_compra"):
            _rodar(f"{nome}.down.sql")
        assert not (set(TABELAS) & _tabelas())
        assert not (set(COLUNAS_COMPRA) & set(_colunas_de_items()))
    finally:
        for _ in range(2):  # idempotência
            for nome in ("0401_items_compra", "0403_manutencoes", "0404_transferencias", "0405_baixas_apagamento"):
                _rodar(f"{nome}.sql")
    assert set(TABELAS) <= _tabelas()
    assert set(COLUNAS_COMPRA) <= set(_colunas_de_items())
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT valor_compra FROM items WHERE id=%s", (item["id"],))
        row = cur.fetchone()
    assert row is not None and row["valor_compra"] is None
