"""Migração 0801: tabelas do módulo de acessos e carga inicial de sistemas."""
import re

import pymysql
import pytest

from tests.acessos_helpers import contar, criar_acesso, criar_colaborador, criar_sistema

pytestmark = pytest.mark.integration

TABELAS = ("sistemas", "pacotes_funcao", "colaborador_acessos", "acesso_eventos", "acesso_tarefas")


def _tabelas_existentes() -> set:
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT table_name AS nome FROM information_schema.tables "
            "WHERE table_schema = DATABASE()"
        )
        return {linha["nome"] for linha in cur.fetchall()}


def _executar_arquivo(nome: str) -> None:
    from app.db.database_mysql import get_cursor
    from app.db.migrations import MIGRATIONS_DIR, executar_sql as executar_migracao

    texto = (MIGRATIONS_DIR / nome).read_text(encoding="utf-8")
    with get_cursor() as cur:
        executar_migracao(cur, texto)


def test_tabelas_de_acessos_existem(limpar_banco):
    assert set(TABELAS) <= _tabelas_existentes()


def test_nenhuma_coluna_de_senha(limpar_banco):
    from app.db.database_mysql import get_cursor

    padrao = re.compile(r"senha|pass|pwd", re.I)
    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT table_name AS tabela, column_name AS coluna FROM information_schema.columns "
            "WHERE table_schema = DATABASE() AND table_name IN %s",
            (TABELAS,),
        )
        colunas = cur.fetchall()
    assert colunas, "as tabelas de acessos deveriam ter colunas"
    assert [c for c in colunas if padrao.search(c["coluna"])] == []


def test_carga_inicial_de_sistemas_idempotente(limpar_banco):
    from app.db.database_mysql import get_cursor

    _executar_arquivo("0801_acessos.sql")
    _executar_arquivo("0801_acessos.sql")
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT nome, ativo FROM sistemas ORDER BY nome")
        linhas = [(r["nome"], r["ativo"]) for r in cur.fetchall()]
    assert linhas == [
        ("AD", 1), ("Biz Deliveries", 1), ("E-mail", 1), ("Promax", 1), ("VPN", 1), ("WMS", 1),
    ]


def test_down_remove_e_up_recria(limpar_banco):
    _executar_arquivo("0801_acessos.down.sql")
    assert not (set(TABELAS) & _tabelas_existentes())
    _executar_arquivo("0801_acessos.sql")
    assert set(TABELAS) <= _tabelas_existentes()
    assert contar("sistemas") == 6


def test_login_unico_ignora_caixa(limpar_banco):
    c = criar_colaborador()
    s = criar_sistema("AD")
    criar_acesso(c, s, "Maria.Silva")
    with pytest.raises(pymysql.err.IntegrityError) as exc:
        criar_acesso(c, s, "maria.silva")
    assert exc.value.args[0] == 1062
