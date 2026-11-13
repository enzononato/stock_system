"""Tabela `termo_solicitacoes` e `TermoDBManager`."""
from datetime import timedelta

import pytest

from tests.fixtures_termos import colaborador, criar_colaborador  # noqa: F401

pytestmark = pytest.mark.integration


@pytest.fixture
def termos(_test_schema, limpar_banco):
    from app.db.termo_db import TermoDBManager

    return TermoDBManager()


def _criar(termos, item_id, *, tipo="entrega", status="manual", colaborador_id=None, **extra):
    from app.core.tempo import agora_local

    return termos.criar(
        item_id=item_id, tipo=tipo, status=status, solicitado_por="gestor.teste",
        solicitado_em=agora_local(), colaborador_id=colaborador_id, **extra,
    )


def test_migracao_cria_tabela_com_enums(_test_schema, executar_sql):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT column_name AS nome, column_type AS tipo FROM information_schema.columns "
            "WHERE table_schema = DATABASE() AND table_name = 'termo_solicitacoes'"
        )
        colunas = {r["nome"]: r["tipo"] for r in cur.fetchall()}
    assert colunas["status"] == (
        "enum('agendado','enviado','assinado','recusado','expirado','cancelado','erro','manual')"
    )
    assert colunas["tipo"] == "enum('entrega','devolucao','uso')"


def test_criar_obter_e_atualizar(termos, criar_item, colaborador):
    item = criar_item()
    termo_id = _criar(termos, item["id"], colaborador_id=colaborador["id"])

    linha = termos.obter(termo_id)
    assert linha["colaborador_nome"] == "Maria Souza"
    assert linha["item_descricao"] == "Notebook Dell Latitude 5420"

    assert termos.atualizar(termo_id, status="enviado", revsign_documento_id="1001") is True
    assert termos.obter_por_documento("1001")["id"] == termo_id
    assert termos.atualizar(termo_id, somente_se_status=("agendado",), status="cancelado") is False
    assert termos.obter(termo_id)["status"] == "enviado"
    with pytest.raises(ValueError):
        termos.atualizar(termo_id, item_id=2)


def test_pendentes_do_item_e_por_tipo(termos, criar_item):
    item = criar_item()
    manual = _criar(termos, item["id"], status="manual", tipo="entrega")
    _criar(termos, item["id"], status="assinado", tipo="entrega")
    enviado = _criar(termos, item["id"], status="enviado", tipo="uso")

    assert [t["id"] for t in termos.pendentes_do_item(item["id"])] == [enviado, manual]
    assert [t["id"] for t in termos.pendentes_do_item(item["id"], tipo="uso")] == [enviado]


def test_vigente_ignora_cancelada_e_item_disponivel(termos, criar_item, executar_sql):
    item = criar_item()
    assert termos.vigente(item["id"]) is None

    executar_sql("UPDATE items SET status='Indisponível' WHERE id=%s", (item["id"],))
    assinada = _criar(termos, item["id"], status="assinado")
    _criar(termos, item["id"], status="cancelado")
    vigente = termos.vigente(item["id"])
    assert vigente["id"] == assinada
    assert vigente["legado"] is False

    executar_sql("UPDATE items SET status='Disponível' WHERE id=%s", (item["id"],))
    assert termos.vigente(item["id"]) is None


def test_vigente_legado_vem_do_historico(termos, item_indisponivel):
    vigente = termos.vigente(item_indisponivel["id"])
    assert vigente["legado"] is True
    assert vigente["id"] is None
    assert vigente["status"] == "assinado"
    assert vigente["arquivo_assinado_key"] == "termos_assinados/fixture.pdf"
    assert vigente["solicitado_por"] == "fixture_teste"


def test_vigente_legado_ignora_confirmacao_de_emprestimo_anterior(
    termos, inv_manager, item_indisponivel, forcar_devolucao_iniciada
):
    """Um segundo empréstimo ainda pendente não herda o termo do ciclo anterior."""
    from datetime import datetime

    item_id = item_indisponivel["id"]
    forcar_devolucao_iniciada(item_id)
    assert inv_manager.confirm_return(item_id, "fixture_teste", "termos_devolucao_assinados/x.pdf")[0]
    ok, msg = inv_manager.issue(
        item_id, "Outro", "11144477735", "101 - Puxada", "Analista", "TI", "Revalle Juazeiro",
        datetime.now().strftime("%d/%m/%Y"), "fixture_teste",
    )
    assert ok, msg
    assert termos.vigente(item_id) is None


def test_agendados_vencidos_e_enviados(termos, criar_item):
    from app.core.tempo import agora_local

    item = criar_item()
    agora = agora_local()
    agendado = _criar(termos, item["id"], status="agendado", agendado_para=agora + timedelta(hours=1))
    assert termos.agendados_vencidos(agora) == []
    assert [t["id"] for t in termos.agendados_vencidos(agora + timedelta(minutes=61))] == [agendado]

    com_documento = _criar(termos, item["id"], status="enviado", tipo="uso")
    termos.atualizar(com_documento, revsign_documento_id="2001")
    _criar(termos, item["id"], status="enviado", tipo="devolucao")  # sem documento
    assert [t["id"] for t in termos.enviados()] == [com_documento]


def test_listar_por_colaborador_pagina(termos, criar_item, colaborador):
    item = criar_item()
    for _ in range(3):
        _criar(termos, item["id"], colaborador_id=colaborador["id"])
    linhas, total = termos.listar_por_colaborador(colaborador["id"], limit=2, offset=0)
    assert len(linhas) == 2 and total == 3
    linhas, total = termos.listar_por_item(item["id"], limit=10, offset=2)
    assert len(linhas) == 1 and total == 3


def test_migracao_0300_vai_e_volta(_test_schema):
    """O `.down.sql` derruba a tabela e reaplicar o `.sql` a recria (idempotente)."""
    from app.db import migrations
    from app.db.database_mysql import get_cursor

    por_versao = {m.versao: m for m in migrations.listar_migracoes()}
    mig = por_versao[300]
    assert mig.down_path is not None

    def existe() -> bool:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT COUNT(*) AS n FROM information_schema.tables "
                "WHERE table_schema = DATABASE() AND table_name = 'termo_solicitacoes'"
            )
            return cur.fetchone()["n"] == 1

    assert existe()
    try:
        with get_cursor() as cur:
            migrations.executar_sql(cur, mig.down_path.read_text(encoding="utf-8"))
        assert not existe()
    finally:
        with get_cursor() as cur:
            migrations.executar_sql(cur, mig.up_path.read_text(encoding="utf-8"))
    assert existe()
