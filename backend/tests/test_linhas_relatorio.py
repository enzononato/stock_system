"""Relatório de custo mensal por unidade e exportação (F5)."""
from io import BytesIO

import openpyxl
import pytest

from tests.fixtures_linhas import *  # noqa: F401,F403

pytestmark = pytest.mark.integration


@pytest.fixture
def cenario(client_gestor, nova_linha, criar_unidade):
    """A: ativas 49.90 + 30.10 + 0.05, suspensa 100.00 e cancelada 200.00; B: ativa 15.55; C: só suspensa."""
    a = criar_unidade(nome="Unidade A")
    b = criar_unidade(nome="Unidade B")
    c = criar_unidade(nome="Unidade C")

    def status(linha, novo):
        resp = client_gestor.post(f"/api/linhas/{linha['id']}/status", json={"status": novo})
        assert resp.status_code == 200, resp.text

    for custo in ("49.90", "30.10", "0.05"):
        nova_linha(unidade_id=a["id"], custo_mensal=custo)
    status(nova_linha(unidade_id=a["id"], custo_mensal="100.00"), "suspensa")
    status(nova_linha(unidade_id=a["id"], custo_mensal="200.00"), "cancelada")
    nova_linha(unidade_id=b["id"], custo_mensal="15.55")
    status(nova_linha(unidade_id=c["id"], custo_mensal="10.00"), "suspensa")
    return a, b, c


def test_relatorio_soma_por_unidade_so_ativas(client_gestor, cenario):
    a, b, _ = cenario

    resp = client_gestor.get("/api/linhas/relatorio/custo")

    assert resp.status_code == 200, resp.text
    corpo = resp.json()
    assert corpo["unidades"] == [
        {"unidade_id": a["id"], "unidade_nome": "Unidade A", "quantidade": 3, "custo_total": "80.05"},
        {"unidade_id": b["id"], "unidade_nome": "Unidade B", "quantidade": 1, "custo_total": "15.55"},
    ]
    assert corpo["total_linhas"] == 4
    assert corpo["custo_total"] == "95.60"
    assert corpo["gerado_em"]


def test_relatorio_sem_linhas(client_gestor):
    corpo = client_gestor.get("/api/linhas/relatorio/custo").json()

    assert corpo["unidades"] == []
    assert corpo["total_linhas"] == 0
    assert corpo["custo_total"] == "0.00"


def test_exportar_excel(client_gestor, cenario):
    resp = client_gestor.get("/api/linhas/relatorio/custo/export", params={"formato": "xlsx"})

    assert resp.status_code == 200, resp.text
    assert resp.headers["content-type"] == (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert "custo_linhas_" in resp.headers["content-disposition"]
    ws = openpyxl.load_workbook(BytesIO(resp.content)).active
    linhas = {row[0]: row for row in ws.iter_rows(values_only=True) if row and row[0]}
    assert linhas["UNIDADE A"][1] == 3
    assert float(linhas["UNIDADE A"][2]) == pytest.approx(80.05)
    assert float(linhas["UNIDADE B"][2]) == pytest.approx(15.55)
    assert "UNIDADE C" not in linhas
    assert linhas["TOTAL"][1] == 4
    assert float(linhas["TOTAL"][2]) == pytest.approx(95.60)


def test_exportar_pdf(client_gestor, cenario):
    resp = client_gestor.get("/api/linhas/relatorio/custo/export", params={"formato": "pdf"})

    assert resp.status_code == 200, resp.text
    assert resp.content[:4] == b"%PDF"
    assert resp.headers["content-type"] == "application/pdf"


def test_formato_padrao_e_excel(client_gestor, cenario):
    resp = client_gestor.get("/api/linhas/relatorio/custo/export")

    assert resp.status_code == 200
    assert resp.headers["content-type"].endswith("spreadsheetml.sheet")


def test_formato_invalido_422(client_gestor):
    resp = client_gestor.get("/api/linhas/relatorio/custo/export", params={"formato": "csv"})

    assert resp.status_code == 422
    assert resp.json()["detail"] == "Formato inválido: use xlsx ou pdf."


def test_relatorio_exige_gestor_ou_tecnico(client_aprendiz, client_tecnico):
    assert client_aprendiz.get("/api/linhas/relatorio/custo").status_code == 403
    assert client_aprendiz.get("/api/linhas/relatorio/custo/export").status_code == 403
    assert client_tecnico.get("/api/linhas/relatorio/custo").status_code == 200
    assert client_tecnico.get("/api/linhas/relatorio/custo/export").status_code == 200
