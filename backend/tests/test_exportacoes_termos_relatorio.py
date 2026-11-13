"""F7 / Task 6: exportações de Termos pendentes e do Relatório mensal (Excel e PDF)."""
from datetime import datetime, timedelta
from io import BytesIO

import openpyxl
import pytest

from tests.fixtures_termos import criar_colaborador  # noqa: F401

pytestmark = pytest.mark.integration


def _planilha(resp) -> list[list]:
    assert resp.status_code == 200, resp.text
    ws = openpyxl.load_workbook(BytesIO(resp.content)).active
    return [list(linha) for linha in ws.iter_rows(min_row=5, values_only=True)]


def _ws(resp):
    return openpyxl.load_workbook(BytesIO(resp.content)).active


def _cabecalho(resp) -> list[str]:
    return [c.value for c in _ws(resp)[4]]


def _celulas(resp) -> list[str]:
    return [str(c) for linha in _ws(resp).iter_rows(values_only=True) for c in linha if c is not None]


def _termo(item_id, *, status="manual", ha=timedelta(days=4), **extra):
    from app.core.tempo import agora_local
    from app.db.termo_db import TermoDBManager

    return TermoDBManager().criar(
        item_id=item_id, tipo="entrega", status=status, solicitado_por="gestor.teste",
        solicitado_em=agora_local() - ha, **extra,
    )


# ── Termos pendentes ────────────────────────────────────────────────────────

@pytest.fixture
def cenario_termos(criar_item, criar_colaborador, item_pendente, executar_sql):
    colab = criar_colaborador()
    manual, enviado, assinado = criar_item(model="Modelo-Manual"), criar_item(), criar_item()
    _termo(manual["id"], ha=timedelta(days=5, hours=1), colaborador_id=colab["id"])
    _termo(
        enviado["id"], status="enviado", ha=timedelta(days=3, hours=1), colaborador_id=colab["id"],
        destinatario_email="maria.souza@gmail.com",
    )
    _termo(assinado["id"], status="assinado", ha=timedelta(days=6))
    executar_sql(
        "UPDATE history SET data_operacao=%s WHERE item_id=%s AND operation='Empréstimo'",
        (datetime.now() - timedelta(days=10, hours=1), item_pendente["id"]),
    )
    return {"manual": manual, "enviado": enviado, "legado": item_pendente}


def test_termos_pendentes_excel(client_gestor, cenario_termos):
    resp = client_gestor.get("/api/exportacoes/termos-pendentes", params={"formato": "xlsx"})
    linhas = _planilha(resp)
    cab = _cabecalho(resp)
    assert len(linhas) == 3
    assert not any("@" in c for c in _celulas(resp))
    legado, manual, enviado = linhas  # da mais antiga para a mais nova
    assert legado[cab.index("Solicitação")] == "—"
    assert legado[cab.index("Status")] == "Sem solicitação"
    assert legado[cab.index("Termo")] == "Entrega"
    assert legado[cab.index("Dias pendente")] == 10
    assert manual[cab.index("Status")] == "Assinatura manual"
    assert manual[cab.index("Dias pendente")] == 5
    assert manual[cab.index("Colaborador")] == "Maria Souza"
    assert manual[cab.index("Solicitado por")] == "gestor.teste"
    assert enviado[cab.index("Status")] == "Enviado"
    assert isinstance(manual[cab.index("Solicitação")], int)
    assert cab[-1] == "Motivo"
    assert [l[cab.index("Unidade")] for l in linhas] == ["REVALLE JUAZEIRO"] * 3


def test_termos_pendentes_busca(client_gestor, cenario_termos):
    resp = client_gestor.get(
        "/api/exportacoes/termos-pendentes", params={"formato": "xlsx", "busca": "Modelo-Manual"}
    )
    linhas = _planilha(resp)
    assert len(linhas) == 1
    assert "Busca: Modelo-Manual" in _ws(resp)["A3"].value


def test_termos_pendentes_pdf(client_gestor, cenario_termos):
    resp = client_gestor.get("/api/exportacoes/termos-pendentes", params={"formato": "pdf"})
    assert resp.status_code == 200
    assert resp.content[:4] == b"%PDF"


def test_termos_pendentes_aprendiz_403(client_aprendiz):
    resp = client_aprendiz.get("/api/exportacoes/termos-pendentes", params={"formato": "xlsx"})
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."


# ── Relatório mensal ────────────────────────────────────────────────────────

def _emprestar(inv_manager, item):
    ok, msg = inv_manager.issue(
        item["id"], "Ciclano", "52998224725", "101 - Puxada", "Analista", "TI", "Revalle Juazeiro",
        datetime.now().strftime("%d/%m/%Y"), "gestor.teste",
    )
    assert ok, msg


def test_relatorio_mensal_excel_mascara_cpf(client_gestor, inv_manager, criar_item):
    _emprestar(inv_manager, criar_item())
    agora = datetime.now()
    params = {"formato": "xlsx", "year": agora.year, "month": agora.month}

    resp = client_gestor.get("/api/exportacoes/relatorio-mensal", params=params)

    linhas = _planilha(resp)
    cab = _cabecalho(resp)
    assert cab[:3] == ["ID Histórico", "ID Item", "ID Periférico"]
    assert "CPF" in cab and "Data Empréstimo" in cab
    emprestimos = [l for l in linhas if l[cab.index("Operação")] == "Empréstimo"]
    assert len(emprestimos) == 1
    assert emprestimos[0][cab.index("CPF")] == "***.982.247-**"
    assert emprestimos[0][cab.index("Revenda")] == "REVALLE JUAZEIRO"
    assert emprestimos[0][cab.index("Usuário")] == "Ciclano"
    assert not any("52998224725" in c or "529.982.247-25" in c for c in _celulas(resp))
    assert f"Período: {agora.month:02d}/{agora.year}" in _ws(resp)["A3"].value
    assert _ws(resp)["A2"].value == f"Relatório mensal {agora.month:02d}/{agora.year}"
    assert f"relatorio_mensal_{agora.year}_{agora.month:02d}_" in resp.headers["content-disposition"]

    # O CSV continua trazendo o CPF como antes.
    csv = client_gestor.get("/api/reports/monthly/export", params={"year": agora.year, "month": agora.month})
    assert csv.status_code == 200
    assert "52998224725" in csv.content.decode("utf-8")


def test_relatorio_mensal_pdf(client_gestor, inv_manager, criar_item):
    _emprestar(inv_manager, criar_item())
    resp = client_gestor.get("/api/exportacoes/relatorio-mensal", params={"formato": "pdf"})
    assert resp.status_code == 200
    assert resp.content[:4] == b"%PDF"


def test_relatorio_mensal_mes_invalido_400(client_gestor):
    resp = client_gestor.get("/api/exportacoes/relatorio-mensal", params={"formato": "xlsx", "month": 13})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Mês inválido."


def test_relatorio_mensal_so_gestor(client_tecnico, client_aprendiz):
    resp = client_tecnico.get("/api/exportacoes/relatorio-mensal", params={"formato": "xlsx"})
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito ao Gestor."
    assert client_aprendiz.get(
        "/api/exportacoes/relatorio-mensal", params={"formato": "xlsx"}
    ).status_code == 403


def test_mascarar_cpf():
    from app.services.exportacoes_telas import mascarar_cpf

    assert mascarar_cpf("52998224725") == "***.982.247-**"
    assert mascarar_cpf("529.982.247-25") == "***.982.247-**"
    assert mascarar_cpf("123") == "***"
    assert mascarar_cpf(None) is None
    assert mascarar_cpf("") is None
