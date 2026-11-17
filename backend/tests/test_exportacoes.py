"""F7 / Task 5: exportações Excel e PDF de Estoque, Histórico, Manutenções e Linhas."""
from datetime import datetime
from io import BytesIO

import openpyxl
import pytest

from tests.fixtures_linhas import criar_aparelho, dados_linha, nova_linha  # noqa: F401
from tests.fixtures_termos import criar_colaborador  # noqa: F401

pytestmark = pytest.mark.integration

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
JUAZEIRO = "Revalle Juazeiro"
PETROLINA = "Revalle Petrolina"


def _planilha(resp) -> list[list]:
    assert resp.status_code == 200, resp.text
    ws = openpyxl.load_workbook(BytesIO(resp.content)).active
    return [list(linha) for linha in ws.iter_rows(min_row=5, values_only=True)]


def _filtros(resp) -> str:
    return openpyxl.load_workbook(BytesIO(resp.content)).active["A3"].value


def _cabecalho(resp) -> list[str]:
    ws = openpyxl.load_workbook(BytesIO(resp.content)).active
    return [c.value for c in ws[4]]


def _todas_as_celulas(resp) -> list[str]:
    ws = openpyxl.load_workbook(BytesIO(resp.content)).active
    return [str(c) for linha in ws.iter_rows(values_only=True) for c in linha if c is not None]


def _item(criar_item, executar_sql, status="Disponível", revenda=JUAZEIRO, **extra):
    item = criar_item(revenda=revenda, **extra)
    if status != "Disponível":
        executar_sql("UPDATE items SET status=%s WHERE id=%s", (status, item["id"]))
    return item


# ── Estoque ─────────────────────────────────────────────────────────────────

def test_estoque_excel_respeita_filtros_e_ordem(client_gestor, criar_item, executar_sql):
    for _ in range(2):
        _item(criar_item, executar_sql)
    _item(criar_item, executar_sql, status="Indisponível")
    _item(criar_item, executar_sql, revenda=PETROLINA)
    params = {"status": "Disponível", "revenda": JUAZEIRO, "sort": "id", "dir": "asc"}

    resp = client_gestor.get("/api/exportacoes/estoque", params={**params, "formato": "xlsx"})

    assert resp.status_code == 200
    assert resp.headers["content-type"] == XLSX
    assert "estoque_" in resp.headers["content-disposition"]
    linhas = _planilha(resp)
    esperado = client_gestor.get("/api/items", params=params).json()["items"]
    assert [l[0] for l in linhas] == [i["id"] for i in esperado]
    assert len(linhas) == 2
    cab = _cabecalho(resp)
    assert cab[0] == "Patrimônio" and "Unidade" in cab and "Valor de compra" in cab
    assert {l[cab.index("Unidade")] for l in linhas} == {"REVALLE JUAZEIRO"}
    filtros = _filtros(resp)
    assert "Status: Disponível" in filtros
    assert "Unidade: REVALLE JUAZEIRO" in filtros
    assert "Ordenação: Patrimônio (crescente)" in filtros


def test_estoque_ordenado_por_colaborador(client_gestor, criar_item, executar_sql):
    for nome in ("Ana", "Carla", "Beto"):
        item = criar_item()
        executar_sql("UPDATE items SET assigned_to=%s WHERE id=%s", (nome, item["id"]))
    params = {"sort": "colaborador", "dir": "desc"}
    resp = client_gestor.get("/api/exportacoes/estoque", params={**params, "formato": "xlsx"})
    esperado = client_gestor.get("/api/items", params=params).json()["items"]
    assert [l[0] for l in _planilha(resp)] == [i["id"] for i in esperado]
    cab = _cabecalho(resp)
    assert [l[cab.index("Colaborador")] for l in _planilha(resp)] == ["Carla", "Beto", "Ana"]
    assert "Ordenação: Colaborador (decrescente)" in _filtros(resp)


def test_estoque_filtra_por_colaborador_e_busca(client_gestor, criar_item, criar_colaborador, executar_sql):
    colab = criar_colaborador()
    a, b = criar_item(model="Alpha"), criar_item(model="Beta")
    executar_sql("UPDATE items SET colaborador_id=%s WHERE id=%s", (colab["id"], a["id"]))
    resp = client_gestor.get(
        "/api/exportacoes/estoque", params={"formato": "xlsx", "colaborador_id": colab["id"]}
    )
    assert [l[0] for l in _planilha(resp)] == [a["id"]]
    assert "Colaborador: Maria Souza" in _filtros(resp)
    resp = client_gestor.get("/api/exportacoes/estoque", params={"formato": "xlsx", "search": "Beta"})
    assert [l[0] for l in _planilha(resp)] == [b["id"]]
    assert "Busca: Beta" in _filtros(resp)


def test_estoque_pdf(client_gestor, criar_item):
    criar_item()
    resp = client_gestor.get("/api/exportacoes/estoque", params={"formato": "pdf"})
    assert resp.status_code == 200
    assert resp.content[:4] == b"%PDF"
    assert resp.headers["content-type"] == "application/pdf"


def test_estoque_sem_cpf(client_gestor, item_pendente):
    resp = client_gestor.get("/api/exportacoes/estoque", params={"formato": "xlsx"})
    celulas = _todas_as_celulas(resp)
    assert any("Fulano de Tal" in c for c in celulas)
    assert not any("11122233344" in c or "111.222.333-44" in c for c in celulas)


def test_formato_invalido_422(client_gestor):
    resp = client_gestor.get("/api/exportacoes/estoque", params={"formato": "csv"})
    assert resp.status_code == 422
    assert resp.json()["detail"].startswith("Formato")
    assert client_gestor.get("/api/exportacoes/estoque").status_code == 422


def test_ordenacao_invalida_422(client_gestor):
    resp = client_gestor.get("/api/exportacoes/estoque", params={"formato": "xlsx", "sort": "senha"})
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Ordenação inválida."


def test_pdf_acima_do_limite_400(client_gestor, criar_item, monkeypatch):
    from app.services import exportacoes_telas

    monkeypatch.setitem(exportacoes_telas.LIMITE_LINHAS, "pdf", 2)
    for _ in range(3):
        criar_item()
    resp = client_gestor.get("/api/exportacoes/estoque", params={"formato": "pdf"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == (
        "A exportação em PDF está limitada a 2 linhas e esta consulta tem 3. "
        "Refine os filtros ou exporte em Excel."
    )
    # O Excel não sofre o limite do PDF.
    assert client_gestor.get("/api/exportacoes/estoque", params={"formato": "xlsx"}).status_code == 200


def test_checar_limite_usa_separador_de_milhar():
    from app.services.exportacoes_telas import ExportacaoGrande, checar_limite

    checar_limite(5000, "pdf")
    with pytest.raises(ExportacaoGrande) as e:
        checar_limite(7000, "pdf")
    assert e.value.detail == (
        "A exportação em PDF está limitada a 5.000 linhas e esta consulta tem 7.000. "
        "Refine os filtros ou exporte em Excel."
    )
    with pytest.raises(ExportacaoGrande) as e:
        checar_limite(60000, "xlsx")
    assert e.value.detail == (
        "A exportação em Excel está limitada a 50.000 linhas e esta consulta tem 60.000. Refine os filtros."
    )


# ── Histórico ───────────────────────────────────────────────────────────────

def test_historico_excel_com_busca(client_gestor, inv_manager, criar_item):
    a, b = criar_item(), criar_item()
    hoje = datetime.now().strftime("%d/%m/%Y")
    for item, nome in ((a, "Fulano de Tal"), (b, "Beltrano")):
        ok, msg = inv_manager.issue(
            item["id"], nome, "52998224725", "101 - Puxada", "Analista", "TI", JUAZEIRO, hoje, "gestor.teste"
        )
        assert ok, msg

    resp = client_gestor.get(
        "/api/exportacoes/historico", params={"formato": "xlsx", "search": a["identificador"]}
    )
    linhas = _planilha(resp)
    cab = _cabecalho(resp)
    assert linhas and {l[cab.index("Identificador")] for l in linhas} == {a["identificador"]}
    assert f"Busca: {a['identificador']}" in _filtros(resp)
    assert "Detalhes" in cab
    assert not any("52998224725" in c for c in _todas_as_celulas(resp))
    assert {"Empréstimo", "Cadastro"} <= {l[cab.index("Operação")] for l in linhas}
    assert {l[cab.index("Unidade")] for l in linhas} <= {"REVALLE JUAZEIRO", None}

    por_item = client_gestor.get(
        "/api/exportacoes/historico", params={"formato": "xlsx", "item_id": b["id"]}
    )
    assert {l[cab.index("Identificador")] for l in _planilha(por_item)} == {b["identificador"]}
    assert client_gestor.get("/api/exportacoes/historico", params={"formato": "pdf"}).content[:4] == b"%PDF"


# ── Manutenções ─────────────────────────────────────────────────────────────

def _manutencao_com_custo(client, item_id, custo):
    resp = client.post("/api/manutencoes", json={"item_id": item_id, "defeito": "Tela quebrada"})
    assert resp.status_code == 200, resp.text
    mid = resp.json()["id"]
    resp = client.post(f"/api/manutencoes/{mid}/retorno", data={"resultado": "retornada", "custo": custo})
    assert resp.status_code == 200, resp.text
    return mid


def test_manutencoes_excel_com_total(client_gestor, criar_item):
    a, b = criar_item(), criar_item()
    _manutencao_com_custo(client_gestor, a["id"], "150.00")
    _manutencao_com_custo(client_gestor, a["id"], "89.90")
    _manutencao_com_custo(client_gestor, b["id"], "1000.00")

    resp = client_gestor.get("/api/exportacoes/manutencoes", params={"formato": "xlsx", "item_id": a["id"]})
    linhas = _planilha(resp)
    cab = _cabecalho(resp)
    assert len(linhas) == 3
    assert linhas[-1][0] == "TOTAL"
    assert linhas[-1][cab.index("Custo")] == pytest.approx(239.9)
    assert [l[cab.index("Status")] for l in linhas[:2]] == ["Retornada", "Retornada"]
    assert "Custo acumulado do equipamento" in cab
    assert f"Equipamento: #{a['id']}" in _filtros(resp)

    todas = client_gestor.get("/api/exportacoes/manutencoes", params={"formato": "xlsx"})
    assert _planilha(todas)[-1][cab.index("Custo")] == pytest.approx(1239.9)


def test_manutencoes_filtro_em_andamento_e_pdf(client_gestor, criar_item):
    a, b = criar_item(), criar_item()
    _manutencao_com_custo(client_gestor, a["id"], "10.00")
    client_gestor.post("/api/manutencoes", json={"item_id": b["id"], "defeito": "Teclado"})
    resp = client_gestor.get(
        "/api/exportacoes/manutencoes", params={"formato": "xlsx", "status": "em_andamento"}
    )
    linhas = _planilha(resp)
    assert len(linhas) == 2  # a em andamento + TOTAL
    assert "Status: Em andamento" in _filtros(resp)
    pdf = client_gestor.get("/api/exportacoes/manutencoes", params={"formato": "pdf"})
    assert pdf.content[:4] == b"%PDF"
    invalido = client_gestor.get("/api/exportacoes/manutencoes", params={"formato": "pdf", "status": "xx"})
    assert invalido.status_code == 422


# ── Linhas ──────────────────────────────────────────────────────────────────

def test_linhas_excel_com_filtro_de_status(client_gestor, criar_aparelho, nova_linha, executar_sql):
    aparelho = criar_aparelho()
    nova_linha(custo_mensal="49.90", item_id=aparelho["id"])
    suspensa = nova_linha(custo_mensal="30.10")
    client_gestor.post(f"/api/linhas/{suspensa['id']}/status", json={"status": "suspensa"})
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (aparelho["id"],))

    resp = client_gestor.get("/api/exportacoes/linhas", params={"formato": "xlsx", "status": "ativa"})
    linhas = _planilha(resp)
    cab = _cabecalho(resp)
    assert len(linhas) == 2  # a linha ativa + total
    assert linhas[0][cab.index("Status")] == "Ativa"
    assert "(baixado)" in linhas[0][cab.index("Aparelho")]
    assert linhas[0][cab.index("Unidade")] == linhas[0][cab.index("Unidade")].upper()
    assert linhas[-1][0] == "Total mensal (linhas ativas)"
    assert linhas[-1][cab.index("Custo mensal")] == pytest.approx(49.9)
    assert "ICCID" in cab
    assert "Status: Ativa" in _filtros(resp)

    todas = client_gestor.get("/api/exportacoes/linhas", params={"formato": "xlsx"})
    assert _planilha(todas)[-1][cab.index("Custo mensal")] == pytest.approx(49.9)  # só as ativas somam
    assert len(_planilha(todas)) == 3
    assert client_gestor.get("/api/exportacoes/linhas", params={"formato": "pdf"}).content[:4] == b"%PDF"


# ── Permissões ──────────────────────────────────────────────────────────────

@pytest.mark.parametrize("rota", ["estoque", "historico", "manutencoes", "linhas"])
def test_exportacoes_exigem_gestor_ou_tecnico(client_aprendiz, client_tecnico, client, rota):
    resp = client_aprendiz.get(f"/api/exportacoes/{rota}", params={"formato": "xlsx"})
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
    assert client_tecnico.get(f"/api/exportacoes/{rota}", params={"formato": "xlsx"}).status_code == 200
    assert client.get(f"/api/exportacoes/{rota}", params={"formato": "xlsx"}).status_code == 401
