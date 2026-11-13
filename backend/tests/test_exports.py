"""Exportação em Excel e PDF (app/services/exports.py): puro, sem banco nem HTTP."""
import re
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO

import openpyxl
import pytest

from app.services.exports import Coluna, gerar_excel, gerar_pdf, resposta_arquivo

pytestmark = pytest.mark.unit

COLUNAS = [
    Coluna("nome", "Nome", largura=30),
    Coluna("unidade", "Unidade", maiusculas=True),
    Coluna("qtd", "Quantidade", tipo="numero"),
    Coluna("valor", "Valor", tipo="moeda"),
    Coluna("data", "Data", tipo="data"),
]
LINHAS = [
    {"nome": "Notebook Dell", "unidade": "Revalle Juazeiro", "qtd": 3, "valor": Decimal("1234.5"),
     "data": date(2026, 3, 9)},
    {"nome": "Mouse", "unidade": None, "qtd": None, "valor": None, "data": None},
]


def _planilha(**kwargs):
    conteudo = gerar_excel("Inventário", COLUNAS, LINHAS, **kwargs)
    return openpyxl.load_workbook(BytesIO(conteudo)).active


def test_gerar_excel_tem_cabecalho_e_linhas():
    ws = _planilha(filtros={"Unidade": "Juazeiro", "Status": "Disponível"})

    assert ws["A1"].value == "Inventário"
    assert ws["A1"].font.bold
    sub = ws["A2"].value
    assert re.match(r"^Gerado em \d{2}/\d{2}/\d{4} \d{2}:\d{2}", sub)
    assert "Unidade: Juazeiro" in sub and "Status: Disponível" in sub

    cabecalho = [c.value for c in ws[3]]
    assert cabecalho == ["Nome", "Unidade", "Quantidade", "Valor", "Data"]
    for celula in ws[3]:
        assert celula.font.bold
        assert celula.font.color.rgb.endswith("FFFFFF")
        assert celula.fill.fgColor.rgb.endswith("0131CD")
    assert ws.freeze_panes == "A4"
    assert ws.column_dimensions["A"].width == 30

    assert [c.value for c in ws[4]][:3] == ["Notebook Dell", "REVALLE JUAZEIRO", 3]
    assert ws["D4"].value == pytest.approx(1234.5)
    assert ws["D4"].number_format == "R$ #,##0.00"
    assert ws["E4"].value in (datetime(2026, 3, 9), date(2026, 3, 9))
    assert ws["E4"].number_format == "DD/MM/YYYY"
    # valores ausentes viram células vazias
    assert [c.value for c in ws[5]] == ["Mouse", None, None, None, None]
    assert ws.max_row == 5


def test_gerar_excel_sem_filtros_e_sem_linhas():
    conteudo = gerar_excel("Vazio", COLUNAS, [])
    ws = openpyxl.load_workbook(BytesIO(conteudo)).active
    assert ws["A1"].value == "Vazio"
    assert ws["A2"].value.startswith("Gerado em ")
    assert [c.value for c in ws[3]][0] == "Nome"
    assert ws.max_row == 3


def test_gerar_excel_neutraliza_formula_em_texto():
    colunas = [Coluna("nome", "Nome")]
    conteudo = gerar_excel("T", colunas, [{"nome": "=HYPERLINK(\"http://x\")"}, {"nome": "+1"}])
    ws = openpyxl.load_workbook(BytesIO(conteudo)).active
    assert ws["A4"].data_type == "s"  # texto, nunca fórmula
    assert ws["A4"].value == '=HYPERLINK("http://x")'
    assert ws["A5"].value == "+1"


def test_maiusculas_aplicadas():
    ws = _planilha()
    assert ws["B4"].value == "REVALLE JUAZEIRO"
    assert ws["A4"].value == "Notebook Dell"  # coluna sem `maiusculas` fica como veio


def test_formatacao_de_texto_no_pdf():
    from app.services.exports import _texto_pdf

    assert _texto_pdf("Revalle São Paulo", Coluna("u", "Unidade", maiusculas=True)) == "REVALLE SÃO PAULO"
    assert _texto_pdf("Revalle", Coluna("u", "Unidade")) == "Revalle"
    assert _texto_pdf(1234.5, Coluna("v", "Valor", tipo="moeda")) == "R$ 1.234,50"
    assert _texto_pdf(date(2026, 3, 9), Coluna("d", "Data", tipo="data")) == "09/03/2026"


def test_gerar_pdf_comeca_com_assinatura_pdf():
    linhas = [
        {"nome": "Impressora – Térmica › Zebra…", "unidade": "São Paulo", "qtd": 1,
         "valor": 10, "data": "09/03/2026"},
        {"nome": "Ação com ç, ã, é, õ, ü e emoji 😀", "unidade": "Petrolina", "qtd": 2,
         "valor": 0, "data": None},
    ]
    conteudo = gerar_pdf("Relatório de Inventário – Março", COLUNAS, linhas,
                         filtros={"Local": "Sala › 1"})
    assert conteudo[:4] == b"%PDF"
    assert len(conteudo) > 500

    retrato = gerar_pdf("Retrato", COLUNAS, linhas, paisagem=False)
    assert retrato[:4] == b"%PDF"


def test_gerar_pdf_com_muitas_linhas_quebra_pagina():
    pequeno = gerar_pdf("T", COLUNAS, LINHAS)
    grande = gerar_pdf("T", COLUNAS, LINHAS * 80)
    assert len(re.findall(rb"/Type /Page\b", pequeno)) == 1
    assert len(re.findall(rb"/Type /Page\b", grande)) > 1


def test_gerar_pdf_sem_linhas():
    assert gerar_pdf("Vazio", COLUNAS, [])[:4] == b"%PDF"


def test_sanitizar_texto_latin1():
    from app.services.exports import sanitizar_latin1

    assert sanitizar_latin1("Ação – São Paulo › 😀…") == "Ação ? São Paulo ? ??"
    assert sanitizar_latin1(None) == ""
    assert sanitizar_latin1(12) == "12"


def test_resposta_arquivo_cabecalhos():
    xlsx = resposta_arquivo(b"abc", "inventario", "xlsx")
    assert xlsx.body == b"abc"
    assert xlsx.media_type == (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    hoje = datetime.now().strftime("%Y%m%d")
    assert xlsx.headers["content-disposition"] == f'attachment; filename="inventario_{hoje}.xlsx"'

    pdf = resposta_arquivo(b"%PDF", "inventario", "pdf")
    assert pdf.media_type == "application/pdf"
    assert pdf.headers["content-disposition"] == f'attachment; filename="inventario_{hoje}.pdf"'
