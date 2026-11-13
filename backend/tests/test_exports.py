"""Exportação em Excel e PDF (app/services/exports.py): puro, sem banco nem HTTP."""
import re
import zlib
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO

import openpyxl
import pytest

from app.services.exports import (
    Coluna,
    gerar_excel,
    gerar_pdf,
    resposta_arquivo,
    sanitizar_latin1,
)

pytestmark = pytest.mark.unit

MARCA = "Revalle · Patrimônio"

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


def _conteudo_pdf(pdf: bytes) -> bytes:
    """Bytes do PDF com os streams (comprimidos) descomprimidos, para procurar texto.
    Com fontes core o texto sai em latin-1, dentro de `(...) Tj`."""
    partes = [pdf]
    for m in re.finditer(rb"stream\r?\n(.*?)\r?\nendstream", pdf, re.DOTALL):
        try:
            partes.append(zlib.decompress(m.group(1)))
        except zlib.error:
            pass
    return b"\n".join(partes)


def _paginas(pdf: bytes) -> int:
    return len(re.findall(rb"/Type /Page\b", pdf))


def test_gerar_excel_tem_cabecalho_e_linhas():
    ws = _planilha(filtros={"Unidade": "Juazeiro", "Status": "Disponível"})

    assert ws["A1"].value == MARCA
    assert ws["A2"].value == "Inventário"
    assert ws["A2"].font.bold
    sub = ws["A3"].value
    assert re.match(r"^Gerado em \d{2}/\d{2}/\d{4} \d{2}:\d{2}", sub)
    assert "Unidade: Juazeiro" in sub and "Status: Disponível" in sub

    cabecalho = [c.value for c in ws[4]]
    assert cabecalho == ["Nome", "Unidade", "Quantidade", "Valor", "Data"]
    for celula in ws[4]:
        assert celula.font.bold
        assert celula.font.color.rgb.endswith("FFFFFF")
        assert celula.fill.fgColor.rgb.endswith("0131CD")
    assert ws.freeze_panes == "A5"
    assert ws.column_dimensions["A"].width == 30

    assert [c.value for c in ws[5]][:3] == ["Notebook Dell", "REVALLE JUAZEIRO", 3]
    assert ws["D5"].value == pytest.approx(1234.5)
    assert ws["D5"].number_format == "R$ #,##0.00"
    assert ws["E5"].value in (datetime(2026, 3, 9), date(2026, 3, 9))
    assert ws["E5"].number_format == "DD/MM/YYYY"
    # valores ausentes viram células vazias
    assert [c.value for c in ws[6]] == ["Mouse", None, None, None, None]
    assert ws.max_row == 6


def test_gerar_excel_sem_filtros_e_sem_linhas():
    conteudo = gerar_excel("Vazio", COLUNAS, [])
    ws = openpyxl.load_workbook(BytesIO(conteudo)).active
    assert ws["A1"].value == MARCA
    assert ws["A2"].value == "Vazio"
    assert ws["A3"].value.startswith("Gerado em ")
    assert [c.value for c in ws[4]][0] == "Nome"
    assert ws.max_row == 4


def test_gerar_excel_neutraliza_formula_em_texto():
    colunas = [Coluna("nome", "Nome")]
    conteudo = gerar_excel("T", colunas, [{"nome": "=HYPERLINK(\"http://x\")"}, {"nome": "+1"}])
    ws = openpyxl.load_workbook(BytesIO(conteudo)).active
    assert ws["A5"].data_type == "s"  # texto, nunca fórmula
    assert ws["A5"].value == '=HYPERLINK("http://x")'
    assert ws["A6"].value == "+1"


def test_gerar_excel_remove_caracteres_de_controle_ilegais():
    sujo = "A\x00B\x01C\x08D\x0bE\x0cF\x0eG\x1fH"
    colunas = [Coluna("nome", "Nome\x07"), Coluna("data", "Data", tipo="data")]
    conteudo = gerar_excel(
        "Título\x00", colunas, [{"nome": sujo + "\tcom tab\ne quebra", "data": "9/3\x02"}],
        filtros={"F\x03": "v\x04"},
    )
    ws = openpyxl.load_workbook(BytesIO(conteudo)).active
    assert ws["A2"].value == "Título"
    assert ws["A4"].value == "Nome"
    assert ws["A5"].value == "ABCDEFGH\tcom tab\ne quebra"  # tab e quebra de linha são válidos
    assert ws["B5"].value == "9/3"
    assert "F: v" in ws["A3"].value


def test_maiusculas_aplicadas():
    ws = _planilha()
    assert ws["B5"].value == "REVALLE JUAZEIRO"
    assert ws["A5"].value == "Notebook Dell"  # coluna sem `maiusculas` fica como veio


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


def test_pdf_tem_marca_e_preserva_acentos():
    pdf = gerar_pdf("Relatório de Manutenção", COLUNAS, LINHAS)
    texto = _conteudo_pdf(pdf)
    # latin-1: ó = \xf3, ç = \xe7, ã = \xe3, ô = \xf4, · = \xb7
    assert "Relatório de Manutenção".encode("latin-1") in texto
    assert MARCA.encode("latin-1") in texto


def test_gerar_pdf_com_muitas_linhas_quebra_pagina_e_repete_cabecalho():
    pequeno = gerar_pdf("T", COLUNAS, LINHAS)
    grande = gerar_pdf("T", COLUNAS, LINHAS * 80)
    assert _paginas(pequeno) == 1
    paginas = _paginas(grande)
    assert paginas > 1
    # o título de coluna aparece uma vez por página (cabeçalho repetido)
    assert _conteudo_pdf(grande).count(b"(Quantidade)") >= paginas
    assert _conteudo_pdf(pequeno).count(b"(Quantidade)") == 1


def test_gerar_pdf_sem_linhas():
    assert gerar_pdf("Vazio", COLUNAS, [])[:4] == b"%PDF"


def test_sanitizar_texto_latin1():
    assert sanitizar_latin1("Ação – São Paulo › 😀…") == "Ação - São Paulo > ?..."
    assert sanitizar_latin1(None) == ""
    assert sanitizar_latin1(12) == "12"


def test_sanitizar_equivalentes_de_pontuacao():
    assert sanitizar_latin1("a–b—c") == "a-b-c"
    assert sanitizar_latin1("fim…") == "fim..."
    assert sanitizar_latin1("‹x› “y” ‘z’") == "<x> \"y\" 'z'"
    assert sanitizar_latin1("a b") == "a b"
    assert sanitizar_latin1("Relatório de Manutenção") == "Relatório de Manutenção"
    # o que continua fora do latin-1 não derruba nada
    assert sanitizar_latin1("日本語 😀") == "??? ?"
    sanitizar_latin1("日本語").encode("latin-1")


def test_resposta_arquivo_cabecalhos():
    hoje = datetime.now().strftime("%Y%m%d")
    xlsx = resposta_arquivo(b"abc", "inventario", "xlsx")
    assert xlsx.body == b"abc"
    assert xlsx.media_type == (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert xlsx.headers["content-disposition"] == (
        f"attachment; filename=\"inventario_{hoje}.xlsx\"; filename*=UTF-8''inventario_{hoje}.xlsx"
    )

    pdf = resposta_arquivo(b"%PDF", "inventario", "pdf")
    assert pdf.media_type == "application/pdf"
    assert pdf.headers["content-disposition"].startswith(
        f'attachment; filename="inventario_{hoje}.pdf"'
    )


def test_resposta_arquivo_nome_com_acentos():
    hoje = datetime.now().strftime("%Y%m%d")
    resp = resposta_arquivo(b"x", "Relatório de Manutenção", "pdf")
    cabecalho = resp.headers["content-disposition"]
    assert f'filename="Relatorio_de_Manutencao_{hoje}.pdf"' in cabecalho
    assert (
        f"filename*=UTF-8''Relat%C3%B3rio%20de%20Manuten%C3%A7%C3%A3o_{hoje}.pdf" in cabecalho
    )
    cabecalho.encode("latin-1")  # o cabeçalho HTTP é só ASCII
    assert cabecalho.isascii()
