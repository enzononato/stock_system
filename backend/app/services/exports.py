"""Exportação de listagens em Excel (.xlsx) e PDF.

Cada tela exportável descreve suas colunas com `Coluna` e entrega as linhas como
dicts; este módulo cuida da apresentação (cabeçalho, formatos, quebra de página).

- Unidade/revenda saem em MAIÚSCULAS (`Coluna.maiusculas`).
- CPF nunca é exportado por inteiro, a menos que uma coluna explicitamente o peça:
  quem monta as linhas decide, este módulo só escreve o que recebe.
- As fontes padrão do PDF (Helvetica) só cobrem latin-1: o texto passa por
  `sanitizar_latin1`, que troca o que não cabe por "?" em vez de quebrar.
"""
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from typing import Any, Literal, Optional
from urllib.parse import quote

from fastapi import Response
from fpdf import FPDF
from fpdf.fonts import FontFace
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

Tipo = Literal["texto", "numero", "moeda", "data"]

COR_CABECALHO = "0131CD"
FORMATO_MOEDA_EXCEL = "R$ #,##0.00"
FORMATO_DATA_EXCEL = "DD/MM/YYYY"

_MEDIA_TYPES = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pdf": "application/pdf",
}


@dataclass
class Coluna:
    chave: str
    titulo: str
    largura: int = 18
    tipo: Tipo = "texto"
    maiusculas: bool = False


MARCA = "Revalle · Patrimônio"

# Pontuação comum fora do latin-1, trocada por equivalente legível antes do "?" genérico.
_EQUIVALENTES_LATIN1 = str.maketrans({
    "–": "-", "—": "-", "…": "...", "›": ">", "‹": "<",
    "“": '"', "”": '"', "‘": "'", "’": "'", " ": " ",
})

# Caracteres de controle que o XML do .xlsx não aceita (o openpyxl levanta erro).
_CONTROLE_ILEGAL_EXCEL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def sanitizar_latin1(valor: Any) -> str:
    """Texto seguro para as fontes core do PDF: pontuação tipográfica ganha um
    equivalente ASCII ("–" -> "-", "…" -> "...", aspas curvas -> retas) e o que
    ainda não existe em latin-1 (emojis, por exemplo) vira "?"."""
    if valor is None:
        return ""
    texto = str(valor).translate(_EQUIVALENTES_LATIN1)
    return texto.encode("latin-1", "replace").decode("latin-1")


def _limpar_excel(texto: str) -> str:
    return _CONTROLE_ILEGAL_EXCEL.sub("", texto)


def _agora_texto() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M")


def _subtitulo(filtros: Optional[dict[str, str]]) -> str:
    partes = [f"Gerado em {_agora_texto()}"]
    for nome, valor in (filtros or {}).items():
        partes.append(f"{nome}: {valor}")
    return " | ".join(partes)


def _texto(valor: Any, coluna: Coluna) -> Optional[str]:
    if valor is None or valor == "":
        return None
    texto = str(valor)
    return texto.upper() if coluna.maiusculas else texto


def _numero(valor: Any) -> Any:
    """int/float/Decimal viram número; qualquer outra coisa (texto) fica como veio."""
    if isinstance(valor, bool):
        return int(valor)
    if isinstance(valor, Decimal):
        return float(valor)
    return valor


# ── Excel ────────────────────────────────────────────────────────────────────

def _valor_excel(valor: Any, coluna: Coluna) -> Any:
    if valor is None or valor == "":
        return None
    if coluna.tipo in ("numero", "moeda"):
        return _numero(valor)
    if coluna.tipo == "data":
        return valor if isinstance(valor, (date, datetime)) else str(valor)
    return _texto(valor, coluna)


def gerar_excel(
    titulo: str,
    colunas: list[Coluna],
    linhas: list[dict],
    filtros: Optional[dict[str, str]] = None,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = re.sub(r"[\[\]\*\?/\\:]", " ", _limpar_excel(titulo))[:31] or "Dados"

    # Linhas: 1 marca, 2 título, 3 "Gerado em ..." + filtros, 4 cabeçalho, 5+ dados.
    ws["A1"] = MARCA
    ws["A1"].font = Font(bold=True, size=10, color=COR_CABECALHO)
    ws["A2"] = _limpar_excel(titulo)
    ws["A2"].font = Font(bold=True, size=14)
    ws["A3"] = _limpar_excel(_subtitulo(filtros))
    ws["A3"].font = Font(italic=True, color="595959")

    cabecalho_fundo = PatternFill("solid", fgColor=COR_CABECALHO)
    for i, coluna in enumerate(colunas, start=1):
        celula = ws.cell(row=4, column=i, value=_limpar_excel(coluna.titulo))
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = cabecalho_fundo
        celula.alignment = Alignment(vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = coluna.largura

    for r, linha in enumerate(linhas, start=5):
        for i, coluna in enumerate(colunas, start=1):
            valor = _valor_excel(linha.get(coluna.chave), coluna)
            if isinstance(valor, str):
                valor = _limpar_excel(valor)
            celula = ws.cell(row=r, column=i, value=valor)
            if isinstance(valor, str):
                # openpyxl trata texto começado por "=" como fórmula; aqui é sempre texto.
                celula.data_type = "s"
            elif coluna.tipo == "moeda" and valor is not None:
                celula.number_format = FORMATO_MOEDA_EXCEL
            elif coluna.tipo == "data" and isinstance(valor, (date, datetime)):
                celula.number_format = FORMATO_DATA_EXCEL

    ws.freeze_panes = "A5"
    if colunas:
        ws.auto_filter.ref = f"A4:{get_column_letter(len(colunas))}{4 + len(linhas)}"

    saida = BytesIO()
    wb.save(saida)
    return saida.getvalue()


# ── PDF ──────────────────────────────────────────────────────────────────────

def _moeda_br(valor: Any) -> str:
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return str(valor)
    texto = f"{numero:,.2f}"  # 1,234.50
    return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _texto_pdf(valor: Any, coluna: Coluna) -> str:
    if valor is None or valor == "":
        return ""
    if coluna.tipo == "moeda":
        texto = _moeda_br(valor)
    elif coluna.tipo == "data" and isinstance(valor, datetime):
        texto = valor.strftime("%d/%m/%Y")
    elif coluna.tipo == "data" and isinstance(valor, date):
        texto = valor.strftime("%d/%m/%Y")
    elif coluna.tipo == "numero" and isinstance(valor, Decimal):
        texto = format(valor.normalize(), "f")
    else:
        texto = str(valor)
    if coluna.maiusculas:
        texto = texto.upper()
    return sanitizar_latin1(texto)


class _Relatorio(FPDF):
    def footer(self) -> None:
        self.set_y(-12)
        self.set_font("Helvetica", size=8)
        self.set_text_color(110, 110, 110)
        self.cell(0, 6, f"Página {self.page_no()}", align="C")


def gerar_pdf(
    titulo: str,
    colunas: list[Coluna],
    linhas: list[dict],
    filtros: Optional[dict[str, str]] = None,
    paisagem: bool = True,
) -> bytes:
    pdf = _Relatorio(orientation="L" if paisagem else "P", unit="mm", format="A4")
    pdf.set_margins(10, 10, 10)
    pdf.set_auto_page_break(auto=True, margin=16)
    pdf.add_page()

    pdf.set_text_color(1, 49, 205)
    pdf.set_font("Helvetica", style="B", size=9)
    pdf.cell(0, 5, sanitizar_latin1(MARCA), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", style="B", size=14)
    pdf.cell(0, 8, sanitizar_latin1(titulo), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=8)
    pdf.set_text_color(90, 90, 90)
    pdf.multi_cell(0, 4, sanitizar_latin1(_subtitulo(filtros)), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    if not colunas:
        return bytes(pdf.output())

    # As larguras das colunas são proporcionais e ocupam a largura útil da página.
    util = pdf.w - pdf.l_margin - pdf.r_margin
    total = sum(max(c.largura, 1) for c in colunas)
    larguras = [util * max(c.largura, 1) / total for c in colunas]

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", size=8)
    with pdf.table(
        col_widths=larguras,
        line_height=4.5,
        text_align="LEFT",
        headings_style=_estilo_cabecalho(),
        padding=1,
        repeat_headings=1,
    ) as tabela:
        cabecalho = tabela.row()
        for coluna in colunas:
            cabecalho.cell(sanitizar_latin1(coluna.titulo))
        for linha in linhas:
            registro = tabela.row()
            for coluna in colunas:
                registro.cell(_texto_pdf(linha.get(coluna.chave), coluna))
    return bytes(pdf.output())


def _estilo_cabecalho() -> FontFace:
    return FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(1, 49, 205))


# ── Resposta HTTP ────────────────────────────────────────────────────────────

def resposta_arquivo(
    conteudo: bytes, nome_base: str, formato: Literal["xlsx", "pdf"]
) -> Response:
    """Download do arquivo: `<nome_base>_<AAAAMMDD>.<formato>`.

    `filename` leva o nome transliterado para ASCII (acentos removidos, o resto
    vira "_"); `filename*` (RFC 5987) leva o nome original em UTF-8 percent-encoded,
    que os navegadores preferem quando entendem.
    """
    sufixo = f"_{datetime.now().strftime('%Y%m%d')}.{formato}"
    ascii_base = unicodedata.normalize("NFKD", nome_base).encode("ascii", "ignore").decode("ascii")
    ascii_base = re.sub(r"[^A-Za-z0-9_-]", "_", ascii_base) or "exportacao"
    original = quote(f"{nome_base}{sufixo}", safe="")
    return Response(
        content=conteudo,
        media_type=_MEDIA_TYPES[formato],
        headers={
            "Content-Disposition": (
                f'attachment; filename="{ascii_base}{sufixo}"; filename*=UTF-8\'\'{original}'
            )
        },
    )
