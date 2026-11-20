"""Põe o marcador `{{assinatura}}` nos modelos de termo (backend/modelos).

O RevSign acha o texto `{{assinatura}}` no PDF do termo, cobre-o de branco e cria o campo de
assinatura do colaborador (180 x 54 pt) com o canto superior esquerdo no início do marcador. Por isso
o marcador:
- é um run próprio, em fonte branca de 8 pt (na impressão do .docx manual não aparece);
- fica ACIMA da linha de assinatura, com espaço livre embaixo para o campo: o campo não pode cobrir
  texto do termo (a linha de assinatura em branco pode ser coberta);
- fica à direita do rótulo, ou centralizado sobre a linha, conforme o modelo.

O texto visível dos modelos não muda. O script é idempotente: modelo que já tem o marcador não é
mexido. Para refazer depois de trocar um modelo: `python scripts/marcar_assinatura.py` (a partir de
backend/); depois confira o PDF convertido (seção 3 de docs/integracoes/revsign-contrato-patrimonio.md).

Uso: python scripts/marcar_assinatura.py [arquivo.docx ...]   (sem argumentos: os 15 modelos)
"""
import copy
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

MARCADOR = "{{assinatura}}"
_RE_MARCADOR = re.compile(r"\{\{\s*assinatura\s*\}\}", re.IGNORECASE)
LARGURA_CAMPO = Pt(180)
TAMANHO_FONTE = Pt(8)
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
# Espaço extra, em termo.docx, em volta do marcador: depois dele, para o campo (54 pt de altura) terminar
# antes de "Assinatura do empregado"; antes dele, para não encostar na linha da data.
ESPACO_ANTES_TERMO = Pt(6)
ESPACO_DEPOIS_TERMO = Pt(36)

MODELOS = Path(__file__).resolve().parent.parent / "modelos"


class ModeloInesperado(Exception):
    """O modelo não tem a estrutura esperada em volta da assinatura; ajuste o script, não o modelo às cegas."""


def _todos_os_paragrafos(doc):
    containers = [doc]
    for secao in doc.sections:
        containers += [
            secao.header, secao.footer, secao.first_page_header,
            secao.first_page_footer, secao.even_page_header, secao.even_page_footer,
        ]

    def varrer(container):
        yield from container.paragraphs
        for tabela in container.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    yield from varrer(celula)

    for container in containers:
        yield from varrer(container)


def tem_marcador(doc) -> bool:
    return any(_RE_MARCADOR.search(p.text) for p in _todos_os_paragrafos(doc))


def _largura_util(doc):
    secao = doc.sections[-1]
    return secao.page_width - secao.left_margin - secao.right_margin


def _recuo_centralizado(doc):
    """Recuo à esquerda que centraliza o campo de 180 pt entre as margens."""
    return int(_largura_util(doc) / 2 - LARGURA_CAMPO / 2)


def _adicionar_marcador(paragrafo):
    run = paragrafo.add_run(MARCADOR)
    run.font.size = TAMANHO_FONTE
    run.font.color.rgb = BRANCO
    return run


def _ultimo_paragrafo(doc, prefixo: str):
    achados = [i for i, p in enumerate(doc.paragraphs) if p.text.strip().startswith(prefixo)]
    if not achados:
        raise ModeloInesperado(f"parágrafo '{prefixo}…' não encontrado")
    return achados[-1]


def _marcar_devolucao(doc):
    """Parágrafo novo, à direita do rótulo "Assinatura do Colaborador:", logo acima da linha de sublinhados."""
    paragrafos = doc.paragraphs
    rotulo = paragrafos[_ultimo_paragrafo(doc, "Assinatura do Colaborador")]
    novo = copy.deepcopy(rotulo._p)
    rotulo._p.addprevious(novo)
    for filho in list(novo):
        if filho.tag != qn("w:pPr"):
            novo.remove(filho)
    # o clone herdou o estilo/formatação do rótulo; só interessa a fonte pequena do parágrafo
    from docx.text.paragraph import Paragraph

    marcador = Paragraph(novo, rotulo._parent)
    formato = marcador.paragraph_format
    formato.left_indent = Pt(190)  # depois do rótulo "Assinatura do Colaborador:" (~170 pt)
    formato.first_line_indent = 0
    formato.space_before = 0
    formato.space_after = 0
    formato.line_spacing = 1
    _adicionar_marcador(marcador)
    marcador._p.get_or_add_pPr()  # garante pPr
    mark = marcador._p.pPr.find(qn("w:rPr"))
    if mark is not None:
        marcador._p.pPr.remove(mark)
    _tirar_quebra_final(rotulo)
    return marcador


def _tirar_quebra_final(paragrafo):
    """O rótulo termina com uma quebra de linha solta (linha em branco depois da assinatura). O marcador
    acrescenta ~9 pt acima; tirar essa linha em branco compensa e evita empurrar o fim do termo (que
    já fica rente à margem quando a lista de periféricos é longa) para uma página nova."""
    runs = paragrafo._p.findall(qn("w:r"))
    if not runs:
        return
    ultimo = runs[-1]
    conteudo = [c for c in ultimo if c.tag != qn("w:rPr")]
    if conteudo and all(c.tag == qn("w:br") and c.get(qn("w:type")) in (None, "textWrapping") for c in conteudo):
        paragrafo._p.remove(ultimo)


def _marcar_termo(doc):
    """termo.docx: o parágrafo vazio entre a data e a linha vira o marcador, com espaço depois dele."""
    rotulo = _ultimo_paragrafo(doc, "Assinatura do empregado")
    linha = doc.paragraphs[rotulo - 1]
    vazio = doc.paragraphs[rotulo - 2]
    if not set(linha.text.strip()) <= {"_"} or not linha.text.strip():
        raise ModeloInesperado("a linha de sublinhados não está logo acima de 'Assinatura do empregado'")
    if vazio.text.strip():
        raise ModeloInesperado("o parágrafo acima da linha de assinatura não está vazio")
    vazio.alignment = WD_ALIGN_PARAGRAPH.LEFT
    vazio.paragraph_format.left_indent = _recuo_centralizado(doc)
    vazio.paragraph_format.space_before = ESPACO_ANTES_TERMO
    vazio.paragraph_format.space_after = ESPACO_DEPOIS_TERMO
    _adicionar_marcador(vazio)
    return vazio


def _marcar_unidade(doc):
    """Modelos por unidade: 3 parágrafos vazios acima de "Assinatura" (o último com a linha desenhada);
    o marcador vai no primeiro deles, centralizado, e o campo termina antes do rótulo."""
    rotulo = _ultimo_paragrafo(doc, "Assinatura")
    paragrafos = doc.paragraphs
    acima = paragrafos[rotulo - 3: rotulo]
    if len(acima) != 3 or any(p.text.strip() for p in acima):
        raise ModeloInesperado("esperava 3 parágrafos vazios acima de 'Assinatura'")
    if acima[2]._p.pPr is None or acima[2]._p.pPr.find(qn("w:pBdr")) is None:
        raise ModeloInesperado("esperava a linha de assinatura no último parágrafo vazio")
    alvo = acima[0]
    alvo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    alvo.paragraph_format.left_indent = _recuo_centralizado(doc)
    _adicionar_marcador(alvo)
    return alvo


def tipo_do_modelo(caminho: Path) -> str:
    if caminho.parent.name == "devolucao":
        return "devolucao"
    return "termo" if caminho.name == "termo.docx" else "unidade"


_MARCADORES = {"devolucao": _marcar_devolucao, "termo": _marcar_termo, "unidade": _marcar_unidade}


def marcar(caminho: Path) -> bool:
    """Marca o modelo. Devolve True se alterou o arquivo, False se já tinha o marcador."""
    doc = Document(str(caminho))
    if tem_marcador(doc):
        return False
    _MARCADORES[tipo_do_modelo(caminho)](doc)
    doc.save(str(caminho))
    return True


def modelos():
    return sorted(MODELOS.glob("*.docx")) + sorted((MODELOS / "devolucao").glob("*.docx"))


def main(argv) -> int:
    alvos = [Path(a) for a in argv] or modelos()
    for caminho in alvos:
        try:
            alterado = marcar(caminho)
        except ModeloInesperado as exc:
            print(f"ERRO {caminho.name}: {exc}", file=sys.stderr)
            return 1
        print(f"{'marcado' if alterado else 'já tinha'}: {caminho.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
