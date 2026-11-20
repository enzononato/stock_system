"""Marcador `{{assinatura}}` nos modelos de termo: é onde o RevSign cria o campo de assinatura."""
import io
import re
import shutil
import subprocess
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Pt, RGBColor

from tests.fixtures_termos import unidade_juazeiro  # noqa: F401

MODELOS = Path(__file__).resolve().parent.parent / "modelos"
ARQUIVOS = sorted(MODELOS.glob("*.docx")) + sorted((MODELOS / "devolucao").glob("*.docx"))
MARCADOR = "{{assinatura}}"
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)


def _containers(doc):
    yield doc
    for secao in doc.sections:
        for parte in (
            secao.header, secao.footer, secao.first_page_header,
            secao.first_page_footer, secao.even_page_header, secao.even_page_footer,
        ):
            yield parte


def _paragrafos(container):
    yield from container.paragraphs
    for tabela in container.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                yield from _paragrafos(celula)


def _runs_com_marcador(doc):
    runs = []
    for container in _containers(doc):
        for paragrafo in _paragrafos(container):
            runs.extend(r for r in paragrafo.runs if MARCADOR in r.text.lower())
    return runs


def _texto_total(doc):
    return "\n".join(p.text for c in _containers(doc) for p in _paragrafos(c))


def test_ha_quinze_modelos():
    assert len(ARQUIVOS) == 15


@pytest.mark.parametrize("arquivo", ARQUIVOS, ids=lambda p: p.stem)
def test_cada_modelo_tem_um_marcador_branco(arquivo):
    doc = Document(str(arquivo))
    runs = _runs_com_marcador(doc)
    assert len(runs) == 1
    assert runs[0].text.strip().lower() == MARCADOR
    assert runs[0].font.color.rgb == BRANCO
    assert runs[0].font.size == Pt(8)
    # inteiro num só run (o RevSign precisa dele numa linha) e sem repetição no resto do texto
    assert _texto_total(doc).lower().count(MARCADOR) == 1


def _contem_marcador(docx_bytes: bytes) -> bool:
    return len(_runs_com_marcador(Document(io.BytesIO(docx_bytes)))) == 1


@pytest.mark.integration
def test_termo_preenchido_mantem_o_marcador(
    inv_manager, item_indisponivel, unidade_juazeiro, termo_emprestimo_disponivel, termo_devolucao_disponivel
):
    if not (termo_emprestimo_disponivel and termo_devolucao_disponivel):
        pytest.skip("Modelos de termo ausentes neste ambiente.")
    item_id = item_indisponivel["id"]

    ok, entrega, _ = inv_manager.render_loan_term(item_id, status_permitidos=("Indisponível",))
    assert ok, entrega
    assert _contem_marcador(entrega)

    ok, devolucao, _ = inv_manager.render_return_term(item_id)
    assert ok, devolucao
    assert _contem_marcador(devolucao)


@pytest.mark.skipif(
    shutil.which("soffice") is None or shutil.which("pdftotext") is None,
    reason="LibreOffice/poppler não instalados",
)
@pytest.mark.parametrize("arquivo", ARQUIVOS, ids=lambda p: p.stem)
def test_pdf_tem_o_marcador(arquivo):
    from app.services.termo_pdf import docx_para_pdf

    pdf = docx_para_pdf(arquivo.read_bytes())
    texto = subprocess.run(
        ["pdftotext", "-", "-"], input=pdf, capture_output=True, check=True, timeout=60,
    ).stdout.decode("utf-8", "replace")
    assert len(re.findall(r"\{\{assinatura\}\}", texto, flags=re.I)) == 1
