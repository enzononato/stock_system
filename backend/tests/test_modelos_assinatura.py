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


def _rotulo_idx(paragrafos, inicio):
    return next(
        i for i, p in enumerate(paragrafos) if i > inicio and p.text.strip().startswith("Assinatura")
    )


@pytest.mark.parametrize("arquivo", ARQUIVOS, ids=lambda p: p.stem)
def test_marcador_fica_ligado_ao_rotulo(arquivo):
    """Do marcador até o parágrafo antes de "Assinatura…" tudo é keep-with-next: se o termo preenchido
    passar de página, o campo de assinatura não pode ficar numa página e o rótulo em outra."""
    paragrafos = Document(str(arquivo)).paragraphs
    inicio = next(i for i, p in enumerate(paragrafos) if MARCADOR in p.text.lower())
    fim = _rotulo_idx(paragrafos, inicio)
    assert fim > inicio
    for p in paragrafos[inicio:fim]:
        assert p.paragraph_format.keep_with_next is True, repr(p.text)


@pytest.mark.parametrize("arquivo", sorted((MODELOS / "devolucao").glob("*.docx")), ids=lambda p: p.stem)
def test_paragrafo_novo_da_devolucao_nao_repete_id_do_word(arquivo):
    # (o termo.docx original já traz um paraId repetido, anterior a este marcador; aqui só vale o parágrafo novo)
    ids = Document(str(arquivo)).element.body.xpath("//w:p/@w14:paraId")
    assert len(ids) == len(set(ids))


def test_script_e_idempotente_e_completa_o_que_falta(tmp_path):
    from scripts.marcar_assinatura import marcar

    for origem in (
        MODELOS / "termo.docx", MODELOS / "termo_juazeiro.docx",
        MODELOS / "devolucao" / "termo_devolucao_juazeiro.docx",
    ):
        copia = tmp_path / origem.name
        copia.write_bytes(origem.read_bytes())
        assert marcar(copia) is False  # já está pronto: não regrava
        assert copia.read_bytes() == origem.read_bytes()

        doc = Document(str(copia))
        for p in doc.paragraphs:
            p.paragraph_format.keep_with_next = None
        doc.save(str(copia))
        assert marcar(copia) is True  # refaz só o keep-with-next
        assert marcar(copia) is False
        doc = Document(str(copia))
        assert len(_runs_com_marcador(doc)) == 1


def _pagina_do_marcador_e_do_rotulo(docx_bytes: bytes, rotulo: str):
    from app.services.termo_pdf import docx_para_pdf

    pdf = docx_para_pdf(docx_bytes)
    paginas = {"marcador": None, "rotulo": None}
    for pagina in range(1, 9):
        texto = subprocess.run(
            ["pdftotext", "-f", str(pagina), "-l", str(pagina), "-", "-"],
            input=pdf, capture_output=True, timeout=60,  # página inexistente: falha e sai vazio
        ).stdout.decode("utf-8", "replace")
        if MARCADOR in texto.lower():
            paginas["marcador"] = pagina
        if rotulo in texto:
            paginas["rotulo"] = pagina
    return paginas["marcador"], paginas["rotulo"]


def _com_espaco_antes_do_marcador(arquivo: Path, pontos: int) -> bytes:
    """Simula termo preenchido mais longo: espaço extra no parágrafo logo acima do marcador, o que faz o
    marcador cair em qualquer posição perto do fim da página."""
    doc = Document(str(arquivo))
    paragrafos = doc.paragraphs
    inicio = next(i for i, p in enumerate(paragrafos) if MARCADOR in p.text.lower())
    paragrafos[inicio - 1].paragraph_format.space_after = Pt(pontos)
    saida = io.BytesIO()
    doc.save(saida)
    return saida.getvalue()


requer_pdf = pytest.mark.skipif(
    shutil.which("soffice") is None or shutil.which("pdftotext") is None,
    reason="LibreOffice/poppler não instalados",
)


@requer_pdf
@pytest.mark.parametrize("pontos", [0, 60, 92, 96, 110])
def test_devolucao_longa_mantem_marcador_e_rotulo_na_mesma_pagina(pontos):
    arquivo = MODELOS / "devolucao" / "termo_devolucao_juazeiro.docx"
    marcador, rotulo = _pagina_do_marcador_e_do_rotulo(
        _com_espaco_antes_do_marcador(arquivo, pontos), "Assinatura do Colaborador"
    )
    assert marcador is not None and marcador == rotulo
    if pontos >= 92:
        assert marcador == 2  # o caso passou da página: os dois foram juntos para a seguinte


@requer_pdf
@pytest.mark.parametrize("pontos", [416, 448])
def test_termo_longo_mantem_marcador_e_rotulo_na_mesma_pagina(pontos):
    marcador, rotulo = _pagina_do_marcador_e_do_rotulo(
        _com_espaco_antes_do_marcador(MODELOS / "termo.docx", pontos), "Assinatura do empregado"
    )
    assert marcador is not None and marcador == rotulo
    assert marcador == 5  # o termo tem 4 páginas sem o espaço extra
