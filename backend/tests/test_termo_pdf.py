"""Conversão do termo .docx para PDF (LibreOffice headless)."""
import os
import shutil
import subprocess
from pathlib import Path

import pytest

from app.services import termo_pdf
from app.services.termo_pdf import ConversaoPdfFalhou, docx_para_pdf

MODELO = Path(__file__).resolve().parent.parent / "modelos" / "termo.docx"


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice não instalado")
def test_converte_modelo_real():
    pdf = docx_para_pdf(MODELO.read_bytes())
    assert pdf.startswith(b"%PDF")


def test_sem_libreoffice(monkeypatch):
    monkeypatch.setattr(termo_pdf.shutil, "which", lambda *_a, **_k: None)
    with pytest.raises(ConversaoPdfFalhou, match="LibreOffice não está instalado no servidor."):
        docx_para_pdf(b"qualquer")


def test_timeout(monkeypatch):
    monkeypatch.setattr(termo_pdf.shutil, "which", lambda *_a, **_k: "/usr/bin/soffice")

    def _trava(cmd, **kwargs):
        raise subprocess.TimeoutExpired(cmd, kwargs.get("timeout"))

    monkeypatch.setattr(termo_pdf.subprocess, "run", _trava)
    with pytest.raises(ConversaoPdfFalhou, match="A conversão do termo para PDF demorou demais."):
        docx_para_pdf(b"qualquer", timeout=1)


def test_saida_que_nao_e_pdf(monkeypatch):
    monkeypatch.setattr(termo_pdf.shutil, "which", lambda *_a, **_k: "/usr/bin/soffice")

    def _lixo(cmd, **kwargs):
        outdir = cmd[cmd.index("--outdir") + 1]
        Path(outdir, "termo.pdf").write_bytes(b"isto nao e um pdf")
        return subprocess.CompletedProcess(cmd, 0, b"", b"")

    monkeypatch.setattr(termo_pdf.subprocess, "run", _lixo)
    with pytest.raises(ConversaoPdfFalhou, match="Não foi possível converter o termo para PDF."):
        docx_para_pdf(b"qualquer")


def test_falha_ao_executar_o_binario(monkeypatch):
    monkeypatch.setattr(termo_pdf.shutil, "which", lambda *_a, **_k: "/usr/bin/soffice")

    def _quebra(cmd, **kwargs):
        raise OSError("sem permissão")

    monkeypatch.setattr(termo_pdf.subprocess, "run", _quebra)
    with pytest.raises(ConversaoPdfFalhou, match="Não foi possível converter o termo para PDF."):
        docx_para_pdf(b"qualquer")


def test_sem_arquivo_de_saida(monkeypatch):
    monkeypatch.setattr(termo_pdf.shutil, "which", lambda *_a, **_k: "/usr/bin/soffice")
    monkeypatch.setattr(
        termo_pdf.subprocess, "run",
        lambda cmd, **kw: subprocess.CompletedProcess(cmd, 1, b"", b"erro"),
    )
    with pytest.raises(ConversaoPdfFalhou, match="Não foi possível converter o termo para PDF."):
        docx_para_pdf(b"qualquer")


def test_diretorio_temporario_e_removido(monkeypatch):
    monkeypatch.setattr(termo_pdf.shutil, "which", lambda *_a, **_k: "/usr/bin/soffice")
    vistos = []

    def _grava(cmd, **kwargs):
        outdir = cmd[cmd.index("--outdir") + 1]
        vistos.append(outdir)
        Path(outdir, "termo.pdf").write_bytes(b"%PDF-1.7 ok")
        return subprocess.CompletedProcess(cmd, 0, b"", b"")

    monkeypatch.setattr(termo_pdf.subprocess, "run", _grava)
    assert docx_para_pdf(b"x") == b"%PDF-1.7 ok"
    assert vistos and not os.path.exists(vistos[0])
