"""Conversão do termo .docx para PDF (LibreOffice headless)."""
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pytest

from app.services import termo_pdf
from app.services.termo_pdf import ConversaoPdfFalhou, docx_para_pdf

MODELO = Path(__file__).resolve().parent.parent / "modelos" / "termo.docx"
posix = pytest.mark.skipif(not sys.platform.startswith("linux"), reason="usa grupos de processos POSIX")


class PopenFalso:
    """Substitui subprocess.Popen: registra a chamada e simula o comportamento do LibreOffice."""

    chamadas: list = []

    def __init__(self, cmd, **kwargs):
        self.cmd, self.kwargs = cmd, kwargs
        self.pid = 2 ** 22 + 12345  # não existe: os.killpg levanta ProcessLookupError
        self.returncode = 0
        type(self).chamadas.append(self)
        self.ao_iniciar()

    def ao_iniciar(self):
        pass

    @property
    def outdir(self):
        return self.cmd[self.cmd.index("--outdir") + 1]

    def communicate(self, timeout=None):
        return b"", b""

    def kill(self):
        pass


@pytest.fixture
def com_soffice(monkeypatch):
    monkeypatch.setattr(termo_pdf.shutil, "which", lambda nome, *_a, **_k: f"/usr/bin/{nome}")
    PopenFalso.chamadas = []


def _usa_popen(monkeypatch, classe):
    monkeypatch.setattr(termo_pdf.subprocess, "Popen", classe)


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice não instalado")
def test_converte_modelo_real():
    pdf = docx_para_pdf(MODELO.read_bytes())
    assert pdf.startswith(b"%PDF")


def test_sem_libreoffice(monkeypatch):
    monkeypatch.setattr(termo_pdf.shutil, "which", lambda *_a, **_k: None)
    with pytest.raises(ConversaoPdfFalhou, match="LibreOffice não está instalado no servidor."):
        docx_para_pdf(b"qualquer")


def test_usa_libreoffice_quando_nao_ha_soffice(monkeypatch):
    monkeypatch.setattr(
        termo_pdf.shutil, "which", lambda nome, *_a, **_k: "/usr/bin/libreoffice" if nome == "libreoffice" else None
    )

    class Ok(PopenFalso):
        def ao_iniciar(self):
            Path(self.outdir, "termo.pdf").write_bytes(b"%PDF-1.7 ok")

    PopenFalso.chamadas = []
    _usa_popen(monkeypatch, Ok)
    assert docx_para_pdf(b"x") == b"%PDF-1.7 ok"
    assert PopenFalso.chamadas[0].cmd[0] == "/usr/bin/libreoffice"


def test_timeout(monkeypatch, com_soffice):
    class Trava(PopenFalso):
        def communicate(self, timeout=None):
            if timeout == 1:
                raise subprocess.TimeoutExpired(self.cmd, timeout)
            return b"", b""

    _usa_popen(monkeypatch, Trava)
    with pytest.raises(ConversaoPdfFalhou, match="A conversão do termo para PDF demorou demais."):
        docx_para_pdf(b"qualquer", timeout=1)


def test_saida_que_nao_e_pdf(monkeypatch, com_soffice):
    class Lixo(PopenFalso):
        def ao_iniciar(self):
            Path(self.outdir, "termo.pdf").write_bytes(b"isto nao e um pdf")

    _usa_popen(monkeypatch, Lixo)
    with pytest.raises(ConversaoPdfFalhou, match="Não foi possível converter o termo para PDF."):
        docx_para_pdf(b"qualquer")


def test_falha_ao_executar_o_binario(monkeypatch, com_soffice):
    def quebra(cmd, **kwargs):
        raise OSError("sem permissão")

    _usa_popen(monkeypatch, quebra)
    with pytest.raises(ConversaoPdfFalhou, match="Não foi possível converter o termo para PDF."):
        docx_para_pdf(b"qualquer")


def test_sem_arquivo_de_saida(monkeypatch, com_soffice):
    class Falha(PopenFalso):
        def communicate(self, timeout=None):
            self.returncode = 1
            return b"", b"erro"

    _usa_popen(monkeypatch, Falha)
    with pytest.raises(ConversaoPdfFalhou, match="Não foi possível converter o termo para PDF."):
        docx_para_pdf(b"qualquer")


def test_diretorio_temporario_e_removido(monkeypatch, com_soffice):
    class Ok(PopenFalso):
        def ao_iniciar(self):
            Path(self.outdir, "termo.pdf").write_bytes(b"%PDF-1.7 ok")

    _usa_popen(monkeypatch, Ok)
    assert docx_para_pdf(b"x") == b"%PDF-1.7 ok"
    assert not os.path.exists(PopenFalso.chamadas[0].outdir)


def test_comando_e_ambiente(monkeypatch, com_soffice):
    monkeypatch.setenv("DB_PASSWORD", "segredo-do-banco")
    monkeypatch.setenv("REVSIGN_TOKEN", "segredo-da-api")
    monkeypatch.setenv("LANG", "C.UTF-8")

    class Ok(PopenFalso):
        def ao_iniciar(self):
            Path(self.outdir, "termo.pdf").write_bytes(b"%PDF-1.7 ok")

    _usa_popen(monkeypatch, Ok)
    docx_para_pdf(b"x")
    chamada = PopenFalso.chamadas[0]
    tmp = chamada.outdir
    assert chamada.cmd[:3] == ["/usr/bin/soffice", "--headless", "--norestore"]
    assert f"-env:UserInstallation=file://" in chamada.cmd[3] and chamada.cmd[3].endswith("/perfil")
    assert chamada.cmd[4:6] == ["--convert-to", "pdf"]
    assert chamada.cmd[-1].endswith("termo.docx")
    env = chamada.kwargs["env"]
    assert env["HOME"] == tmp and env["TMPDIR"] == tmp and env["LANG"] == "C.UTF-8" and "PATH" in env
    assert "DB_PASSWORD" not in env and "REVSIGN_TOKEN" not in env
    assert set(env) <= {"PATH", "HOME", "TMPDIR", "LANG", "LC_ALL"}
    assert chamada.kwargs["start_new_session"] is True


@posix
def test_timeout_real_mata_toda_a_arvore(monkeypatch, tmp_path):
    """Binário falso que, como o oosplash/soffice.bin, deixa um neto dormindo: o timeout tem
    de matar o grupo inteiro, não só o filho direto."""
    pid_neto = tmp_path / "neto.pid"
    falso = tmp_path / "soffice"
    falso.write_text(f"#!/bin/sh\nsleep 300 &\necho $! > '{pid_neto}'\nsleep 300\n")
    falso.chmod(0o755)
    monkeypatch.setattr(termo_pdf.shutil, "which", lambda nome, *_a, **_k: str(falso) if nome == "soffice" else None)

    inicio = time.monotonic()
    with pytest.raises(ConversaoPdfFalhou, match="A conversão do termo para PDF demorou demais."):
        docx_para_pdf(b"x", timeout=2)
    assert time.monotonic() - inicio < 6

    neto = int(pid_neto.read_text())
    for _ in range(30):  # dá tempo ao kernel para reciclar o processo
        try:
            os.kill(neto, 0)
        except ProcessLookupError:
            break
        # um zumbi ainda "existe" até o init reapá-lo; confere o estado em /proc
        try:
            estado = Path(f"/proc/{neto}/stat").read_text().rsplit(")", 1)[1].split()[0]
        except FileNotFoundError:
            break
        if estado == "Z":
            break
        time.sleep(0.1)
    else:
        pytest.fail("o processo neto sobreviveu ao timeout")
