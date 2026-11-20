"""Conversão do termo (.docx) para PDF com o LibreOffice em modo headless.

A API do RevSign só aceita PDF; o sistema gera o termo como .docx (python-docx).
Qualquer falha (binário ausente, travamento, saída inválida) vira `ConversaoPdfFalhou`
com mensagem em português, para o chamador cair em `manual` sem derrubar a requisição.
"""
import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)

MSG_SEM_LIBREOFFICE = "LibreOffice não está instalado no servidor."
MSG_DEMOROU = "A conversão do termo para PDF demorou demais."
MSG_FALHOU = "Não foi possível converter o termo para PDF."


class ConversaoPdfFalhou(Exception):
    """A conversão do termo para PDF não foi possível (a mensagem é própria para o usuário)."""


def docx_para_pdf(conteudo: bytes, *, timeout: int = 60) -> bytes:
    binario = shutil.which("soffice") or shutil.which("libreoffice")
    if not binario:
        raise ConversaoPdfFalhou(MSG_SEM_LIBREOFFICE)

    with tempfile.TemporaryDirectory(prefix="termo_pdf_") as tmp:
        origem = Path(tmp, "termo.docx")
        destino = Path(tmp, "termo.pdf")
        origem.write_bytes(conteudo)
        # Perfil próprio por chamada: evita lock de perfil entre conversões simultâneas.
        perfil = Path(tmp, "perfil").as_uri()
        cmd = [
            binario, "--headless", "--norestore",
            f"-env:UserInstallation={perfil}",
            "--convert-to", "pdf", "--outdir", tmp, str(origem),
        ]
        env = {**os.environ, "HOME": tmp}
        try:
            proc = subprocess.run(cmd, timeout=timeout, capture_output=True, env=env)
        except subprocess.TimeoutExpired as exc:
            raise ConversaoPdfFalhou(MSG_DEMOROU) from exc
        except OSError as exc:
            logger.warning("Falha ao executar o LibreOffice: %s", exc)
            raise ConversaoPdfFalhou(MSG_FALHOU) from exc

        pdf = destino.read_bytes() if destino.is_file() else b""
        if not pdf.startswith(b"%PDF"):
            logger.warning(
                "LibreOffice não gerou um PDF válido (código %s): %s",
                proc.returncode, (proc.stderr or b"")[-500:].decode("utf-8", "replace"),
            )
            raise ConversaoPdfFalhou(MSG_FALHOU)
        return pdf
