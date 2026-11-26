"""Conversão do termo (.docx) para PDF com o LibreOffice em modo headless.

A API do RevSign só aceita PDF; o sistema gera o termo como .docx (python-docx).
Qualquer falha (binário ausente, travamento, saída inválida) vira `ConversaoPdfFalhou`
com mensagem em português, para o chamador cair em `manual` sem derrubar a requisição.
"""
import logging
import os
import shutil
import signal
import subprocess
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)

MSG_SEM_LIBREOFFICE = "LibreOffice não está instalado no servidor."
MSG_DEMOROU = "A conversão do termo para PDF demorou demais."
MSG_FALHOU = "Não foi possível converter o termo para PDF."


class ConversaoPdfFalhou(Exception):
    """A conversão do termo para PDF não foi possível (a mensagem é própria para o usuário)."""


def _ambiente_minimo(tmp: str) -> dict:
    """Ambiente mínimo para o LibreOffice: o do app tem segredos (banco, APIs) que ele não usa."""
    env = {"PATH": os.environ.get("PATH", "/usr/local/bin:/usr/bin:/bin"), "HOME": tmp, "TMPDIR": tmp}
    for chave in ("LANG", "LC_ALL"):
        if os.environ.get(chave):
            env[chave] = os.environ[chave]
    return env


def _matar_arvore(proc: subprocess.Popen) -> None:
    """O `soffice` do Debian termina em `exec oosplash`, que sobe o `soffice.bin`: mata o grupo
    de processos inteiro, não só o filho direto."""
    try:
        os.killpg(proc.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass
    except OSError as exc:
        logger.warning("Não foi possível encerrar o grupo do LibreOffice: %s", exc)
        proc.kill()


def docx_para_pdf(conteudo: bytes, *, timeout: int = 60) -> bytes:
    binario = shutil.which("soffice") or shutil.which("libreoffice")
    if not binario:
        raise ConversaoPdfFalhou(MSG_SEM_LIBREOFFICE)

    # ignore_cleanup_errors: um resto de processo não pode trocar o erro claro por um OSError.
    with tempfile.TemporaryDirectory(prefix="termo_pdf_", ignore_cleanup_errors=True) as tmp:
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
        try:
            proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                env=_ambiente_minimo(tmp), start_new_session=True,
            )
        except OSError as exc:
            logger.warning("Falha ao executar o LibreOffice: %s", exc)
            raise ConversaoPdfFalhou(MSG_FALHOU) from exc

        try:
            _, erro = proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            _matar_arvore(proc)
            try:
                proc.communicate(timeout=5)
            except Exception:  # noqa: BLE001 - já estamos falhando; só não deixar zumbi
                pass
            raise ConversaoPdfFalhou(MSG_DEMOROU) from exc
        except BaseException:
            _matar_arvore(proc)
            raise
        # O processo principal saiu; garante que nada do grupo ficou para trás.
        _matar_arvore(proc)

        pdf = destino.read_bytes() if destino.is_file() else b""
        if not pdf.startswith(b"%PDF"):
            logger.warning(
                "LibreOffice não gerou um PDF válido (código %s): %s",
                proc.returncode, (erro or b"")[-500:].decode("utf-8", "replace"),
            )
            raise ConversaoPdfFalhou(MSG_FALHOU)
        return pdf
