"""Serviço de envio de e-mail (SMTP) com backend em memória para testes.

As configurações são lidas de `settings` no momento da chamada (não no import),
para que os testes possam sobrescrevê-las. Nunca registra em log o corpo das
mensagens nem a senha do SMTP.
"""
import html as _html
import logging
import re
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path

from app.core.config import settings

logger = logging.getLogger(__name__)

SMTP_TIMEOUT = 15
_TEMPLATE_BASE = Path(__file__).resolve().parent.parent / "templates" / "email" / "base.html"


class MailError(Exception):
    """Falha ao enviar e-mail. A mensagem é segura para exibir ao usuário."""


@dataclass
class MailMessage:
    to: list[str]
    subject: str
    html: str
    text: str


# Mensagens "enviadas" quando MAIL_BACKEND == "memory" (usado nos testes).
outbox: list[MailMessage] = []


def limpar_outbox() -> None:
    outbox.clear()


def _html_para_texto(html: str) -> str:
    texto = re.sub(r"(?is)<(script|style|head)\b.*?</\1\s*>", "", html)
    texto = re.sub(r"(?i)<br\s*/?>|</(p|div|h[1-6]|tr|li)\s*>", "\n", texto)
    texto = re.sub(r"<[^>]+>", "", texto)
    texto = _html.unescape(texto)
    linhas = [re.sub(r"[ \t\r\f\v]+", " ", l).strip() for l in texto.split("\n")]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(linhas)).strip()


def send_mail(to: str | list[str], subject: str, html: str, text: str | None = None) -> None:
    destinatarios = [to] if isinstance(to, str) else list(to)
    assunto = re.sub(r"[\r\n]+", " ", subject).strip()
    corpo_texto = text if text is not None else _html_para_texto(html)

    if settings.MAIL_BACKEND == "memory":
        outbox.append(MailMessage(destinatarios, assunto, html, corpo_texto))
        return

    if not settings.SMTP_HOST:
        raise MailError("Envio de e-mail não configurado.")

    try:
        msg = EmailMessage()
        msg["From"] = settings.smtp_from_effective
        msg["To"] = ", ".join(destinatarios)
        msg["Subject"] = assunto
        msg.set_content(corpo_texto)
        msg.add_alternative(html, subtype="html")

        if settings.SMTP_PORT == 465:
            conexao = smtplib.SMTP_SSL(settings.SMTP_HOST, settings.SMTP_PORT, timeout=SMTP_TIMEOUT)
        else:
            conexao = smtplib.SMTP(settings.SMTP_HOST, settings.SMTP_PORT, timeout=SMTP_TIMEOUT)
        with conexao as smtp:
            if settings.SMTP_PORT != 465:
                smtp.starttls()
            if settings.SMTP_USER:
                smtp.login(settings.SMTP_USER, settings.SMTP_PASS)
            smtp.send_message(msg)
    except (smtplib.SMTPException, OSError) as exc:
        # Só o tipo do erro: nada de corpo, destinatários ou credenciais no log.
        logger.error("Falha ao enviar e-mail via SMTP (%s).", type(exc).__name__)
        raise MailError("Não foi possível enviar o e-mail.") from exc


def render_email(
    titulo: str,
    corpo_html: str,
    acao_url: str | None = None,
    acao_rotulo: str | None = None,
) -> str:
    """Envolve `corpo_html` (já seguro/confiável) no layout padrão.

    `titulo`, `acao_url` e `acao_rotulo` são escapados; `corpo_html` é inserido como está.
    """
    acao = ""
    if acao_url and acao_rotulo:
        acao = (
            '<p style="margin:24px 0;"><a href="'
            + _html.escape(acao_url, quote=True)
            + '" style="background:#0131CD;color:#ffffff;text-decoration:none;'
            'padding:12px 24px;border-radius:6px;font-weight:bold;display:inline-block;">'
            + _html.escape(acao_rotulo)
            + "</a></p>"
        )
    modelo = _TEMPLATE_BASE.read_text(encoding="utf-8")
    valores = {"TITULO": _html.escape(titulo), "CORPO": corpo_html, "ACAO": acao}
    # Passada única: um valor que contenha "{{...}}" não é reinterpretado.
    return re.sub(r"\{\{(TITULO|CORPO|ACAO)\}\}", lambda m: valores[m.group(1)], modelo)
