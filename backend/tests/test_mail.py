"""Testes unitários do serviço de e-mail (app/services/mail.py). Sem rede nem banco."""
import smtplib

import pytest

from app.core.config import settings
from app.services import mail
from app.services.mail import MailError, render_email, send_mail

pytestmark = pytest.mark.unit


class _SMTPFalso:
    """Substituto de smtplib.SMTP/SMTP_SSL que registra as chamadas."""

    instancias: list = []
    falhar_em: str | None = None

    def __init__(self, host, port, timeout=None, **kwargs):
        self.host, self.port, self.timeout = host, port, timeout
        self.chamadas: list = []
        self.mensagens: list = []
        _SMTPFalso.instancias.append(self)
        if _SMTPFalso.falhar_em == "init":
            raise smtplib.SMTPConnectError(421, b"indisponivel")

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def starttls(self, *a, **k):
        self.chamadas.append("starttls")

    def login(self, user, senha):
        if _SMTPFalso.falhar_em == "login":
            raise smtplib.SMTPAuthenticationError(535, b"negado")
        self.chamadas.append(("login", user, senha))

    def send_message(self, msg):
        self.chamadas.append("send_message")
        self.mensagens.append(msg)


@pytest.fixture(autouse=True)
def _ambiente(monkeypatch):
    _SMTPFalso.instancias = []
    _SMTPFalso.falhar_em = None
    monkeypatch.setattr(settings, "MAIL_BACKEND", "smtp")
    monkeypatch.setattr(settings, "SMTP_HOST", "smtp.exemplo.com")
    monkeypatch.setattr(settings, "SMTP_PORT", 465)
    monkeypatch.setattr(settings, "SMTP_USER", "naoresponda@revalle.com.br")
    monkeypatch.setattr(settings, "SMTP_PASS", "segredo-smtp")
    monkeypatch.setattr(settings, "SMTP_FROM", "")
    monkeypatch.setattr(smtplib, "SMTP_SSL", _SMTPFalso)
    monkeypatch.setattr(smtplib, "SMTP", _SMTPFalso)
    mail.limpar_outbox()
    yield
    mail.limpar_outbox()


def test_memory_backend_guarda_mensagem(monkeypatch):
    monkeypatch.setattr(settings, "MAIL_BACKEND", "memory")
    send_mail("a@b.com", "Assunto", "<p>Oi</p>")
    assert mail.outbox[-1].to == ["a@b.com"]
    assert mail.outbox[-1].subject == "Assunto"
    assert mail.outbox[-1].text == "Oi"
    assert _SMTPFalso.instancias == []


def test_smtp_ssl_na_porta_465():
    send_mail(["a@b.com", "c@d.com"], "Assunto", "<p>Oi <b>você</b></p>")
    smtp = _SMTPFalso.instancias[0]
    assert (smtp.host, smtp.port, smtp.timeout) == ("smtp.exemplo.com", 465, 15)
    assert ("login", "naoresponda@revalle.com.br", "segredo-smtp") in smtp.chamadas
    assert "starttls" not in smtp.chamadas
    assert "send_message" in smtp.chamadas
    msg = smtp.mensagens[0]
    assert msg["From"] == "naoresponda@revalle.com.br"
    assert msg["To"] == "a@b.com, c@d.com"
    assert msg["Subject"] == "Assunto"
    assert msg.get_body(("plain",)).get_content().strip() == "Oi você"
    assert "<b>" in msg.get_body(("html",)).get_content()


def test_starttls_em_outra_porta(monkeypatch):
    monkeypatch.setattr(settings, "SMTP_PORT", 587)
    send_mail("a@b.com", "Assunto", "<p>Oi</p>")
    smtp = _SMTPFalso.instancias[0]
    assert smtp.port == 587
    assert smtp.chamadas.index("starttls") < smtp.chamadas.index(
        ("login", "naoresponda@revalle.com.br", "segredo-smtp")
    )
    assert "send_message" in smtp.chamadas


def test_from_usa_smtp_from_quando_definido(monkeypatch):
    monkeypatch.setattr(settings, "SMTP_FROM", "Patrimônio <patrimonio@revalle.com.br>")
    send_mail("a@b.com", "Assunto", "<p>Oi</p>")
    assert _SMTPFalso.instancias[0].mensagens[0]["From"].endswith("<patrimonio@revalle.com.br>")


@pytest.mark.parametrize("etapa", ["init", "login"])
def test_falha_smtp_vira_mail_error(etapa):
    _SMTPFalso.falhar_em = etapa
    with pytest.raises(MailError) as exc:
        send_mail("a@b.com", "Assunto", "<p>Oi</p>")
    assert str(exc.value) == "Não foi possível enviar o e-mail."


def test_oserror_vira_mail_error(monkeypatch):
    def _explode(*a, **k):
        raise OSError("rede fora")

    monkeypatch.setattr(smtplib, "SMTP_SSL", _explode)
    with pytest.raises(MailError, match="Não foi possível enviar o e-mail."):
        send_mail("a@b.com", "Assunto", "<p>Oi</p>")


def test_sem_host_configurado_recusa(monkeypatch):
    monkeypatch.setattr(settings, "SMTP_HOST", "")
    with pytest.raises(MailError) as exc:
        send_mail("a@b.com", "Assunto", "<p>Oi</p>")
    assert str(exc.value) == "Envio de e-mail não configurado."
    assert _SMTPFalso.instancias == []


def test_falha_nao_loga_corpo_nem_senha(monkeypatch, caplog):
    _SMTPFalso.falhar_em = "login"
    with caplog.at_level("DEBUG"):
        with pytest.raises(MailError):
            send_mail("a@b.com", "Assunto", "<p>corpo-sigiloso</p>")
    assert "corpo-sigiloso" not in caplog.text
    assert "segredo-smtp" not in caplog.text


def test_render_email_inclui_botao():
    html = render_email(
        "Redefinir <senha>",
        "<p>Clique abaixo.</p>",
        acao_url="https://app.exemplo.com/reset?token=abc&x=1",
        acao_rotulo="Redefinir senha",
    )
    assert "https://app.exemplo.com/reset?token=abc&amp;x=1" in html
    assert "Redefinir senha" in html
    assert "Redefinir &lt;senha&gt;" in html
    assert "Redefinir <senha>" not in html
    assert "<p>Clique abaixo.</p>" in html
    assert "#0131CD" in html
    assert "Revalle · Patrimônio" in html
    assert "Mensagem automática, não responda." in html


def test_render_email_sem_botao():
    html = render_email("Aviso", "<p>Texto</p>")
    assert "<p>Texto</p>" in html
    assert "<a href" not in html


def test_render_email_nao_reinterpreta_marcadores():
    html = render_email("{{CORPO}}", "<p>x</p>")
    assert html.count("<p>x</p>") == 1
