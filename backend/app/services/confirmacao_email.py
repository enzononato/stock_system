"""Confirmação do e-mail de login por um código de 6 dígitos.

Usado na etapa `cadastro_email` do login: a conta sem e-mail confirmado informa o
e-mail corporativo, recebe um código e só entra depois de digitá-lo. O código vale
15 minutos, pode ser pedido de novo depois de 60 segundos (o anterior deixa de
valer) e 5 erros o invalidam. Só o HMAC do código é guardado.
"""
import hashlib
import hmac
import html
import logging
import secrets
from datetime import datetime, timedelta, timezone

from app.core import config
from app.db.email_confirmacao_db import EmailConfirmacaoDB
from app.db.user_manager_db import EmailEmUso, UserDBManager
from app.schemas.validators import validar_email_corporativo
from app.services import audit
from app.services.audit import AuditContext
from app.services.mail import MailError, render_email, send_mail

logger = logging.getLogger(__name__)

VALIDADE_MINUTOS = 15
REENVIO_SEGUNDOS = 60
MAX_TENTATIVAS = 5

MSG_CODIGO_INVALIDO = "Código inválido ou expirado."
MSG_CODIGO_ESGOTADO = "Muitas tentativas. Entre de novo para receber outro código."
MSG_AGUARDE_REENVIO = "Aguarde 1 minuto para pedir outro código."
MSG_EMAIL_EM_USO = "Este e-mail já está em uso por outra conta."
MSG_ENVIO_FALHOU = "Não foi possível enviar o e-mail agora. Tente de novo em instantes."
ASSUNTO = "Código de confirmação — Patrimônio Revalle"


class ConfirmacaoErro(Exception):
    """Erro de negócio da confirmação; `status_code` é o da resposta HTTP."""

    status_code = 400
    mensagem_padrao = ""

    def __init__(self, mensagem: str | None = None):
        super().__init__(mensagem or self.mensagem_padrao)


class CodigoInvalido(ConfirmacaoErro):
    status_code = 400
    mensagem_padrao = MSG_CODIGO_INVALIDO


class CodigoEsgotado(ConfirmacaoErro):
    status_code = 400
    mensagem_padrao = MSG_CODIGO_ESGOTADO


class AguardeReenvio(ConfirmacaoErro):
    status_code = 429
    mensagem_padrao = MSG_AGUARDE_REENVIO


class EmailEmUsoErro(ConfirmacaoErro):
    status_code = 409
    mensagem_padrao = MSG_EMAIL_EM_USO


class DominioInvalido(ConfirmacaoErro):
    status_code = 422


class EnvioFalhou(ConfirmacaoErro):
    status_code = 503
    mensagem_padrao = MSG_ENVIO_FALHOU


def _agora_utc() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def _hash_codigo(usuario_id: int, email: str, codigo: str) -> str:
    """HMAC-SHA256 do código, preso à conta e ao e-mail (lido na chamada: os testes
    recarregam `app.core.config`)."""
    chave = config.settings.JWT_SECRET.encode()
    return hmac.new(chave, f"{usuario_id}:{email}:{codigo}".encode(), hashlib.sha256).hexdigest()


def mascarar_email(email: str | None) -> str | None:
    """`joao.silva@revalle.com.br` -> `j***@revalle.com.br`."""
    if not email:
        return email
    local, _, dominio = email.partition("@")
    return f"{local[:1]}***@{dominio}"


def _enviar(email: str, codigo: str) -> None:
    corpo = (
        "<p>Use o código abaixo para confirmar o seu e-mail no Patrimônio Revalle:</p>"
        f'<p style="font-size:28px;font-weight:bold;letter-spacing:6px;margin:16px 0">{html.escape(codigo)}</p>'
        "<p>Vale por 15 minutos.</p>"
        "<p>Se não foi você que pediu, ignore este e-mail.</p>"
    )
    send_mail(email, ASSUNTO, render_email("Confirme seu e-mail", corpo))


def iniciar(usuario: dict, email: str, ctx: AuditContext) -> str:
    """Valida o e-mail, confere o intervalo de reenvio e se outra conta o usa, gera o
    código, grava só o hash e envia o e-mail. Devolve o e-mail normalizado."""
    try:
        email = validar_email_corporativo(email)
    except ValueError as exc:
        raise DominioInvalido(str(exc))

    db = EmailConfirmacaoDB()
    desde_ultimo = db.segundos_desde_ultimo(usuario["id"])
    if desde_ultimo is not None and desde_ultimo < REENVIO_SEGUNDOS:
        raise AguardeReenvio()
    if UserDBManager().email_em_uso_por_outro(email, usuario["id"]):
        raise EmailEmUsoErro()

    codigo = f"{secrets.randbelow(10**6):06d}"
    confirmacao_id = db.criar(
        usuario["id"],
        email,
        _hash_codigo(usuario["id"], email, codigo),
        _agora_utc() + timedelta(minutes=VALIDADE_MINUTOS),
    )
    try:
        _enviar(email, codigo)
    except Exception as exc:
        # O código não chegou: some com ele, para não travar o reenvio por 60 s.
        db.apagar(confirmacao_id)
        if isinstance(exc, MailError):
            logger.error("Falha ao enviar o código de confirmação de e-mail.")
        else:
            logger.error("Erro inesperado ao enviar o código de confirmação (%s).", type(exc).__name__)
        raise EnvioFalhou()
    audit.record(
        ctx, "usuario", usuario["id"], "solicitar_confirmacao_email",
        depois={"email": mascarar_email(email)},
    )
    return email


def confirmar(usuario: dict, codigo: str, ctx: AuditContext) -> str:
    """Confere o código e grava o e-mail como confirmado. Devolve o e-mail."""
    db = EmailConfirmacaoDB()
    pendente = db.pendente(usuario["id"])
    if (
        pendente is None
        or pendente["expira_em"] <= _agora_utc()
        or int(pendente["tentativas"]) >= MAX_TENTATIVAS
    ):
        raise CodigoInvalido()

    esperado = _hash_codigo(usuario["id"], pendente["email"], (codigo or "").strip())
    if not hmac.compare_digest(esperado, pendente["codigo_hash"]):
        if db.registrar_falha(pendente["id"]) >= MAX_TENTATIVAS:
            db.marcar_usado(pendente["id"])
            raise CodigoEsgotado()
        raise CodigoInvalido()

    if not db.marcar_usado(pendente["id"]):
        raise CodigoInvalido()  # outro pedido simultâneo já gastou o código
    try:
        UserDBManager().confirmar_email(usuario["id"], pendente["email"])
    except EmailEmUso:
        raise EmailEmUsoErro()
    audit.record(
        ctx, "usuario", usuario["id"], "confirmar_email",
        antes={"email": mascarar_email(usuario.get("email"))},
        depois={"email": mascarar_email(pendente["email"])},
    )
    return pendente["email"]
