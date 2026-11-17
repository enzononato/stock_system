"""Esqueci minha senha: pedido de link por e-mail e redefinição."""
import html
import logging
import secrets
from datetime import datetime, timedelta, timezone

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, Request

from app.core.config import settings
from app.core.security import sha256_hex
from app.db.password_reset_db import PasswordResetDBManager
from app.db.user_manager_db import UserDBManager
from app.dependencies import (
    contexto_auditoria_de,
    get_password_reset_db,
    get_user_db,
    ip_real_do_cliente,
    limiter,
)
from app.schemas.auth import ForgotRequest, MensagemResponse, ResetRequest
from app.services import audit
from app.services.mail import MailError, render_email, send_mail

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/auth", tags=["auth"])

RESET_VALIDADE_MINUTOS = 30
RESET_MAX_POR_HORA = 3
CONVITE_VALIDADE_HORAS = 72
MSG_FORGOT = "Se o usuário existir e tiver e-mail cadastrado, enviaremos um link para redefinir a senha."
MSG_LINK_INVALIDO = "Link inválido ou expirado. Peça um novo."
MSG_SENHA_REDEFINIDA = "Senha redefinida. Entre com a nova senha."


def _agora_utc() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def enviar_link_senha(email: str, username: str, token: str) -> None:
    """Envia o link normal de redefinição (30 min). Levanta `MailError` se falhar."""
    link = f"{settings.APP_PUBLIC_URL.rstrip('/')}/redefinir-senha?token={token}"
    corpo = (
        f"<p>Recebemos um pedido para redefinir a senha do usuário <strong>{html.escape(username)}</strong>.</p>"
        "<p>O link vale por 30 minutos e só pode ser usado uma vez. Se não foi você, "
        "ignore este e-mail: sua senha continua a mesma.</p>"
    )
    send_mail(
        email,
        "Redefinição de senha — Patrimônio Revalle",
        render_email("Redefinir senha", corpo, acao_url=link, acao_rotulo="Redefinir senha"),
    )


def enviar_convite(email: str, token: str) -> None:
    """Envia o convite de uma conta nova (link de 72 h para criar a senha). Levanta
    `MailError` se falhar. O `convite=1` no link só muda o título da tela."""
    link = f"{settings.APP_PUBLIC_URL.rstrip('/')}/redefinir-senha?token={token}&convite=1"
    corpo = (
        "<p>Você recebeu acesso ao sistema de Patrimônio da Revalle. Crie sua senha pelo botão abaixo. "
        "O link vale por 72 horas.</p>"
        "<p>Depois, entre com este e-mail e a senha que você criar.</p>"
    )
    send_mail(
        email,
        "Seu acesso ao Patrimônio Revalle",
        render_email("Crie sua senha", corpo, acao_url=link, acao_rotulo="Criar minha senha"),
    )


def emitir_link(usuario_id: int, convite: bool, reset_db: PasswordResetDBManager) -> str:
    """Novo link de senha emitido pelo sistema (convite ou envio do Gestor): os links
    abertos da conta deixam de valer. Convite vale 72 h; o link normal, 30 min. Não
    entra no limite de pedidos por hora do "Esqueci minha senha". Devolve o token."""
    agora = _agora_utc()
    validade = timedelta(hours=CONVITE_VALIDADE_HORAS) if convite else timedelta(minutes=RESET_VALIDADE_MINUTOS)
    token = secrets.token_urlsafe(32)
    # Invalida os anteriores e grava o novo numa transação só (linha do usuário travada).
    reset_db.substituir(usuario_id, sha256_hex(token), agora, agora + validade)
    return token


def enviar_email_redefinicao(email: str, username: str, token: str) -> None:
    """Envia o link de redefinição. Roda numa BackgroundTask: o tempo da resposta
    não depende do SMTP. Falhas só viram log com mensagem fixa (sem e-mail nem token)."""
    try:
        enviar_link_senha(email, username, token)
    except MailError:
        logger.error("Falha ao enviar o e-mail de redefinição de senha.")
    except Exception as exc:
        logger.error(
            "Erro inesperado ao enviar o e-mail de redefinição de senha (%s).", type(exc).__name__
        )


@router.post("/forgot", response_model=MensagemResponse)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def forgot(
    request: Request,
    body: ForgotRequest,
    background_tasks: BackgroundTasks,
    user_db: UserDBManager = Depends(get_user_db),
    reset_db: PasswordResetDBManager = Depends(get_password_reset_db),
):
    """Resposta sempre genérica e no mesmo tempo para qualquer entrada: o trabalho todo
    (busca, limite por usuário, gravação do link, auditoria, e-mail) roda depois da
    resposta, numa BackgroundTask, então o tempo não revela se o usuário existe."""
    background_tasks.add_task(
        _processar_pedido,
        body.identificador.strip(),
        ip_real_do_cliente(request),
        contexto_auditoria_de(None, request),
        user_db,
        reset_db,
    )
    return MensagemResponse(detail=MSG_FORGOT)


def _processar_pedido(
    identificador: str,
    ip: str,
    ctx,
    user_db: UserDBManager,
    reset_db: PasswordResetDBManager,
) -> None:
    try:
        usuario = None
        if identificador:
            if "@" in identificador:
                usuario = user_db.get_user_by_email(identificador)
            else:
                usuario = user_db.get_user_by_username(identificador)
        # Só para e-mail confirmado: um e-mail digitado errado (pelo Gestor ou na
        # conta antiga) nunca recebe o link.
        if not usuario or not usuario.get("email") or not usuario.get("email_confirmado"):
            return

        agora = _agora_utc()
        token = secrets.token_urlsafe(32)
        # Contagem + gravação atômicas (linha do usuário travada): no máximo
        # RESET_MAX_POR_HORA e-mails por hora, mesmo com pedidos simultâneos.
        criado = reset_db.criar_se_dentro_do_limite(
            usuario["id"],
            sha256_hex(token),
            agora,
            agora + timedelta(minutes=RESET_VALIDADE_MINUTOS),
            ip,
            agora - timedelta(hours=1),
            RESET_MAX_POR_HORA,
        )
        if not criado:
            return
        audit.record(
            ctx, "usuario", usuario["id"], "solicitar_redefinicao_senha", depois={"canal": "email"}
        )
        enviar_email_redefinicao(usuario["email"], usuario["username"], token)
    except Exception as exc:
        # Só o tipo do erro: nada de usuário, e-mail ou token no log.
        logger.error("Falha ao processar o pedido de redefinição de senha (%s).", type(exc).__name__)


@router.post("/reset", response_model=MensagemResponse)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def reset(
    request: Request,
    body: ResetRequest,
    user_db: UserDBManager = Depends(get_user_db),
    reset_db: PasswordResetDBManager = Depends(get_password_reset_db),
):
    # Mesma regra de força de hoje (senha em branco); checada ANTES de gastar o link.
    if not body.nova_senha:
        raise HTTPException(status_code=400, detail="A nova senha não pode estar em branco.")

    agora = _agora_utc()
    usuario_id = reset_db.consumir(sha256_hex(body.token), agora)
    if usuario_id is None:
        raise HTTPException(status_code=400, detail=MSG_LINK_INVALIDO)
    usuario = user_db.get_user_by_id(usuario_id)
    if usuario is None:
        raise HTTPException(status_code=400, detail=MSG_LINK_INVALIDO)

    ok, msg = user_db.definir_senha_pelo_link(usuario_id, body.nova_senha)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    reset_db.invalidar_abertos(usuario_id, agora)
    user_db.revogar_sessoes(usuario_id)
    audit.record(
        contexto_auditoria_de(usuario, request),
        "usuario", usuario_id, "redefinir_senha",
        depois={
            "senha": "redefinida" if usuario.get("senha_definida", 1) else "criada pelo convite",
            "sessoes": "encerradas",
            "email_confirmado": True,
        },
    )
    return MensagemResponse(detail=MSG_SENHA_REDEFINIDA)
