import logging
import secrets

from fastapi import APIRouter, HTTPException, Depends

from app.schemas.users import UserCreate, UserResponse, UserUpdate
from app.db.user_manager_db import UserDBManager
from app.db.mfa_db import MfaDBManager
from app.db.password_reset_db import PasswordResetDBManager
from app.dependencies import (
    gestor_only,
    get_user_db,
    get_mfa_db,
    get_audit_context,
    get_password_reset_db,
    CurrentUser,
)
from app.routers import auth_senha
from app.services import audit, segundo_fator
from app.services.audit import AuditContext
from app.services.confirmacao_email import mascarar_email

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/users", tags=["users"])

MSG_FALHA_CONVITE = "Conta criada, mas não foi possível enviar o e-mail. Use “Reenviar convite”."
MSG_FALHA_ENVIO = "Não foi possível enviar o e-mail agora. Tente de novo em instantes."
MSG_SEM_EMAIL = "Cadastre o e-mail desta conta primeiro."


@router.get("", response_model=list[UserResponse])
def list_users(_: CurrentUser = Depends(gestor_only), user_db: UserDBManager = Depends(get_user_db)):
    return user_db.get_all_users()


def _enviar_link(alvo: dict, reset_db: PasswordResetDBManager, ctx: AuditContext) -> bool:
    """Convite (conta sem senha) ou link normal, para o e-mail atual da conta. Os links
    anteriores deixam de valer. Devolve False se o e-mail não saiu."""
    convite = not alvo.get("senha_definida", 1)
    token = auth_senha.emitir_link(alvo["id"], convite, reset_db)
    try:
        if convite:
            auth_senha.enviar_convite(alvo["email"], token)
        else:
            auth_senha.enviar_link_senha(alvo["email"], alvo["username"], token)
    except Exception as exc:
        # Só o tipo do erro: nada de e-mail nem token no log.
        logger.error("Falha ao enviar o link de senha (%s).", type(exc).__name__)
        return False
    audit.record(
        ctx, "usuario", alvo["id"], "enviar_convite" if convite else "enviar_link_senha",
        depois={"email": mascarar_email(alvo["email"])},
    )
    return True


@router.post("", response_model=dict)
def create_user(
    body: UserCreate,
    _: CurrentUser = Depends(gestor_only),
    user_db: UserDBManager = Depends(get_user_db),
    reset_db: PasswordResetDBManager = Depends(get_password_reset_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    """Cria a conta pelo e-mail e envia o convite. Ninguém define a senha de outra
    pessoa: a conta nasce com uma senha inutilizável e `senha_definida = 0`."""
    username = user_db.gerar_username(body.email)
    ok, msg = user_db.add_user(
        username, secrets.token_urlsafe(32), body.role, body.email, senha_definida=False
    )
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    # get_user_by_username traz o hash da senha: relê sem ele.
    criado = user_db.get_user_by_id(user_db.get_user_by_username(username)["id"])
    # O usuário já foi gravado: a auditoria não pode derrubar a resposta.
    try:
        audit.record(
            ctx, "usuario", criado["id"], "criar",
            depois={"id": criado["id"], "username": username, "role": body.role, "email": body.email},
        )
    except Exception:
        logger.exception("Falha ao preparar auditoria (entidade=usuario, acao=criar).")
    if not _enviar_link(criado, reset_db, ctx):
        return {"detail": MSG_FALHA_CONVITE, "convite_enviado": False}
    return {"detail": f"Conta criada. Enviamos o convite para {body.email}.", "convite_enviado": True}


@router.post("/{user_id}/enviar-link-senha")
def enviar_link_senha(
    user_id: int,
    _: CurrentUser = Depends(gestor_only),
    user_db: UserDBManager = Depends(get_user_db),
    reset_db: PasswordResetDBManager = Depends(get_password_reset_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    """O Gestor não define senhas: manda um novo convite (conta sem senha) ou um link
    para criar nova senha, para o e-mail atual da conta."""
    alvo = user_db.get_user_by_id(user_id)
    if not alvo:
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    if not alvo.get("email"):
        raise HTTPException(status_code=400, detail=MSG_SEM_EMAIL)
    if not _enviar_link(alvo, reset_db, ctx):
        raise HTTPException(status_code=503, detail=MSG_FALHA_ENVIO)
    if not alvo.get("senha_definida", 1):
        return {"detail": f"Enviamos o convite para {alvo['email']}."}
    return {"detail": f"Enviamos o link para {alvo['email']}."}


@router.put("/{user_id}", response_model=UserResponse)
def update_user(
    user_id: int,
    body: UserUpdate,
    _: CurrentUser = Depends(gestor_only),
    user_db: UserDBManager = Depends(get_user_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    """Só o Gestor edita o e-mail (login e recuperação de senha). Um e-mail novo volta a
    `email_confirmado = 0`: a pessoa confirma no próximo login."""
    antes = user_db.get_user_by_id(user_id)
    if not antes:
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    ok, msg = user_db.update_email(user_id, body.email)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    audit.record_diff(
        ctx, "usuario", user_id, "editar",
        {"email": antes.get("email")},
        depois={"email": body.email},
    )
    atual = user_db.get_user_by_id(user_id) or {}
    return UserResponse(
        id=atual.get("id", user_id),
        username=atual.get("username", antes["username"]),
        role=atual.get("role", antes["role"]),
        email=atual.get("email"),
        email_confirmado=bool(atual.get("email_confirmado")),
        senha_definida=bool(atual.get("senha_definida", 1)),
        totp_ativo=bool(atual.get("totp_ativo")),
    )


@router.delete("/{user_id}/2fa")
def zerar_2fa_do_usuario(
    user_id: int,
    current_user: CurrentUser = Depends(gestor_only),
    user_db: UserDBManager = Depends(get_user_db),
    mfa_db: MfaDBManager = Depends(get_mfa_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    """O Gestor zera o 2FA de outro usuário (perdeu o celular e os códigos). Não precisa da chave TOTP."""
    if user_id == current_user.id:
        raise HTTPException(
            status_code=400,
            detail=(
                "Você não pode zerar a sua própria verificação em duas etapas. "
                "Peça a outro Gestor ou use o comando no servidor."
            ),
        )
    alvo = user_db.get_user_by_id(user_id)
    if not alvo:
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    if not alvo.get("totp_ativo"):
        raise HTTPException(
            status_code=400, detail="Este usuário não tem a verificação em duas etapas ativa."
        )
    segundo_fator.zerar_2fa(mfa_db, alvo, ctx, origem="tela de usuários")
    return {"detail": "Verificação em duas etapas zerada."}


@router.delete("/{user_id}")
def remove_user(
    user_id: int,
    current_user: CurrentUser = Depends(gestor_only),
    user_db: UserDBManager = Depends(get_user_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    # T5: um Gestor não pode remover a si mesmo, nem remover o último
    # Gestor do sistema (o que causaria lockout — ninguém mais poderia
    # administrar usuários).
    if user_id == current_user.id:
        raise HTTPException(status_code=400, detail="Você não pode remover seu próprio usuário.")

    target = user_db.get_user_by_id(user_id)
    if not target:
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")

    if target.get("role") == "Gestor" and user_db.count_gestores() <= 1:
        raise HTTPException(status_code=400, detail="Não é possível remover o último Gestor do sistema.")

    ok, msg = user_db.remove_user(user_id)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    audit.record(
        ctx, "usuario", user_id, "excluir",
        antes={"id": target["id"], "username": target["username"], "role": target["role"]},
    )
    return {"detail": msg}


# Observação: não existe endpoint de alteração de role de usuário (apenas criação,
# edição do e-mail, envio de link de senha e remoção). A troca de senha pelo Gestor
# (`PUT /api/users/{id}/password`) foi removida: ninguém sabe a senha de outra pessoa,
# o Gestor só envia o link. Caso um endpoint de role seja adicionado no futuro,
# aplique a mesma proteção usada em remove_user: um Gestor não pode rebaixar a si
# mesmo nem rebaixar o último Gestor do sistema (mesmo risco de lockout administrativo).
