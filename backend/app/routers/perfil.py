"""Perfil do usuário logado e 2FA opcional (para quem não é obrigado a tê-lo)."""
from fastapi import APIRouter, Depends, HTTPException, Request

from app.core.config import settings
from app.core.security import verify_password
from app.db.mfa_db import MfaDBManager
from app.db.user_manager_db import UserDBManager
from app.dependencies import (
    CurrentUser,
    get_audit_context,
    get_current_user,
    get_mfa_db,
    get_user_db,
    limiter,
)
from app.schemas.auth import CadastroTotpResponse, MensagemResponse
from app.schemas.perfil import CodigoRequest, CodigosRecuperacaoResponse, PerfilResponse, SenhaRequest
from app.services import segundo_fator
from app.services.audit import AuditContext
from app.services.mfa import MfaErro, mfa_configurado

router = APIRouter(prefix="/api/perfil", tags=["perfil"])


def _usuario(user_db: UserDBManager, atual: CurrentUser) -> dict:
    usuario = user_db.get_user_by_id(atual.id)
    if usuario is None:
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    return usuario


def _erro(exc: MfaErro) -> HTTPException:
    return HTTPException(status_code=exc.status_code, detail=str(exc))


def _exigir_senha(user_db: UserDBManager, atual: CurrentUser, senha: str) -> None:
    com_hash = user_db.get_user_by_username(atual.username)
    if not com_hash or not verify_password(senha, com_hash["password"]):
        raise HTTPException(status_code=400, detail="Senha incorreta.")


@router.get("", response_model=PerfilResponse)
def obter_perfil(
    atual: CurrentUser = Depends(get_current_user),
    user_db: UserDBManager = Depends(get_user_db),
    mfa_db: MfaDBManager = Depends(get_mfa_db),
):
    usuario = _usuario(user_db, atual)
    ativo = bool(usuario.get("totp_ativo"))
    return PerfilResponse(
        id=usuario["id"],
        username=usuario["username"],
        role=usuario["role"],
        email=usuario.get("email"),
        totp_ativo=ativo,
        mfa_obrigatorio=usuario["role"] == "Gestor",
        mfa_disponivel=mfa_configurado(),
        codigos_recuperacao_restantes=mfa_db.contar_codigos_restantes(usuario["id"]) if ativo else 0,
    )


@router.post("/2fa/iniciar", response_model=CadastroTotpResponse)
def iniciar_2fa(
    atual: CurrentUser = Depends(get_current_user),
    user_db: UserDBManager = Depends(get_user_db),
    mfa_db: MfaDBManager = Depends(get_mfa_db),
):
    try:
        segredo, uri = segundo_fator.iniciar_cadastro(mfa_db, _usuario(user_db, atual))
    except MfaErro as exc:
        raise _erro(exc)
    return CadastroTotpResponse(segredo=segredo, otpauth_uri=uri)


@router.post("/2fa/confirmar", response_model=CodigosRecuperacaoResponse)
def confirmar_2fa(
    body: CodigoRequest,
    atual: CurrentUser = Depends(get_current_user),
    user_db: UserDBManager = Depends(get_user_db),
    mfa_db: MfaDBManager = Depends(get_mfa_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        codigos = segundo_fator.confirmar_cadastro(mfa_db, _usuario(user_db, atual), body.codigo, ctx)
    except MfaErro as exc:
        raise _erro(exc)
    return CodigosRecuperacaoResponse(codigos_recuperacao=codigos)


@router.post("/2fa/desativar", response_model=MensagemResponse)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def desativar_2fa(
    request: Request,
    body: SenhaRequest,
    atual: CurrentUser = Depends(get_current_user),
    user_db: UserDBManager = Depends(get_user_db),
    mfa_db: MfaDBManager = Depends(get_mfa_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    _exigir_senha(user_db, atual, body.senha)
    try:
        segundo_fator.desativar_2fa(mfa_db, _usuario(user_db, atual), ctx)
    except MfaErro as exc:
        raise _erro(exc)
    return MensagemResponse(detail="Verificação em duas etapas desativada.")


@router.post("/2fa/codigos", response_model=CodigosRecuperacaoResponse)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def regenerar_codigos(
    request: Request,
    body: SenhaRequest,
    atual: CurrentUser = Depends(get_current_user),
    user_db: UserDBManager = Depends(get_user_db),
    mfa_db: MfaDBManager = Depends(get_mfa_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    _exigir_senha(user_db, atual, body.senha)
    try:
        codigos = segundo_fator.regenerar_codigos(mfa_db, _usuario(user_db, atual), ctx)
    except MfaErro as exc:
        raise _erro(exc)
    return CodigosRecuperacaoResponse(codigos_recuperacao=codigos)
