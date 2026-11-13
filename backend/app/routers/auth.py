from datetime import datetime, timezone
from typing import Optional

import pymysql
from fastapi import APIRouter, HTTPException, Response, Cookie, status, Depends, Request

from app.schemas.auth import (
    CadastroTotpResponse,
    LoginRequest,
    LoginResponse,
    MfaConfirmarRequest,
    MfaTokenRequest,
    MfaVerificarRequest,
    TokenResponse,
    UserMe,
)
from app.core.security import (
    DUMMY_PASSWORD_HASH,
    MFA_CADASTRO_MINUTOS,
    MFA_VERIFICAR_MINUTOS,
    create_access_token,
    create_mfa_token,
    create_refresh_token,
    decode_token,
    verify_password,
)
from app.core.config import settings
from app.db.user_manager_db import UserDBManager
from app.db.token_db import TokenDBManager
from app.services.sessao import dados_do_token, exige_cadastro_2fa, validar_sessao_ativa
from app.db.mfa_db import MfaDBManager
from app.dependencies import (
    CurrentUser,
    contexto_auditoria_de,
    get_current_user,
    get_mfa_db,
    get_token_db,
    get_user_db,
    limiter,
)
from app.services.mfa import CodigoInvalido, MfaErro, TokenEsgotado
from app.services.segundo_fator import (
    confirmar_cadastro,
    exigir_token_nao_esgotado,
    exigir_verificacao_liberada,
    iniciar_cadastro,
    registrar_falha_segundo_fator,
    verificar_segundo_fator,
)

router = APIRouter(prefix="/api/auth", tags=["auth"])


def _set_refresh_cookie(response: Response, refresh_token: str) -> None:
    """Centraliza os atributos do cookie do refresh token (T1): secure e
    samesite vêm da configuração por ambiente, em vez de valor fixo no
    código — o que antes fazia o cookie nem ser aceito/enviado em cenários
    divergentes de dev/produção."""
    response.set_cookie(
        key="refresh_token",
        value=refresh_token,
        httponly=True,
        secure=settings.cookie_secure_effective,
        samesite=settings.COOKIE_SAMESITE,
        max_age=60 * 60 * 24 * settings.REFRESH_TOKEN_EXPIRE_DAYS,
    )


MSG_ETAPA_EXPIRADA = "Etapa de verificação expirada. Entre novamente."


def _emitir_sessao(response: Response, usuario: dict, **extras) -> LoginResponse:
    """Abre a sessão: access token no corpo e refresh token no cookie."""
    # O claim "sub" precisa ser string: é exigência do próprio JWT (RFC 7519)
    # e o python-jose valida isso ao decodificar. get_current_user converte
    # de volta para int ao montar o CurrentUser.
    token_data = dados_do_token(usuario)
    _set_refresh_cookie(response, create_refresh_token(token_data))
    return LoginResponse(etapa="ok", access_token=create_access_token(token_data), **extras)


def _decodificar_mfa(
    mfa_token: str, proposito: str, token_db: TokenDBManager, user_db: UserDBManager
) -> tuple[dict, dict]:
    """Valida o token da etapa de 2FA e devolve (payload, usuário).

    Token inválido/expirado, de outro tipo ou propósito, já usado, de usuário
    removido ou de sessões já encerradas: 400 (nunca 401 — o interceptor do
    frontend trata 401 como sessão expirada)."""
    try:
        payload = decode_token(mfa_token, expected_type="mfa_pendente")
    except HTTPException:
        raise HTTPException(status_code=400, detail=MSG_ETAPA_EXPIRADA)
    if payload.get("proposito") != proposito:
        raise HTTPException(status_code=400, detail=MSG_ETAPA_EXPIRADA)
    jti = payload.get("jti")
    if not jti or token_db.is_revoked(jti):
        raise HTTPException(status_code=400, detail=MSG_ETAPA_EXPIRADA)
    try:
        usuario = user_db.get_user_by_id(int(payload["sub"]))
    except (KeyError, TypeError, ValueError):
        usuario = None
    if usuario is None or int(payload.get("sv", 0)) != int(usuario.get("sessao_versao") or 0):
        raise HTTPException(status_code=400, detail=MSG_ETAPA_EXPIRADA)
    return payload, usuario


def _revogar_mfa(payload: dict, token_db: TokenDBManager) -> None:
    """Uso único: o `jti` do token da etapa é revogado quando ela é concluída."""
    token_db.revoke(payload["jti"], datetime.fromtimestamp(payload["exp"], tz=timezone.utc))


def _erro_mfa(exc: MfaErro) -> HTTPException:
    return HTTPException(status_code=exc.status_code, detail=str(exc))


@router.post("/login", response_model=LoginResponse)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def login(
    request: Request,
    body: LoginRequest,
    response: Response,
    user_db: UserDBManager = Depends(get_user_db),
):
    user = user_db.get_user_by_username(body.username)
    # Usuário inexistente também paga o custo do bcrypt: o tempo de resposta não
    # revela se o usuário existe.
    hash_a_conferir = user["password"] if user else DUMMY_PASSWORD_HASH
    senha_ok = verify_password(body.password, hash_a_conferir)
    if not user or not senha_ok:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Usuário ou senha inválidos.")

    if user.get("totp_ativo"):
        # Senha certa, mas falta o código: nenhum cookie nem access token ainda.
        # (Vale mesmo sem a chave TOTP no servidor: não há bypass para quem já tem 2FA.)
        return LoginResponse(
            etapa="mfa",
            mfa_token=create_mfa_token(dados_do_token(user), "verificar"),
            expira_em_segundos=MFA_VERIFICAR_MINUTOS * 60,
        )
    if exige_cadastro_2fa(user):
        # Gestor sem autenticador: não recebe access nem refresh token antes de cadastrar.
        return LoginResponse(
            etapa="cadastro_2fa",
            mfa_token=create_mfa_token(dados_do_token(user), "cadastrar"),
            expira_em_segundos=MFA_CADASTRO_MINUTOS * 60,
        )
    return _emitir_sessao(response, user)


@router.post("/mfa/verificar", response_model=LoginResponse)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def verificar_mfa(
    request: Request,
    body: MfaVerificarRequest,
    response: Response,
    user_db: UserDBManager = Depends(get_user_db),
    token_db: TokenDBManager = Depends(get_token_db),
    mfa_db: MfaDBManager = Depends(get_mfa_db),
):
    payload, usuario = _decodificar_mfa(body.mfa_token, "verificar", token_db, user_db)
    ctx = contexto_auditoria_de(usuario, request)
    erro: Optional[MfaErro] = None
    revogar_token = False
    restantes = None
    try:
        # Checar o bloqueio, verificar o código e registrar a falha/sucesso numa transação só,
        # com a linha do usuário travada: pedidos simultâneos do mesmo usuário se serializam
        # (nenhuma falha se perde nem é contada a menos), em qualquer worker.
        with mfa_db.transacao_do_usuario(usuario["id"]):
            try:
                exigir_verificacao_liberada(mfa_db, usuario)
                exigir_token_nao_esgotado(mfa_db, usuario, payload["jti"])
                restantes = verificar_segundo_fator(mfa_db, usuario, body.codigo, ctx)
                mfa_db.limpar_falhas(usuario["id"])
            except CodigoInvalido as exc:
                # Código errado: conta a falha. 5 no mesmo token o revogam; 10 do usuário em
                # 15 min bloqueiam a 2ª etapa por 15 min (contadores no banco).
                try:
                    registrar_falha_segundo_fator(mfa_db, usuario, payload["jti"], ctx)
                    erro = exc
                except MfaErro as limite:
                    erro, revogar_token = limite, True
            except MfaErro as exc:
                erro = exc
                revogar_token = isinstance(exc, TokenEsgotado)  # garante a revogação do token esgotado
    except pymysql.MySQLError as exc:
        if exc.args and exc.args[0] in (1205, 1213):  # espera de trava / deadlock: nunca 500
            raise HTTPException(
                status_code=409, detail="Não foi possível verificar agora. Tente novamente."
            )
        raise
    if erro is not None:
        if revogar_token:
            _revogar_mfa(payload, token_db)
        raise _erro_mfa(erro)
    _revogar_mfa(payload, token_db)
    return _emitir_sessao(response, usuario, codigos_recuperacao_restantes=restantes)


@router.post("/mfa/cadastro/iniciar", response_model=CadastroTotpResponse)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def iniciar_cadastro_mfa(
    request: Request,
    body: MfaTokenRequest,
    user_db: UserDBManager = Depends(get_user_db),
    token_db: TokenDBManager = Depends(get_token_db),
    mfa_db: MfaDBManager = Depends(get_mfa_db),
):
    """Gestor sem 2FA, durante o login: gera o segredo e o QR. Não gasta o token da etapa."""
    _, usuario = _decodificar_mfa(body.mfa_token, "cadastrar", token_db, user_db)
    try:
        segredo, uri = iniciar_cadastro(mfa_db, usuario)
    except MfaErro as exc:
        raise _erro_mfa(exc)
    return CadastroTotpResponse(segredo=segredo, otpauth_uri=uri)


@router.post("/mfa/cadastro/confirmar", response_model=LoginResponse)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def confirmar_cadastro_mfa(
    request: Request,
    body: MfaConfirmarRequest,
    response: Response,
    user_db: UserDBManager = Depends(get_user_db),
    token_db: TokenDBManager = Depends(get_token_db),
    mfa_db: MfaDBManager = Depends(get_mfa_db),
):
    """Confirma o código, ativa o 2FA, abre a sessão e devolve os códigos de recuperação (uma vez)."""
    payload, usuario = _decodificar_mfa(body.mfa_token, "cadastrar", token_db, user_db)
    try:
        codigos = confirmar_cadastro(mfa_db, usuario, body.codigo, contexto_auditoria_de(usuario, request))
    except MfaErro as exc:
        raise _erro_mfa(exc)
    _revogar_mfa(payload, token_db)
    return _emitir_sessao(response, usuario, codigos_recuperacao=codigos)


@router.post("/refresh", response_model=TokenResponse)
def refresh(
    response: Response,
    refresh_token: Optional[str] = Cookie(default=None),
    token_db: TokenDBManager = Depends(get_token_db),
):
    if not refresh_token:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Refresh token não encontrado.")
    payload = decode_token(refresh_token, expected_type="refresh")

    jti = payload.get("jti")
    if jti and token_db.is_revoked(jti):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Sessão encerrada. Faça login novamente.",
        )

    validar_sessao_ativa(payload)

    token_data = {
        "sub": payload["sub"], "username": payload["username"], "role": payload["role"],
        "sv": payload.get("sv", 0),
    }
    new_access = create_access_token(token_data)
    new_refresh = create_refresh_token(token_data)

    # Rotação com invalidação (T3): o refresh token antigo é revogado assim
    # que um novo é emitido, para que não possa ser reaproveitado caso tenha
    # vazado (ex: replay de um token capturado).
    if jti:
        expires_at = datetime.fromtimestamp(payload["exp"], tz=timezone.utc)
        token_db.revoke(jti, expires_at)

    _set_refresh_cookie(response, new_refresh)
    return TokenResponse(access_token=new_access)


@router.post("/logout")
def logout(
    response: Response,
    refresh_token: Optional[str] = Cookie(default=None),
    token_db: TokenDBManager = Depends(get_token_db),
):
    """Invalida o refresh token atual (se houver) e remove o cookie (T3).

    Idempotente: cookie ausente, expirado ou já inválido não impedem o
    logout de retornar sucesso — o objetivo (usuário deslogado) é sempre
    alcançado, com ou sem revogação efetiva no banco.
    """
    if refresh_token:
        try:
            payload = decode_token(refresh_token, expected_type="refresh")
            jti = payload.get("jti")
            if jti:
                expires_at = datetime.fromtimestamp(payload["exp"], tz=timezone.utc)
                token_db.revoke(jti, expires_at)
        except HTTPException:
            # Token ausente/expirado/malformado: nada a revogar, mas o
            # logout deve seguir normalmente.
            pass

    response.delete_cookie("refresh_token")
    return {"detail": "Logout realizado com sucesso."}


@router.get("/me", response_model=UserMe)
def me(current_user: CurrentUser = Depends(get_current_user)):
    return UserMe(id=current_user.id, username=current_user.username, role=current_user.role)
