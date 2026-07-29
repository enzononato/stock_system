from datetime import datetime, timezone
from typing import Optional

from fastapi import APIRouter, HTTPException, Response, Cookie, status, Depends, Request

from app.schemas.auth import LoginRequest, TokenResponse, UserMe
from app.core.security import verify_password, create_access_token, create_refresh_token, decode_token
from app.core.config import settings
from app.db.user_manager_db import UserDBManager
from app.db.token_db import TokenDBManager
from app.dependencies import get_current_user, get_user_db, get_token_db, limiter, CurrentUser

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


@router.post("/login", response_model=TokenResponse)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def login(
    request: Request,
    body: LoginRequest,
    response: Response,
    user_db: UserDBManager = Depends(get_user_db),
):
    user = user_db.get_user_by_username(body.username)
    if not user or not verify_password(body.password, user["password"]):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Usuário ou senha inválidos.")

    # O claim "sub" precisa ser string: é exigência do próprio JWT (RFC 7519)
    # e o python-jose valida isso ao decodificar. get_current_user converte
    # de volta para int ao montar o CurrentUser.
    token_data = {"sub": str(user["id"]), "username": user["username"], "role": user["role"]}
    access_token = create_access_token(token_data)
    refresh_token = create_refresh_token(token_data)

    _set_refresh_cookie(response, refresh_token)
    return TokenResponse(access_token=access_token)


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

    token_data = {"sub": payload["sub"], "username": payload["username"], "role": payload["role"]}
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
