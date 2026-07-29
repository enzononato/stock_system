from fastapi import APIRouter, HTTPException, Response, Cookie, status
from typing import Optional

from app.schemas.auth import LoginRequest, TokenResponse, UserMe
from app.core.security import verify_password, create_access_token, create_refresh_token, decode_token
from app.db.user_manager_db import UserDBManager
from app.dependencies import get_current_user, CurrentUser
from fastapi import Depends

router = APIRouter(prefix="/api/auth", tags=["auth"])
user_db = UserDBManager()


@router.post("/login", response_model=TokenResponse)
def login(body: LoginRequest, response: Response):
    user = user_db.get_user_by_username(body.username)
    if not user or not verify_password(body.password, user["password"]):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Usuário ou senha inválidos.")

    token_data = {"sub": str(user["id"]), "username": user["username"], "role": user["role"]}
    access_token = create_access_token(token_data)
    refresh_token = create_refresh_token(token_data)

    # Refresh token em cookie HttpOnly
    response.set_cookie(
        key="refresh_token",
        value=refresh_token,
        httponly=True,
        secure=False,  # False para dev (HTTP); True em produção (HTTPS)
        samesite="lax",
        max_age=60 * 60 * 24 * 7,
    )
    return TokenResponse(access_token=access_token)


@router.post("/refresh", response_model=TokenResponse)
def refresh(response: Response, refresh_token: Optional[str] = Cookie(default=None)):
    if not refresh_token:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Refresh token não encontrado.")
    payload = decode_token(refresh_token, expected_type="refresh")
    token_data = {"sub": payload["sub"], "username": payload["username"], "role": payload["role"]}
    new_access = create_access_token(token_data)
    new_refresh = create_refresh_token(token_data)
    response.set_cookie(
        key="refresh_token",
        value=new_refresh,
        httponly=True,
        secure=False,  # False para dev (HTTP); True em produção (HTTPS)
        samesite="lax",
        max_age=60 * 60 * 24 * 7,
    )
    return TokenResponse(access_token=new_access)


@router.post("/logout")
def logout(response: Response):
    response.delete_cookie("refresh_token")
    return {"detail": "Logout realizado com sucesso."}


@router.get("/me", response_model=UserMe)
def me(current_user: CurrentUser = Depends(get_current_user)):
    return UserMe(id=current_user.id, username=current_user.username, role=current_user.role)
