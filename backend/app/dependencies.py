from fastapi import Depends, HTTPException, status, Cookie
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from dataclasses import dataclass
from typing import Optional

from app.core.security import decode_token

bearer_scheme = HTTPBearer(auto_error=False)


@dataclass
class CurrentUser:
    id: int
    username: str
    role: str


def get_current_user(
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(bearer_scheme),
) -> CurrentUser:
    if not credentials:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token de autenticação não fornecido.",
            headers={"WWW-Authenticate": "Bearer"},
        )
    payload = decode_token(credentials.credentials, expected_type="access")
    return CurrentUser(
        id=int(payload["sub"]),
        username=payload["username"],
        role=payload["role"],
    )


def require_role(*allowed_roles: str):
    """Factory que retorna uma dependency que exige um dos roles listados."""
    def checker(user: CurrentUser = Depends(get_current_user)) -> CurrentUser:
        if user.role not in allowed_roles:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"Acesso negado. Roles permitidos: {', '.join(allowed_roles)}.",
            )
        return user
    return checker


# Atalhos prontos para uso nos routers
def gestor_only(user: CurrentUser = Depends(get_current_user)) -> CurrentUser:
    if user.role != "Gestor":
        raise HTTPException(status_code=403, detail="Acesso restrito ao Gestor.")
    return user


def gestor_or_tecnico(user: CurrentUser = Depends(get_current_user)) -> CurrentUser:
    if user.role not in ("Gestor", "Técnico"):
        raise HTTPException(status_code=403, detail="Acesso restrito a Gestor ou Técnico.")
    return user
