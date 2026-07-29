from datetime import datetime, timedelta, timezone
from typing import Optional
from uuid import uuid4
import bcrypt
from jose import JWTError, jwt
from fastapi import HTTPException, status

from app.core.config import settings


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_password(plain: str, hashed: Optional[str]) -> bool:
    """Verifica a senha em texto puro contra o hash armazenado.

    Retorna False (em vez de propagar exceção) quando o hash salvo não está
    em um formato bcrypt válido — cenário de usuários legados/corrompidos.
    Sem isso, `bcrypt.checkpw` levanta ValueError e o endpoint de login
    devolveria 500 em vez de 401 para essas contas.
    """
    if not plain or not hashed:
        return False
    try:
        return bcrypt.checkpw(plain.encode("utf-8"), hashed.encode("utf-8"))
    except (ValueError, TypeError):
        return False


def create_access_token(data: dict) -> str:
    payload = data.copy()
    payload["exp"] = datetime.now(timezone.utc) + timedelta(minutes=settings.ACCESS_TOKEN_EXPIRE_MINUTES)
    payload["type"] = "access"
    return jwt.encode(payload, settings.JWT_SECRET, algorithm=settings.JWT_ALGORITHM)


def create_refresh_token(data: dict) -> str:
    """Cria o refresh token com um claim `jti` (identificador único). O jti
    permite revogar tokens individualmente no logout/rotação, sem precisar
    invalidar o segredo JWT inteiro para todos os usuários."""
    payload = data.copy()
    payload["exp"] = datetime.now(timezone.utc) + timedelta(days=settings.REFRESH_TOKEN_EXPIRE_DAYS)
    payload["type"] = "refresh"
    payload["jti"] = uuid4().hex
    return jwt.encode(payload, settings.JWT_SECRET, algorithm=settings.JWT_ALGORITHM)


def decode_token(token: str, expected_type: str = "access") -> dict:
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Token inválido ou expirado.",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, settings.JWT_SECRET, algorithms=[settings.JWT_ALGORITHM])
        if payload.get("type") != expected_type:
            raise credentials_exception
        return payload
    except JWTError:
        raise credentials_exception
