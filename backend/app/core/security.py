import hashlib
from datetime import datetime, timedelta, timezone
from typing import Literal, Optional
from uuid import uuid4
import bcrypt
from jose import JWTError, jwt
from fastapi import HTTPException, status

from app.core.config import settings


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


# Hash bcrypt (mesmo custo dos reais) usado quando o usuário não existe, para o login
# gastar o mesmo tempo com usuário válido e inválido. Calculado uma vez na importação.
DUMMY_PASSWORD_HASH = hash_password("senha-fictícia-para-igualar-o-tempo")


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


def sha256_hex(texto: str) -> str:
    """SHA-256 em hexadecimal (64 caracteres). Para tokens e códigos de alta
    entropia (redefinição de senha, códigos de recuperação), que dispensam bcrypt."""
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()


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


MFA_VERIFICAR_MINUTOS = 5
MFA_CADASTRO_MINUTOS = 15


def create_mfa_token(data: dict, proposito: Literal["verificar", "cadastrar", "email"]) -> str:
    """Token temporário da etapa de 2FA do login (`type="mfa_pendente"`), de uso
    único (o `jti` é revogado no sucesso) e que não serve como access token.
    `verificar` dura 5 min; `cadastrar` (Gestor sem autenticador) dura 15 min,
    porque instalar o aplicativo leva mais tempo; `email` (confirmar o e-mail de
    login com o código enviado) também dura 15 min. O propósito vai no token e cada
    rota só aceita o seu."""
    minutos = MFA_VERIFICAR_MINUTOS if proposito == "verificar" else MFA_CADASTRO_MINUTOS
    payload = data.copy()
    payload["exp"] = datetime.now(timezone.utc) + timedelta(minutes=minutos)
    payload["type"] = "mfa_pendente"
    payload["proposito"] = proposito
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
