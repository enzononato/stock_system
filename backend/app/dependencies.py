from functools import lru_cache

from fastapi import Depends, HTTPException, status, Request
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from dataclasses import dataclass
from typing import Optional

from slowapi import Limiter
from slowapi.util import get_remote_address

from app.core.security import decode_token
from app.db.inventory_manager_db import InventoryDBManager
from app.db.user_manager_db import UserDBManager
from app.db.token_db import TokenDBManager
from app.db.unidade_db import UnidadeDBManager

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
    # O claim "sub" é sempre uma string no token (exigência do próprio padrão
    # JWT/python-jose); convertemos de volta para int, já que o id do
    # usuário no banco é numérico.
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


# ── Injeção preguiçosa da camada de dados (T11) ─────────────────────────────
# Os managers só tocam o banco de fato dentro de __init__ (que cria tabelas)
# e a cada chamada de método. Antes, `inv = InventoryDBManager()` era
# executado no nível do módulo de cada router — ou seja, na importação do
# pacote, o que roda ANTES da aplicação subir. Se o MySQL estivesse
# indisponível nesse instante, o processo inteiro morria e nem o healthcheck
# respondia. Instanciando os managers apenas na primeira requisição que os
# usa (via Depends + lru_cache, o que também garante uma única instância
# reaproveitada no processo, evitando reconectar a cada chamada), o boot da
# aplicação nunca depende do banco estar de pé.

@lru_cache
def get_inventory_db() -> InventoryDBManager:
    return InventoryDBManager()


@lru_cache
def get_user_db() -> UserDBManager:
    return UserDBManager()


@lru_cache
def get_token_db() -> TokenDBManager:
    return TokenDBManager()


@lru_cache
def get_unidade_db() -> UnidadeDBManager:
    return UnidadeDBManager()


# ── Rate limiting (T2) ──────────────────────────────────────────────────────
# A aplicação roda atrás de nginx: o IP do cliente real chega no header
# X-Forwarded-For. Usar apenas request.client.host (padrão do slowapi)
# rate-limitaria o proxy inteiro como se fosse um único cliente.

def get_client_ip(request: Request) -> str:
    forwarded_for = request.headers.get("X-Forwarded-For")
    if forwarded_for:
        # O header pode conter uma cadeia "cliente, proxy1, proxy2"; o
        # primeiro IP é o do cliente original.
        return forwarded_for.split(",")[0].strip()
    return get_remote_address(request)


limiter = Limiter(key_func=get_client_ip)
