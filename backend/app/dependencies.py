import ipaddress
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
from app.db.colaborador_db import ColaboradorDBManager
from app.db.audit_db import AuditDBManager
from app.db.local_db import LocalDBManager
from app.db.linha_db import LinhaDBManager
from app.db.inventario_db import InventarioDBManager
from app.services.audit import AuditContext

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


@lru_cache
def get_audit_db() -> AuditDBManager:
    return AuditDBManager()


@lru_cache
def get_local_db() -> LocalDBManager:
    return LocalDBManager()


@lru_cache
def get_colaborador_db() -> ColaboradorDBManager:
    return ColaboradorDBManager()


@lru_cache
def get_linha_db() -> LinhaDBManager:
    return LinhaDBManager()


@lru_cache
def get_inventario_db() -> InventarioDBManager:
    return InventarioDBManager()


# ── Rate limiting (T2) ──────────────────────────────────────────────────────
# A aplicação roda atrás de nginx: o IP do cliente real chega no header
# X-Forwarded-For. Usar apenas request.client.host (padrão do slowapi)
# rate-limitaria o proxy inteiro como se fosse um único cliente.

def _parse_ip(texto: str):
    try:
        return ipaddress.ip_address(texto.strip())
    except ValueError:
        return None


def _ip_confiavel(ip) -> bool:
    """Loopback ou rede privada (o nginx do compose)."""
    return ip is not None and (ip.is_loopback or ip.is_private)


def ip_real_do_cliente(request: Request) -> str:
    """IP do cliente real, sem confiar em X-Forwarded-For forjado.

    O nginx do compose usa `$proxy_add_x_forwarded_for`, que ACRESCENTA o IP
    do cliente ao que ele mandou: a cadeia chega como "<forjado>, <real>" e o
    primeiro item é controlado pelo atacante. Por isso:
    - se o par direto não é proxy confiável (loopback/privado), vale o par e o
      header é ignorado;
    - senão, percorre a cadeia da DIREITA para a ESQUERDA e devolve o primeiro
      IP que não é confiável (entradas ilegíveis são puladas);
    - se todas forem confiáveis, devolve a mais à esquerda que seja válida;
      sem nenhuma válida, o par direto.
    """
    peer = request.client.host if request.client else get_remote_address(request)
    if not _ip_confiavel(_parse_ip(peer)):
        return peer
    cadeia = request.headers.get("X-Forwarded-For", "")
    validos = [ip for ip in (_parse_ip(p) for p in cadeia.split(",")) if ip is not None]
    for ip in reversed(validos):
        if not _ip_confiavel(ip):
            return str(ip)
    if validos:
        return str(validos[0])
    return peer


def get_client_ip(request: Request) -> str:
    return ip_real_do_cliente(request)


def get_audit_context(
    request: Request,
    user: CurrentUser = Depends(get_current_user),
) -> AuditContext:
    """Quem está agindo e de onde, para `audit.record(...)`."""
    return AuditContext(
        usuario_id=user.id,
        username=user.username,
        papel=user.role,
        ip=ip_real_do_cliente(request),
    )


limiter = Limiter(key_func=get_client_ip)
