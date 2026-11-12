from datetime import datetime
from typing import Optional

from pydantic import BaseModel


class AuditEntry(BaseModel):
    id: int
    ocorrido_em: datetime  # UTC
    usuario_id: Optional[int] = None
    username: Optional[str] = None
    papel: Optional[str] = None
    ip: Optional[str] = None
    entidade: str
    entidade_id: Optional[str] = None
    acao: str
    antes: Optional[dict] = None
    depois: Optional[dict] = None
