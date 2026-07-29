"""Schemas genéricos, reaproveitados por múltiplos routers."""
from typing import Generic, List, TypeVar
from pydantic import BaseModel

T = TypeVar("T")


class Paginated(BaseModel, Generic[T]):
    """Envelope padrão de resposta paginada: {"items": [...], "total": N}."""

    items: List[T]
    total: int
