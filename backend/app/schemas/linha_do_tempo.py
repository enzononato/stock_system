from datetime import datetime
from typing import Optional

from pydantic import BaseModel


class EventoResponse(BaseModel):
    ocorrido_em: datetime
    origem: str
    tipo: str
    titulo: str
    descricao: Optional[str] = None
    usuario: Optional[str] = None
    referencia_id: Optional[int] = None


class LinhaDoTempoResponse(BaseModel):
    item_id: int
    eventos: list[EventoResponse]
    fontes_com_erro: list[str]
