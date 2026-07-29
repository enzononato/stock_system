from pydantic import BaseModel, field_validator
from typing import Optional
from datetime import datetime

from app.schemas.validators import blank_to_none


class PeripheralCreate(BaseModel):
    tipo: str
    brand: Optional[str] = None
    model: Optional[str] = None
    identificador: Optional[str] = None

    @field_validator("tipo")
    @classmethod
    def _validar_tipo(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("Tipo do periférico é obrigatório.")
        return v

    @field_validator("brand", "model", "identificador")
    @classmethod
    def _normalizar_opcionais(cls, v):
        # Evita gravar "" no lugar de None: identificador tem UNIQUE no banco
        # e strings vazias repetidas causariam falha de unicidade.
        return blank_to_none(v)


class PeripheralResponse(BaseModel):
    id: int
    tipo: str
    brand: Optional[str]
    model: Optional[str]
    identificador: Optional[str]
    status: Optional[str]
    motivo_substituicao: Optional[str]
    date_registered: Optional[datetime]
    link_id: Optional[int] = None

    class Config:
        from_attributes = True


class ReplacePeripheralRequest(BaseModel):
    new_peripheral_id: int
    reason: str

    @field_validator("reason")
    @classmethod
    def _validar_reason(cls, v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("Informe o motivo da substituição.")
        return v
