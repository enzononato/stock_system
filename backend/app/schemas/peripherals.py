from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class PeripheralCreate(BaseModel):
    tipo: str
    brand: Optional[str] = None
    model: Optional[str] = None
    identificador: Optional[str] = None


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
