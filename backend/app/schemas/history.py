from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class HistoryResponse(BaseModel):
    id: int
    item_id: Optional[int]
    peripheral_id: Optional[int]
    operador: Optional[str]
    operation: Optional[str]
    revenda: Optional[str]
    data_operacao: Optional[datetime]
    tipo: Optional[str]
    marca: Optional[str]
    modelo: Optional[str]
    nota_fiscal: Optional[str]
    fornecedor: Optional[str]
    identificador: Optional[str]
    usuario: Optional[str]
    cpf: Optional[str]
    cargo: Optional[str]
    center_cost: Optional[str]
    setor: Optional[str]
    details: Optional[str]

    class Config:
        from_attributes = True
