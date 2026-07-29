from pydantic import BaseModel, field_validator
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
    operacao_anexo: Optional[str] = None
    termo_assinado_anexo: Optional[str] = None

    class Config:
        from_attributes = True


class ReverseRequest(BaseModel):
    """Body do endpoint de estorno: exige reconfirmação de senha do
    operador logado, já que é uma operação destrutiva sobre o histórico."""

    password: str

    @field_validator("password")
    @classmethod
    def _validar_password(cls, v: str) -> str:
        if not v:
            raise ValueError("Informe sua senha para confirmar o estorno.")
        return v
