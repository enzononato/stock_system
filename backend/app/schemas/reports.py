from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class ReportRow(BaseModel):
    history_id: Optional[int]
    item_id: Optional[int]
    peripheral_id: Optional[int]
    operador: Optional[str]
    usuario: Optional[str]
    cpf: Optional[str]
    cargo: Optional[str]
    center_cost: Optional[str]
    setor: Optional[str]
    fornecedor: Optional[str]
    revenda: Optional[str]
    details: Optional[str]
    data_emprestimo: Optional[datetime]
    data_confirmacao: Optional[datetime]
    data_devolucao: Optional[datetime]
    operation_type: Optional[str]
    tipo: Optional[str]
    brand: Optional[str]
    model: Optional[str]
    identificador: Optional[str]
    nota_fiscal: Optional[str]

    class Config:
        from_attributes = True


class ChartData(BaseModel):
    days: list[int]
    values: list[int]
    values2: Optional[list[int]] = None  # Para o gráfico de empréstimos x devoluções
