from pydantic import BaseModel


class LoanRequest(BaseModel):
    item_id: int
    usuario: str
    cpf: str
    center_cost: str
    cargo: str
    setor: str
    revenda: str
    date_issue: str  # dd/mm/yyyy
