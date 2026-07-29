from pydantic import BaseModel, field_validator

from app.schemas.validators import validar_cpf, validar_data_br


class LoanRequest(BaseModel):
    item_id: int
    usuario: str
    cpf: str
    center_cost: str
    cargo: str
    setor: str
    revenda: str
    date_issue: str  # dd/mm/yyyy

    @field_validator("cpf")
    @classmethod
    def _validar_cpf(cls, v: str) -> str:
        return validar_cpf(v)

    @field_validator("date_issue")
    @classmethod
    def _validar_date_issue(cls, v: str) -> str:
        return validar_data_br(v, campo="data de empréstimo", obrigatorio=True)
