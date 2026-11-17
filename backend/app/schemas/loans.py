from typing import Optional

from pydantic import BaseModel, field_validator, model_validator

from app.schemas.validators import validar_cpf, validar_data_br


class LoanRequest(BaseModel):
    item_id: int
    # Colaborador do cadastro (F3). Quando informado, nome, CPF, cargo e setor vêm do
    # cadastro e os campos antigos abaixo são ignorados (cargo/setor só complementam
    # quando o cadastro não os tem). Sem ele, o payload antigo continua aceito.
    colaborador_id: Optional[int] = None
    usuario: Optional[str] = None
    cpf: Optional[str] = None
    center_cost: str
    cargo: Optional[str] = None
    setor: Optional[str] = None
    revenda: str
    date_issue: str  # dd/mm/yyyy
    # Desmarcado por padrão: o termo assume "pessoa física" (o caso comum, já que
    # o campo ao lado é CPF). Marcar apenas quando o recebedor for de fato
    # representante de pessoa jurídica.
    pessoa_juridica: bool = False

    @field_validator("cpf")
    @classmethod
    def _validar_cpf(cls, v: Optional[str]) -> Optional[str]:
        return None if v is None else validar_cpf(v)

    @field_validator("date_issue")
    @classmethod
    def _validar_date_issue(cls, v: str) -> str:
        return validar_data_br(v, campo="data de empréstimo", obrigatorio=True)

    @model_validator(mode="after")
    def _exigir_colaborador_ou_campos_antigos(self) -> "LoanRequest":
        if self.colaborador_id is None and any(
            valor is None for valor in (self.usuario, self.cpf, self.cargo, self.setor)
        ):
            raise ValueError("Informe o colaborador do cadastro ou nome, CPF, cargo e setor.")
        return self


class ReturnInitiateRequest(BaseModel):
    # Colaborador que está devolvendo (necessário para itens sem vínculo de cadastro).
    colaborador_id: Optional[int] = None
