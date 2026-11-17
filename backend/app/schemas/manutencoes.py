from datetime import date, datetime
from decimal import Decimal
from typing import Optional

from pydantic import BaseModel, field_validator

from app.schemas.validators import blank_to_none, validar_data_br

_CENTAVOS = Decimal("0.01")


def _texto_obrigatorio(v: Optional[str], vazio: str, maximo: int, estouro: str) -> str:
    v = blank_to_none(v)
    if v is None:
        raise ValueError(vazio)
    if len(v) > maximo:
        raise ValueError(estouro)
    return v


def _texto_opcional(v: Optional[str], maximo: int, estouro: str) -> Optional[str]:
    v = blank_to_none(v)
    if v is not None and len(v) > maximo:
        raise ValueError(estouro)
    return v


class ManutencaoCreate(BaseModel):
    item_id: int
    defeito: str
    assistencia: Optional[str] = None
    observacoes: Optional[str] = None

    @field_validator("defeito")
    @classmethod
    def _validar_defeito(cls, v):
        return _texto_obrigatorio(v, "Descreva o defeito.", 500, "O defeito deve ter no máximo 500 caracteres.")

    @field_validator("assistencia")
    @classmethod
    def _validar_assistencia(cls, v):
        return _texto_opcional(v, 150, "A assistência técnica deve ter no máximo 150 caracteres.")

    @field_validator("observacoes")
    @classmethod
    def _validar_observacoes(cls, v):
        return _texto_opcional(v, 500, "As observações devem ter no máximo 500 caracteres.")


class ManutencaoEnvio(BaseModel):
    assistencia: str
    previsao_retorno: Optional[str] = None  # dd/mm/aaaa

    @field_validator("assistencia")
    @classmethod
    def _validar_assistencia(cls, v):
        return _texto_obrigatorio(
            v, "Informe a assistência técnica.", 150, "A assistência técnica deve ter no máximo 150 caracteres."
        )

    @field_validator("previsao_retorno")
    @classmethod
    def _validar_previsao(cls, v):
        return validar_data_br(v, campo="previsão de retorno")


class ManutencaoEmReparo(BaseModel):
    previsao_retorno: Optional[str] = None  # dd/mm/aaaa

    @field_validator("previsao_retorno")
    @classmethod
    def _validar_previsao(cls, v):
        return validar_data_br(v, campo="previsão de retorno")


class ManutencaoCancelamento(BaseModel):
    motivo: str

    @field_validator("motivo")
    @classmethod
    def _validar_motivo(cls, v):
        return _texto_obrigatorio(
            v, "Informe o motivo do cancelamento.", 255, "O motivo deve ter no máximo 255 caracteres."
        )


class ManutencaoResponse(BaseModel):
    id: int
    item_id: int
    status: str
    defeito: str
    assistencia: Optional[str] = None
    aberta_por: str
    aberta_em: datetime
    enviada_em: Optional[datetime] = None
    em_reparo_em: Optional[datetime] = None
    previsao_retorno: Optional[date] = None
    retornada_em: Optional[datetime] = None
    cancelada_em: Optional[datetime] = None
    encerrada_por: Optional[str] = None
    custo: Optional[Decimal] = None
    anexo_key: Optional[str] = None
    observacoes: Optional[str] = None
    item_tipo: Optional[str] = None
    item_brand: Optional[str] = None
    item_model: Optional[str] = None
    item_identificador: Optional[str] = None
    item_revenda: Optional[str] = None
    item_status: Optional[str] = None
    custo_acumulado_item: Decimal = Decimal("0.00")

    @field_validator("custo", "custo_acumulado_item")
    @classmethod
    def _centavos(cls, v):
        return None if v is None else Decimal(v).quantize(_CENTAVOS)


class ManutencaoLista(BaseModel):
    items: list[ManutencaoResponse]
    total: int
    custo_total: Decimal

    @field_validator("custo_total")
    @classmethod
    def _centavos(cls, v):
        return Decimal(v).quantize(_CENTAVOS)
