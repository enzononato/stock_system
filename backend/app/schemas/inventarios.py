"""Schemas do inventário físico e das etiquetas."""
from datetime import datetime
from typing import Literal, Optional

from pydantic import BaseModel, field_validator

from app.schemas.validators import blank_to_none

MSG_TEXTO_INVALIDO = "Valor inválido: informe um texto."
MSG_OBSERVACAO_LONGA = "Observação muito longa (máximo 500 caracteres)."


def _texto_opcional(valor, limite: int, mensagem_longo: str):
    if valor is not None and not isinstance(valor, str):
        raise ValueError(MSG_TEXTO_INVALIDO)
    valor = blank_to_none(valor)
    if valor is not None and len(valor) > limite:
        raise ValueError(mensagem_longo)
    return valor


class InventarioCreate(BaseModel):
    unidade_id: int
    observacao: Optional[str] = None

    @field_validator("observacao", mode="before")
    @classmethod
    def _v_observacao(cls, v):
        return _texto_opcional(v, 500, MSG_OBSERVACAO_LONGA)


class CancelarInventario(BaseModel):
    motivo: str

    @field_validator("motivo", mode="before")
    @classmethod
    def _v_motivo(cls, v):
        if not isinstance(v, str):
            raise ValueError(MSG_TEXTO_INVALIDO)
        v = v.strip()
        if len(v) < 3:
            raise ValueError("Informe o motivo do cancelamento.")
        if len(v) > 255:
            raise ValueError("Motivo muito longo (máximo 255 caracteres).")
        return v


class ResumoInventario(BaseModel):
    esperados: int
    lidos: int
    encontrados: int
    ok: int
    local_errado: int
    usuario_diferente: int
    outra_unidade: int
    nao_cadastrado: int
    nao_encontrados: int


class InventarioResponse(BaseModel):
    id: int
    unidade_id: int
    unidade_nome: str
    status: Literal["aberto", "encerrado", "cancelado"]
    iniciado_por: str
    iniciado_em: datetime
    encerrado_por: Optional[str] = None
    encerrado_em: Optional[datetime] = None
    observacao: Optional[str] = None
    motivo_cancelamento: Optional[str] = None
    total_leituras: int = 0


class InventarioDetalhe(InventarioResponse):
    resumo: ResumoInventario
