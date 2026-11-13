"""Schemas dos locais físicos de cada unidade."""
from typing import Optional

from pydantic import BaseModel, field_validator

from app.schemas.validators import blank_to_none

MSG_NOME_OBRIGATORIO = "Nome é obrigatório."
MSG_NOME_LONGO = "Nome deve ter no máximo 150 caracteres."
MSG_DESCRICAO_LONGA = "Descrição deve ter no máximo 255 caracteres."
MSG_TEXTO_INVALIDO = "Valor inválido: informe um texto."


def _nome(valor):
    if not isinstance(valor, str):
        raise ValueError(MSG_TEXTO_INVALIDO)
    valor = blank_to_none(valor)
    if valor is None:
        raise ValueError(MSG_NOME_OBRIGATORIO)
    if len(valor) > 150:
        raise ValueError(MSG_NOME_LONGO)
    return valor


def _descricao(valor):
    if valor is not None and not isinstance(valor, str):
        raise ValueError(MSG_TEXTO_INVALIDO)
    valor = blank_to_none(valor)
    if valor is not None and len(valor) > 255:
        raise ValueError(MSG_DESCRICAO_LONGA)
    return valor


class LocalCreate(BaseModel):
    unidade_id: int
    nome: str
    descricao: Optional[str] = None

    @field_validator("nome", mode="before")
    @classmethod
    def _v_nome(cls, v):
        return _nome(v)

    @field_validator("descricao", mode="before")
    @classmethod
    def _v_descricao(cls, v):
        return _descricao(v)


class LocalUpdate(BaseModel):
    """Campo ausente = sem mudança; `descricao: null` limpa a descrição."""

    nome: Optional[str] = None
    descricao: Optional[str] = None
    ativo: Optional[bool] = None

    @field_validator("nome", mode="before")
    @classmethod
    def _v_nome(cls, v):
        # `nome: null` explícito é tratado no router (como campo não enviado).
        return None if v is None else _nome(v)

    @field_validator("descricao", mode="before")
    @classmethod
    def _v_descricao(cls, v):
        return _descricao(v)


class LocalResponse(BaseModel):
    id: int
    unidade_id: int
    unidade_nome: str
    nome: str
    descricao: Optional[str] = None
    ativo: bool
