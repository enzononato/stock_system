"""Schemas do cadastro de colaboradores (sincronizado com a base de RH)."""
from datetime import datetime
from typing import Literal, Optional

from pydantic import BaseModel, field_validator

from app.schemas.validators import (
    blank_to_none,
    normalizar_telefone,
    validar_cpf,
    validar_email,
)

Situacao = Literal["ativo", "afastado", "desligado"]


def _nome_obrigatorio(nome: Optional[str]) -> str:
    nome = blank_to_none(nome)
    if nome is None:
        raise ValueError("Nome é obrigatório.")
    return nome


class ColaboradorResponse(BaseModel):
    id: int
    cpf: str
    nome: str
    matricula: Optional[str] = None
    cargo: Optional[str] = None
    setor: Optional[str] = None
    unidade_rh: Optional[str] = None
    unidade_id: Optional[int] = None
    unidade_nome: Optional[str] = None
    situacao: Situacao
    email: Optional[str] = None
    telefone: Optional[str] = None
    email_manual: bool
    telefone_manual: bool
    superior_id: Optional[int] = None
    superior_nome: Optional[str] = None
    origem: Literal["rh", "manual"]
    sincronizado_em: Optional[datetime] = None
    criado_em: datetime
    atualizado_em: datetime


class ColaboradorCpfResponse(ColaboradorResponse):
    """Resposta da consulta por CPF: o cadastro local ou, na falta dele, o que a base
    de RH tem (nada é gravado; `local=False` e `id=None`)."""

    id: Optional[int] = None
    email_manual: bool = False
    telefone_manual: bool = False
    criado_em: Optional[datetime] = None
    atualizado_em: Optional[datetime] = None
    local: bool


class ColaboradorCreate(BaseModel):
    cpf: str
    nome: str
    cargo: Optional[str] = None
    setor: Optional[str] = None
    unidade_id: Optional[int] = None
    email: Optional[str] = None
    telefone: Optional[str] = None
    superior_id: Optional[int] = None
    situacao: Situacao = "ativo"

    @field_validator("cpf")
    @classmethod
    def _v_cpf(cls, v):
        return validar_cpf(v)

    @field_validator("nome")
    @classmethod
    def _v_nome(cls, v):
        return _nome_obrigatorio(v)

    @field_validator("cargo", "setor")
    @classmethod
    def _v_opcionais(cls, v):
        return blank_to_none(v)

    @field_validator("email")
    @classmethod
    def _v_email(cls, v):
        return validar_email(v)

    @field_validator("telefone")
    @classmethod
    def _v_telefone(cls, v):
        return normalizar_telefone(v)


class ColaboradorUpdate(BaseModel):
    """Todos opcionais. Use `model_dump(exclude_unset=True)`: chave ausente = sem mudança,
    chave `null` (ou "") = limpa o campo (e-mail, telefone, superior, cargo, setor, unidade)."""

    nome: Optional[str] = None
    cargo: Optional[str] = None
    setor: Optional[str] = None
    unidade_id: Optional[int] = None
    situacao: Optional[Situacao] = None
    email: Optional[str] = None
    telefone: Optional[str] = None
    superior_id: Optional[int] = None

    @field_validator("nome")
    @classmethod
    def _v_nome(cls, v):
        # O validador só roda quando a chave foi enviada: aqui `None` é "limpar", o que não vale.
        return _nome_obrigatorio(v)

    @field_validator("situacao")
    @classmethod
    def _v_situacao(cls, v):
        if v is None:
            raise ValueError("Situação inválida.")
        return v

    @field_validator("cargo", "setor")
    @classmethod
    def _v_opcionais(cls, v):
        return blank_to_none(v)

    @field_validator("email")
    @classmethod
    def _v_email(cls, v):
        return validar_email(v)

    @field_validator("telefone")
    @classmethod
    def _v_telefone(cls, v):
        return normalizar_telefone(v)


class ResumoSyncResponse(BaseModel):
    total_base: int
    criados: int
    atualizados: int
    desligados: int
    ignorados: int = 0
    duplicados: int = 0


class SyncStatus(BaseModel):
    ultima: Optional[dict] = None
    configurado: bool
