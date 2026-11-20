"""Schemas da coleta de dados da máquina (coletor PowerShell).

A coleta chega de uma máquina que ainda não foi cadastrada, por rota pública: os limites
abaixo (tamanho de texto e de listas) existem para recusar corpo malicioso com 422.
"""
from typing import Annotated, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field

Texto200 = Annotated[str, Field(max_length=200)]

# Campos do formulário de cadastro preenchidos pela coleta (só os que tiverem valor).
# Chaves: brand, model, identificador, dominio, host, endereco_fisico, cpu, ram, storage,
# sistema, licenca, anydesk.
CamposColeta = dict[str, str]


class Disco(BaseModel):
    tipo: Optional[str] = Field(default=None, max_length=20)
    bytes: Optional[int] = Field(default=None, ge=0)
    barramento: Optional[str] = Field(default=None, max_length=20)


class Adaptador(BaseModel):
    nome: Optional[Texto200] = None
    descricao: Optional[Texto200] = None
    mac: Optional[str] = Field(default=None, max_length=40)
    tipo: Literal["ethernet", "wifi", "outro"] = "outro"
    ativo: bool = False
    gateway: bool = False


class ColetaBruta(BaseModel):
    """Dados crus do Windows, como o coletor os envia."""

    model_config = ConfigDict(extra="ignore")

    versao_coletor: str = Field(max_length=20)
    fabricante: Optional[Texto200] = None
    modelo: Optional[Texto200] = None
    versao_produto: Optional[Texto200] = None
    serial: Optional[Texto200] = None
    host: Optional[Texto200] = None
    dominio: Optional[Texto200] = None
    cpu: Optional[Texto200] = None
    sistema: Optional[Texto200] = None
    product_id: Optional[Texto200] = None
    anydesk: Optional[Texto200] = None
    chassi: list[int] = Field(default_factory=list, max_length=10)
    no_dominio: Optional[bool] = None
    ram_bytes: Optional[int] = None
    discos: list[Disco] = Field(default_factory=list, max_length=10)
    adaptadores: list[Adaptador] = Field(default_factory=list, max_length=20)
