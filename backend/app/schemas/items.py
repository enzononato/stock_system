from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class ItemCreate(BaseModel):
    tipo: str
    brand: str
    model: str
    revenda: str
    date_registered: str  # dd/mm/yyyy
    # Campos opcionais (variam por tipo)
    identificador: Optional[str] = None
    nota_fiscal: Optional[str] = None
    fornecedor: Optional[str] = None
    dominio: Optional[str] = None
    host: Optional[str] = None
    endereco_fisico: Optional[str] = None
    cpu: Optional[str] = None
    ram: Optional[str] = None
    storage: Optional[str] = None
    sistema: Optional[str] = None
    licenca: Optional[str] = None
    anydesk: Optional[str] = None
    setor: Optional[str] = None
    ip: Optional[str] = None
    mac: Optional[str] = None
    fornecedor: Optional[str] = None
    potencia_nominal: Optional[str] = None
    autonomia_estimada: Optional[str] = None
    ip_snmp: Optional[str] = None
    codigo_patrimonial: Optional[str] = None
    responsavel: Optional[str] = None
    local_instalacao: Optional[str] = None
    poe: Optional[str] = None
    quantidade_portas: Optional[str] = None


class ItemUpdate(BaseModel):
    tipo: Optional[str] = None
    brand: Optional[str] = None
    model: Optional[str] = None
    revenda: Optional[str] = None
    identificador: Optional[str] = None
    nota_fiscal: Optional[str] = None
    fornecedor: Optional[str] = None
    dominio: Optional[str] = None
    host: Optional[str] = None
    endereco_fisico: Optional[str] = None
    cpu: Optional[str] = None
    ram: Optional[str] = None
    storage: Optional[str] = None
    sistema: Optional[str] = None
    licenca: Optional[str] = None
    anydesk: Optional[str] = None
    setor: Optional[str] = None
    ip: Optional[str] = None
    mac: Optional[str] = None
    potencia_nominal: Optional[str] = None
    autonomia_estimada: Optional[str] = None
    ip_snmp: Optional[str] = None
    codigo_patrimonial: Optional[str] = None
    responsavel: Optional[str] = None
    local_instalacao: Optional[str] = None
    poe: Optional[str] = None
    quantidade_portas: Optional[str] = None


class ItemResponse(BaseModel):
    id: int
    tipo: Optional[str]
    brand: Optional[str]
    model: Optional[str]
    identificador: Optional[str]
    nota_fiscal: Optional[str]
    status: Optional[str]
    assigned_to: Optional[str]
    cpf: Optional[str]
    revenda: Optional[str]
    dominio: Optional[str]
    host: Optional[str]
    endereco_fisico: Optional[str]
    cpu: Optional[str]
    ram: Optional[str]
    storage: Optional[str]
    sistema: Optional[str]
    licenca: Optional[str]
    anydesk: Optional[str]
    setor: Optional[str]
    ip: Optional[str]
    mac: Optional[str]
    fornecedor: Optional[str]
    potencia_nominal: Optional[str]
    autonomia_estimada: Optional[str]
    ip_snmp: Optional[str]
    codigo_patrimonial: Optional[str]
    responsavel: Optional[str]
    local_instalacao: Optional[str]
    poe: Optional[str]
    quantidade_portas: Optional[str]
    date_registered: Optional[datetime]
    date_issued: Optional[datetime]
    peripheral_count: Optional[int] = 0

    class Config:
        from_attributes = True
