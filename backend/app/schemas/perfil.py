from typing import Optional

from pydantic import BaseModel


class PerfilResponse(BaseModel):
    id: int
    username: str
    role: str
    email: Optional[str]
    totp_ativo: bool
    mfa_obrigatorio: bool
    mfa_disponivel: bool
    codigos_recuperacao_restantes: int


class CodigoRequest(BaseModel):
    codigo: str


class SenhaRequest(BaseModel):
    senha: str


class CodigosRecuperacaoResponse(BaseModel):
    codigos_recuperacao: list[str]
