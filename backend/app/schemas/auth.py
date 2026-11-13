from typing import Literal, Optional

from pydantic import BaseModel, Field


class ForgotRequest(BaseModel):
    identificador: str = Field(max_length=200)


class ResetRequest(BaseModel):
    token: str
    nova_senha: str


class MensagemResponse(BaseModel):
    detail: str


class LoginRequest(BaseModel):
    username: str
    password: str


class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"


class LoginResponse(BaseModel):
    """Resposta do login em etapas: `ok` (sessão aberta), `mfa` (falta o código)
    ou `cadastro_2fa` (Gestor sem autenticador: precisa cadastrar antes de entrar)."""

    etapa: Literal["ok", "mfa", "cadastro_2fa"]
    access_token: Optional[str] = None
    token_type: str = "bearer"
    mfa_token: Optional[str] = None
    expira_em_segundos: Optional[int] = None
    codigos_recuperacao_restantes: Optional[int] = None
    codigos_recuperacao: Optional[list[str]] = None


class MfaVerificarRequest(BaseModel):
    mfa_token: str
    codigo: str


class UserMe(BaseModel):
    id: int
    username: str
    role: str


class MfaTokenRequest(BaseModel):
    mfa_token: str


class MfaConfirmarRequest(BaseModel):
    mfa_token: str
    codigo: str


class CadastroTotpResponse(BaseModel):
    segredo: str
    otpauth_uri: str
