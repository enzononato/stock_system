from typing import Literal, Optional

from pydantic import BaseModel, field_validator

from app.schemas.validators import blank_to_none, validar_email_corporativo


class UserCreate(BaseModel):
    """Conta nova: só o e-mail corporativo e o papel. Não há senha — a pessoa recebe
    um convite e cria a própria; o `username` (só para o histórico) é gerado do e-mail."""

    email: str
    role: Literal["Gestor", "Técnico", "Jovem Aprendiz"]

    @field_validator("email")
    @classmethod
    def _email(cls, v: str) -> str:
        return validar_email_corporativo(v)


class UserUpdate(BaseModel):
    # Chave obrigatória (aceita null, que limpa o e-mail).
    email: Optional[str]

    @field_validator("email")
    @classmethod
    def _email(cls, v: Optional[str]) -> Optional[str]:
        v = blank_to_none(v)
        return None if v is None else validar_email_corporativo(v)


class UserResponse(BaseModel):
    id: int
    username: str
    role: str
    email: Optional[str] = None
    email_confirmado: bool = False
    senha_definida: bool = True
    totp_ativo: bool = False
