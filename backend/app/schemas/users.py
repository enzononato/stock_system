from typing import Literal, Optional

from pydantic import BaseModel, field_validator

from app.schemas.validators import validar_email


class UserCreate(BaseModel):
    username: str
    password: str
    role: Literal["Gestor", "Técnico", "Jovem Aprendiz"]
    email: Optional[str] = None

    @field_validator("email")
    @classmethod
    def _email(cls, v: Optional[str]) -> Optional[str]:
        # vazio/só espaços vira None; senão minúsculas e formato validado
        return validar_email(v)


class UserUpdate(BaseModel):
    # Chave obrigatória (aceita null, que limpa o e-mail).
    email: Optional[str]

    @field_validator("email")
    @classmethod
    def _email(cls, v: Optional[str]) -> Optional[str]:
        return validar_email(v)


class UserResponse(BaseModel):
    id: int
    username: str
    role: str
    email: Optional[str] = None
    totp_ativo: bool = False


class PasswordUpdate(BaseModel):
    new_password: str
