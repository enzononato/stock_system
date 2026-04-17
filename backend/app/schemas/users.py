from pydantic import BaseModel
from typing import Literal


class UserCreate(BaseModel):
    username: str
    password: str
    role: Literal["Gestor", "Técnico", "Jovem Aprendiz"]


class UserResponse(BaseModel):
    id: int
    username: str
    role: str


class PasswordUpdate(BaseModel):
    new_password: str
