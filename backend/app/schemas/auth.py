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


class UserMe(BaseModel):
    id: int
    username: str
    role: str
