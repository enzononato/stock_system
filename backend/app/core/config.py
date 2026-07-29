from pydantic_settings import BaseSettings
from pydantic import field_validator
from typing import Optional
import os


class Settings(BaseSettings):
    # Ambiente de execução: "development" (padrão) ou "production".
    # Em produção, cookie_secure_effective força cookies Secure=True automaticamente,
    # mesmo que COOKIE_SECURE esteja configurado incorretamente no .env.
    ENVIRONMENT: str = "development"

    # Banco de dados MySQL — obrigatórios, SEM valor padrão. Nunca commitar
    # credenciais reais no código-fonte: o app deve falhar no boot com uma
    # mensagem clara caso alguma destas variáveis esteja ausente do .env.
    DB_HOST: str
    DB_PORT: int = 3306
    DB_USER: str
    DB_PASSWORD: str
    DB_NAME: str
    DB_CHARSET: str = "utf8mb4"

    # Pool de conexões MySQL (DBUtils) usado pela camada de acesso a dados
    DB_POOL_SIZE: int = 5
    DB_POOL_MAX_CONNECTIONS: int = 20

    # JWT — obrigatório, SEM valor padrão (ver validador abaixo: mínimo 32 caracteres).
    # Gere um valor seguro com: python -c "import secrets; print(secrets.token_urlsafe(48))"
    JWT_SECRET: str
    JWT_ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 15
    REFRESH_TOKEN_EXPIRE_DAYS: int = 7

    # Cookies do refresh token
    COOKIE_SECURE: bool = False
    COOKIE_SAMESITE: str = "lax"

    # Rate limiting do endpoint de login (proteção contra força bruta)
    LOGIN_RATE_LIMIT: str = "5/minute"

    # Paginação padrão das listagens da API
    DEFAULT_PAGE_SIZE: int = 50
    MAX_PAGE_SIZE: int = 500

    # Storage
    STORAGE_BACKEND: str = "local"  # "local" ou "s3"
    LOCAL_STORAGE_PATH: str = "storage"
    S3_BUCKET: Optional[str] = None
    S3_REGION: Optional[str] = "us-east-1"
    AWS_ACCESS_KEY_ID: Optional[str] = None
    AWS_SECRET_ACCESS_KEY: Optional[str] = None

    # Modelos de documentos (.docx): diretório absoluto opcional. Quando vazio,
    # o caminho é calculado a partir de __file__ (ver _modelo_path mais abaixo).
    MODELOS_DIR: str = ""

    # CORS
    CORS_ORIGINS: str = "http://localhost:5173,http://localhost:3000"

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

    @field_validator("JWT_SECRET")
    @classmethod
    def validar_tamanho_jwt_secret(cls, v: str) -> str:
        if len(v) < 32:
            raise ValueError(
                "JWT_SECRET deve ter no mínimo 32 caracteres. Gere um valor seguro com: "
                'python -c "import secrets; print(secrets.token_urlsafe(48))"'
            )
        return v

    @property
    def cors_origins_list(self) -> list[str]:
        return [o.strip() for o in self.CORS_ORIGINS.split(",")]

    @property
    def cookie_secure_effective(self) -> bool:
        """True sempre que ENVIRONMENT == "production", independentemente do
        valor bruto de COOKIE_SECURE — evita servir cookies de refresh token
        sem o atributo Secure em produção por um .env mal configurado."""
        if self.ENVIRONMENT == "production":
            return True
        return self.COOKIE_SECURE


settings = Settings()

# Constantes de domínio (reaproveitadas do config.py original)
CENTER_COST_OPTIONS = [
    "101 - Puxada",
    "202 - Armazém",
    "301 - Administrativo",
    "401 - Vendas",
    "501 - Entrega",
    "601 - CSC",
]

REVENDAS_OPTIONS = [
    "Revalle Juazeiro",
    "Revalle Bonfim",
    "Revalle Petrolina",
    "Revalle Ribeira",
    "Revalle Paulo Afonso",
    "Revalle Alagoinhas",
    "Revalle Serrinha",
]

SETORES_OPTIONS = [
    "Contabilidade",
    "Financeiro",
    "Departamento pessoal",
    "Cultura",
    "Gente",
    "Armazém",
    "Distribuição",
    "Segurança",
    "Compras",
    "TI",
    "CME",
    "Vendas",
    "Puxada",
    "Faturamento",
    "Portaria",
    "Auditório",
    "Caixa",
    "Diretoria",
]

EQUIPMENT_TYPES = [
    "Celular",
    "Notebook",
    "Desktop",
    "Impressora",
    "Tablet",
    "Switch",
    "HD",
    "Nobreak",
    "Access Point",
]

PERIPHERAL_TYPES = [
    "Mouse",
    "Teclado",
    "Monitor",
    "Mousepad",
    "Suporte",
    "Estabilizador",
    "Mochila",
    "Fone",
    "Webcam",
    "Adaptador",
    "Outro",
]

REMOVAL_REASONS = {
    "Roubo": True,
    "Perda": False,
    "Obsolescência": False,
    "Doação": True,
    "Venda": True,
}

# Caminhos dos modelos de termo. Por padrão resolve para backend/modelos (3 níveis
# acima de app/core/config.py — em contêiner, com WORKDIR /app, resolve para
# /app/modelos, que é exatamente o volume somente-leitura montado pelo
# docker-compose). Se MODELOS_DIR estiver preenchido no .env, ele tem prioridade
# sobre o cálculo automático baseado em __file__.
def _modelo_path(filename: str) -> str:
    if settings.MODELOS_DIR:
        base = settings.MODELOS_DIR
    else:
        base = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "modelos")
    return os.path.join(base, filename)

TERMO_MODELOS = {
    "Revalle Juazeiro":     _modelo_path("termo_juazeiro.docx"),
    "Revalle Bonfim":       _modelo_path("termo_bonfim.docx"),
    "Revalle Petrolina":    _modelo_path("termo_petrolina.docx"),
    "Revalle Ribeira":      _modelo_path("termo_ribeira.docx"),
    "Revalle Paulo Afonso": _modelo_path("termo_pauloafonso.docx"),
    "Revalle Alagoinhas":   _modelo_path("termo_alagoinhas.docx"),
    "Revalle Serrinha":     _modelo_path("termo_serrinha.docx"),
}

TERMO_DEVOLUCAO_MODELOS = {
    "Revalle Juazeiro":     _modelo_path("devolucao/termo_devolucao_juazeiro.docx"),
    "Revalle Bonfim":       _modelo_path("devolucao/termo_devolucao_bonfim.docx"),
    "Revalle Petrolina":    _modelo_path("devolucao/termo_devolucao_petrolina.docx"),
    "Revalle Ribeira":      _modelo_path("devolucao/termo_devolucao_ribeira.docx"),
    "Revalle Paulo Afonso": _modelo_path("devolucao/termo_devolucao_pauloafonso.docx"),
    "Revalle Alagoinhas":   _modelo_path("devolucao/termo_devolucao_alagoinhas.docx"),
    "Revalle Serrinha":     _modelo_path("devolucao/termo_devolucao_serrinha.docx"),
}
