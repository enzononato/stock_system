from pydantic_settings import BaseSettings
from typing import Optional
import os


class Settings(BaseSettings):
    # Database
    DB_HOST: str
    DB_USER: str = "root"
    DB_PASSWORD: str
    DB_NAME: str = "stock_sys_db"
    DB_CHARSET: str = "utf8mb4"

    # JWT
    JWT_SECRET: str = "CHANGE_ME_IN_PRODUCTION_USE_A_LONG_RANDOM_STRING"
    JWT_ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 15
    REFRESH_TOKEN_EXPIRE_DAYS: int = 7

    # Storage
    STORAGE_BACKEND: str = "local"  # "local" ou "s3"
    LOCAL_STORAGE_PATH: str = "storage"
    S3_BUCKET: Optional[str] = None
    S3_REGION: Optional[str] = "us-east-1"
    AWS_ACCESS_KEY_ID: Optional[str] = None
    AWS_SECRET_ACCESS_KEY: Optional[str] = None

    # CORS
    CORS_ORIGINS: str = "http://localhost:5173,http://localhost:3000"

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

    @property
    def cors_origins_list(self) -> list[str]:
        return [o.strip() for o in self.CORS_ORIGINS.split(",")]


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

# Caminhos dos modelos de termo (relativos ao diretório backend/)
def _modelo_path(filename: str) -> str:
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
