"""Schemas das linhas telefônicas (F5): número validado, custo em Decimal e eventos."""
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Optional

from pydantic import BaseModel, field_validator

from app.schemas.validators import blank_to_none

OPERADORAS: tuple[str, ...] = ("Vivo", "Claro", "TIM", "Algar", "Outra")
STATUS_LINHA: tuple[str, ...] = ("ativa", "suspensa", "cancelada")
TIPOS_APARELHO: tuple[str, ...] = ("Celular", "Tablet")

DDDS_VALIDOS: frozenset[str] = frozenset(
    [str(n) for n in range(11, 20)]
    + ["21", "22", "24", "27", "28"]
    + [str(n) for n in range(31, 36)]
    + ["37", "38"]
    + [str(n) for n in range(41, 50)]
    + ["51", "53", "54", "55"]
    + [str(n) for n in range(61, 70)]
    + ["71", "73", "74", "75", "77", "79"]
    + [str(n) for n in range(81, 90)]
    + [str(n) for n in range(91, 100)]
)

MSG_NUMERO_TAMANHO = "informe DDD + 8 ou 9 dígitos."
MSG_VAZIO = "não pode ficar vazio."
_LIMITE_CUSTO = Decimal("100000000")


def normalizar_numero_linha(valor: Optional[str]) -> str:
    """Só dígitos (DDD + 8 ou 9 dígitos); `55` de país na frente é removido."""
    digitos = re.sub(r"\D", "", valor or "")
    if len(digitos) in (12, 13) and digitos.startswith("55"):
        digitos = digitos[2:]
    if len(digitos) not in (10, 11):
        raise ValueError(MSG_NUMERO_TAMANHO)
    if digitos[:2] not in DDDS_VALIDOS:
        raise ValueError("DDD inexistente.")
    if len(digitos) == 11 and digitos[2] != "9":
        raise ValueError("celular com 9 dígitos deve começar com 9.")
    return digitos


def formatar_numero_linha(numero: Optional[str]) -> str:
    if not numero:
        return ""
    if len(numero) == 11 and numero.isdigit():
        return f"({numero[:2]}) {numero[2:7]}-{numero[7:]}"
    if len(numero) == 10 and numero.isdigit():
        return f"({numero[:2]}) {numero[2:6]}-{numero[6:]}"
    return numero


def normalizar_iccid(valor: Optional[str]) -> Optional[str]:
    valor = blank_to_none(valor)
    if valor is None:
        return None
    valor = re.sub(r"\s+", "", valor)
    if not valor.isdigit() or not 18 <= len(valor) <= 22:
        raise ValueError("informe de 18 a 22 dígitos.")
    return valor


def validar_custo(valor: object) -> Decimal:
    """Custo mensal em Decimal com 2 casas. Não aceita vírgula (o frontend converte)."""
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        raise ValueError("informe o valor.")
    try:
        numero = Decimal(str(valor).strip())
    except InvalidOperation:
        raise ValueError("valor inválido.")
    if not numero.is_finite():
        raise ValueError("valor inválido.")
    if numero < 0:
        raise ValueError("não pode ser negativo.")
    if numero != numero.quantize(Decimal("0.01")):
        raise ValueError("use no máximo 2 casas decimais.")
    if numero >= _LIMITE_CUSTO:
        raise ValueError("valor acima do permitido.")
    return numero.quantize(Decimal("0.01"))


def validar_operadora(valor: Optional[str]) -> str:
    if isinstance(valor, str):
        for operadora in OPERADORAS:
            if operadora.lower() == valor.strip().lower():
                return operadora
    raise ValueError("escolha Vivo, Claro, TIM, Algar ou Outra.")


def _texto(valor: Optional[str], maximo: int) -> Optional[str]:
    if valor is not None and not isinstance(valor, str):
        raise ValueError("informe um texto.")
    valor = blank_to_none(valor)
    if valor is not None and len(valor) > maximo:
        raise ValueError(f"use no máximo {maximo} caracteres.")
    return valor


def _numero(valor):
    return normalizar_numero_linha(valor if isinstance(valor, str) else None)


class LinhaCreate(BaseModel):
    numero: str
    operadora: str
    plano: Optional[str] = None
    custo_mensal: Decimal
    unidade_id: int
    item_id: Optional[int] = None
    colaborador_id: Optional[int] = None
    iccid: Optional[str] = None
    observacoes: Optional[str] = None

    @field_validator("numero", mode="before")
    @classmethod
    def _v_numero(cls, v):
        return _numero(v)

    @field_validator("operadora", mode="before")
    @classmethod
    def _v_operadora(cls, v):
        return validar_operadora(v)

    @field_validator("custo_mensal", mode="before")
    @classmethod
    def _v_custo(cls, v):
        return validar_custo(v)

    @field_validator("iccid", mode="before")
    @classmethod
    def _v_iccid(cls, v):
        return normalizar_iccid(v)

    @field_validator("plano", mode="before")
    @classmethod
    def _v_plano(cls, v):
        return _texto(v, 150)

    @field_validator("observacoes", mode="before")
    @classmethod
    def _v_observacoes(cls, v):
        return _texto(v, 500)


class LinhaUpdate(BaseModel):
    """Chave ausente = não muda; `null` limpa plano, item_id, colaborador_id, iccid e observacoes.
    `null` em numero, operadora, custo_mensal ou unidade_id é recusado."""

    numero: Optional[str] = None
    operadora: Optional[str] = None
    plano: Optional[str] = None
    custo_mensal: Optional[Decimal] = None
    unidade_id: Optional[int] = None
    item_id: Optional[int] = None
    colaborador_id: Optional[int] = None
    iccid: Optional[str] = None
    observacoes: Optional[str] = None
    motivo: Optional[str] = None

    @field_validator("numero", mode="before")
    @classmethod
    def _v_numero(cls, v):
        if v is None:
            raise ValueError(MSG_VAZIO)
        return _numero(v)

    @field_validator("operadora", mode="before")
    @classmethod
    def _v_operadora(cls, v):
        if v is None:
            raise ValueError(MSG_VAZIO)
        return validar_operadora(v)

    @field_validator("custo_mensal", mode="before")
    @classmethod
    def _v_custo(cls, v):
        if v is None:
            raise ValueError(MSG_VAZIO)
        return validar_custo(v)

    @field_validator("unidade_id", mode="before")
    @classmethod
    def _v_unidade(cls, v):
        if v is None:
            raise ValueError(MSG_VAZIO)
        return v

    @field_validator("iccid", mode="before")
    @classmethod
    def _v_iccid(cls, v):
        return normalizar_iccid(v)

    @field_validator("plano", mode="before")
    @classmethod
    def _v_plano(cls, v):
        return _texto(v, 150)

    @field_validator("observacoes", mode="before")
    @classmethod
    def _v_observacoes(cls, v):
        return _texto(v, 500)

    @field_validator("motivo", mode="before")
    @classmethod
    def _v_motivo(cls, v):
        return _texto(v, 255)


class LinhaStatusChange(BaseModel):
    status: str
    observacao: Optional[str] = None

    @field_validator("status", mode="before")
    @classmethod
    def _v_status(cls, v):
        if v not in STATUS_LINHA:
            raise ValueError("escolha ativa, suspensa ou cancelada.")
        return v

    @field_validator("observacao", mode="before")
    @classmethod
    def _v_observacao(cls, v):
        return _texto(v, 255)


class LinhaResponse(BaseModel):
    id: int
    numero: str
    operadora: str
    plano: Optional[str] = None
    custo_mensal: Decimal
    status: str
    unidade_id: int
    unidade_nome: str
    item_id: Optional[int] = None
    item_tipo: Optional[str] = None
    item_brand: Optional[str] = None
    item_model: Optional[str] = None
    item_identificador: Optional[str] = None
    item_ativo: Optional[bool] = None
    colaborador_id: Optional[int] = None
    colaborador_nome: Optional[str] = None
    colaborador_situacao: Optional[str] = None
    iccid: Optional[str] = None
    observacoes: Optional[str] = None
    criada_em: datetime
    atualizada_em: datetime


class LinhaEventoResponse(BaseModel):
    id: int
    linha_id: int
    linha_numero: str
    tipo: str
    campo: Optional[str] = None
    de_valor: Optional[str] = None
    para_valor: Optional[str] = None
    item_id: Optional[int] = None
    item_id_anterior: Optional[int] = None
    ocorrido_em: datetime
    por_usuario: Optional[str] = None
    observacao: Optional[str] = None


class AparelhoOpcao(BaseModel):
    id: int
    tipo: str
    brand: Optional[str] = None
    model: Optional[str] = None
    identificador: Optional[str] = None
    codigo_patrimonial: Optional[str] = None
    revenda: Optional[str] = None
    status: Optional[str] = None
    responsavel_nome: Optional[str] = None


class CustoUnidade(BaseModel):
    unidade_id: int
    unidade_nome: str
    quantidade: int
    custo_total: Decimal


class RelatorioCusto(BaseModel):
    unidades: list[CustoUnidade]
    total_linhas: int
    custo_total: Decimal
    gerado_em: datetime
