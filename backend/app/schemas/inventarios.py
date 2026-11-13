"""Schemas do inventário físico e das etiquetas."""
from datetime import datetime
from typing import Literal, Optional

from pydantic import BaseModel, field_validator

from app.schemas.validators import blank_to_none

MSG_TEXTO_INVALIDO = "Valor inválido: informe um texto."
MSG_OBSERVACAO_LONGA = "Observação muito longa (máximo 500 caracteres)."


def _texto_opcional(valor, limite: int, mensagem_longo: str):
    if valor is not None and not isinstance(valor, str):
        raise ValueError(MSG_TEXTO_INVALIDO)
    valor = blank_to_none(valor)
    if valor is not None and len(valor) > limite:
        raise ValueError(mensagem_longo)
    return valor


class InventarioCreate(BaseModel):
    unidade_id: int
    observacao: Optional[str] = None

    @field_validator("observacao", mode="before")
    @classmethod
    def _v_observacao(cls, v):
        return _texto_opcional(v, 500, MSG_OBSERVACAO_LONGA)


class CancelarInventario(BaseModel):
    motivo: str

    @field_validator("motivo", mode="before")
    @classmethod
    def _v_motivo(cls, v):
        if not isinstance(v, str):
            raise ValueError(MSG_TEXTO_INVALIDO)
        v = v.strip()
        if len(v) < 3:
            raise ValueError("Informe o motivo do cancelamento.")
        if len(v) > 255:
            raise ValueError("Motivo muito longo (máximo 255 caracteres).")
        return v


class ResumoInventario(BaseModel):
    esperados: int
    lidos: int
    encontrados: int
    ok: int
    local_errado: int
    usuario_diferente: int
    outra_unidade: int
    nao_cadastrado: int
    nao_encontrados: int


class InventarioResponse(BaseModel):
    id: int
    unidade_id: int
    unidade_nome: str
    status: Literal["aberto", "encerrado", "cancelado"]
    iniciado_por: str
    iniciado_em: datetime
    encerrado_por: Optional[str] = None
    encerrado_em: Optional[datetime] = None
    observacao: Optional[str] = None
    motivo_cancelamento: Optional[str] = None
    total_leituras: int = 0


class InventarioDetalhe(InventarioResponse):
    resumo: ResumoInventario


# ── Leituras ────────────────────────────────────────────────────────────────

Resultado = Literal["ok", "local_errado", "usuario_diferente", "outra_unidade", "nao_cadastrado"]


class LeituraCreate(BaseModel):
    codigo: str
    local_id_encontrado: Optional[int] = None
    colaborador_id_encontrado: Optional[int] = None
    usuario_conferido: bool = False
    observacao: Optional[str] = None

    @field_validator("codigo", mode="before")
    @classmethod
    def _v_codigo(cls, v):
        if not isinstance(v, str):
            raise ValueError(MSG_TEXTO_INVALIDO)
        v = v.strip()
        if not v:
            raise ValueError("Informe o código lido.")
        if len(v) > 255:
            raise ValueError("Código muito longo (máximo 255 caracteres).")
        return v

    @field_validator("observacao", mode="before")
    @classmethod
    def _v_observacao(cls, v):
        return _texto_opcional(v, 500, MSG_OBSERVACAO_LONGA)


class LeituraUpdate(BaseModel):
    """Campo ausente = sem mudança; `null` limpa local/colaborador/observação."""

    local_id_encontrado: Optional[int] = None
    colaborador_id_encontrado: Optional[int] = None
    usuario_conferido: Optional[bool] = None
    observacao: Optional[str] = None

    @field_validator("observacao", mode="before")
    @classmethod
    def _v_observacao(cls, v):
        return _texto_opcional(v, 500, MSG_OBSERVACAO_LONGA)


class LeituraResponse(BaseModel):
    id: int
    inventario_id: int
    item_id: Optional[int] = None
    codigo_lido: str
    lido_por: str
    lido_em: datetime
    resultado: Resultado
    detalhe: Optional[str] = None
    observacao: Optional[str] = None
    atualizado_em: Optional[datetime] = None
    local_id_encontrado: Optional[int] = None
    local_encontrado_nome: Optional[str] = None
    usuario_conferido: bool = False
    colaborador_id_encontrado: Optional[int] = None
    colaborador_encontrado_nome: Optional[str] = None
    local_id_esperado: Optional[int] = None
    local_esperado_nome: Optional[str] = None
    usuario_esperado: Optional[str] = None
    item_tipo: Optional[str] = None
    item_brand: Optional[str] = None
    item_model: Optional[str] = None
    item_identificador: Optional[str] = None
    item_revenda: Optional[str] = None
    duplicada: bool = False


# ── Relatório ───────────────────────────────────────────────────────────────

class RelatorioLinha(BaseModel):
    categoria: Literal[
        "ok", "local_errado", "usuario_diferente", "outra_unidade", "nao_cadastrado", "nao_encontrado"
    ]
    leitura_id: Optional[int] = None
    item_id: Optional[int] = None
    tipo: Optional[str] = None
    brand: Optional[str] = None
    model: Optional[str] = None
    identificador: Optional[str] = None
    status_item: Optional[str] = None
    unidade_item: Optional[str] = None
    local_esperado: Optional[str] = None
    local_encontrado: Optional[str] = None
    usuario_esperado: Optional[str] = None
    usuario_encontrado: Optional[str] = None
    codigo_lido: Optional[str] = None
    lido_por: Optional[str] = None
    lido_em: Optional[datetime] = None
    detalhe: Optional[str] = None
    observacao: Optional[str] = None


class RelatorioInventario(BaseModel):
    inventario: InventarioDetalhe
    linhas: list[RelatorioLinha]


# ── Etiquetas ───────────────────────────────────────────────────────────────

class EtiquetasRequest(BaseModel):
    item_ids: list[int]

    @field_validator("item_ids")
    @classmethod
    def _v_item_ids(cls, v):
        v = list(dict.fromkeys(v))  # sem duplicados, na ordem recebida
        if not v:
            raise ValueError("Selecione ao menos um equipamento.")
        if len(v) > 2000:
            raise ValueError("Selecione no máximo 2000 equipamentos por vez.")
        return v


class Etiqueta(BaseModel):
    id: int
    tipo: Optional[str] = None
    brand: Optional[str] = None
    model: Optional[str] = None
    identificador: Optional[str] = None
    revenda: Optional[str] = None
    url: str


class EtiquetasResponse(BaseModel):
    base_url: str
    etiquetas: list[Etiqueta]
    ignorados: list[int]


class IdsEtiquetas(BaseModel):
    ids: list[int]
    total: int
