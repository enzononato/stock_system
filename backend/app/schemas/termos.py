from datetime import datetime
from typing import Literal, Optional

from pydantic import BaseModel, Field, field_validator

from app.core.tempo import FUSO_BRASILIA
from app.schemas.common import Paginated

TipoTermo = Literal["entrega", "devolucao", "uso"]
StatusTermo = Literal["agendado", "enviado", "assinado", "recusado", "expirado", "cancelado", "erro", "manual"]


class TermoResponse(BaseModel):
    # None só no termo "legado" (empréstimo confirmado antes da F3, sem solicitação gravada).
    id: Optional[int] = None
    item_id: int
    item_descricao: Optional[str] = None
    colaborador_id: Optional[int] = None
    colaborador_nome: Optional[str] = None
    tipo: TipoTermo
    history_id: Optional[int] = None
    modelo_versao: Optional[str] = None
    status: StatusTermo
    destinatario_email: Optional[str] = None
    solicitado_por: str
    solicitado_em: datetime
    agendado_para: Optional[datetime] = None
    enviado_em: Optional[datetime] = None
    assinado_em: Optional[datetime] = None
    cancelado_em: Optional[datetime] = None
    motivo: Optional[str] = None
    tem_documento: bool = False
    tem_assinado: bool = False
    legado: bool = False


class TermosDoItemResponse(BaseModel):
    revsign_ativo: bool
    vigente: Optional[TermoResponse] = None
    historico: Paginated[TermoResponse]


class SolicitarTermoRequest(BaseModel):
    # Horário de Brasília. Valor com fuso é convertido para Brasília e fica sem fuso.
    agendado_para: Optional[datetime] = None

    @field_validator("agendado_para")
    @classmethod
    def _para_brasilia(cls, v: Optional[datetime]) -> Optional[datetime]:
        if v is not None and v.tzinfo is not None:
            return v.astimezone(FUSO_BRASILIA).replace(tzinfo=None)
        return v


class CancelarTermoRequest(BaseModel):
    motivo: Optional[str] = Field(None, max_length=300)


class ArquivosTermoResponse(BaseModel):
    documento_url: Optional[str] = None
    assinado_url: Optional[str] = None
