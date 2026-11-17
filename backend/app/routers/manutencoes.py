from contextlib import contextmanager
from datetime import date
from decimal import Decimal, InvalidOperation
import logging
from typing import Literal, Optional

from fastapi import APIRouter, Depends, File, Form, HTTPException, Query, UploadFile

from app.core.config import settings
from app.db.manutencao_db import (
    ManutencaoDBManager, ManutencaoErro, ROTULOS_STATUS, STATUS_EM_ANDAMENTO,
)
from app.dependencies import (
    CurrentUser, get_audit_context, get_manutencao_db, gestor_or_tecnico,
)
from app.schemas.manutencoes import (
    ManutencaoCancelamento, ManutencaoCreate, ManutencaoEmReparo, ManutencaoEnvio,
    ManutencaoLista, ManutencaoResponse,
)
from app.schemas.validators import blank_to_none, data_br_para_date, validar_valor_monetario
from app.services import audit
from app.services.anexos import CATEGORIA_MANUTENCAO, resposta_anexo, salvar_anexo
from app.services.audit import AuditContext

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/manutencoes", tags=["manutencoes"])

MSG_CUSTO_INVALIDO = "Custo inválido: use números com ponto para os centavos (ex.: 150.00)."


@contextmanager
def _erros_de_negocio():
    try:
        yield
    except ManutencaoErro as e:
        raise HTTPException(status_code=e.status_code, detail=e.detail)


def _obter_ou_404(man: ManutencaoDBManager, manutencao_id: int) -> dict:
    m = man.obter(manutencao_id)
    if not m:
        raise HTTPException(status_code=404, detail="Manutenção não encontrada.")
    return m


def _auditar(ctx: AuditContext, man: ManutencaoDBManager, acao: str, manutencao_id: int, antes: Optional[dict]) -> None:
    """Auditoria depois do sucesso: diff antes/depois da manutenção, com o equipamento
    e o status dele em `depois`. Nunca derruba a operação já gravada."""
    depois = None
    try:
        depois = man.obter(manutencao_id)
    except Exception:
        logger.exception("Falha ao ler a manutenção %s para a auditoria.", manutencao_id)
    extra = (
        {"item_id": depois["item_id"], "item_status": depois["item_status"]} if depois else None
    )
    audit.record_diff(ctx, "manutencao", manutencao_id, acao, antes, depois=depois, extra=extra)


@router.get("", response_model=ManutencaoLista)
def listar_manutencoes(
    status: Optional[Literal[
        "aberta", "enviada", "em_reparo", "retornada", "sem_conserto", "cancelada", "em_andamento"
    ]] = None,
    item_id: Optional[int] = None,
    revenda: Optional[str] = None,
    tipo: Optional[str] = None,
    de: Optional[date] = None,
    ate: Optional[date] = None,
    busca: Optional[str] = None,
    limit: int = Query(default=25, ge=1, le=settings.MAX_PAGE_SIZE),
    offset: int = Query(default=0, ge=0),
    _: CurrentUser = Depends(gestor_or_tecnico),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
):
    itens, total, custo_total = man.listar(
        status=status, item_id=item_id, revenda=revenda, tipo=tipo, de=de, ate=ate,
        busca=busca, limit=limit, offset=offset,
    )
    return {"items": itens, "total": total, "custo_total": custo_total}


@router.get("/{manutencao_id}", response_model=ManutencaoResponse)
def obter_manutencao(
    manutencao_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
):
    return _obter_ou_404(man, manutencao_id)


@router.get("/{manutencao_id}/anexo")
def baixar_anexo(
    manutencao_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
):
    m = _obter_ou_404(man, manutencao_id)
    if not m.get("anexo_key"):
        raise HTTPException(status_code=404, detail="Nenhum anexo nesta manutenção.")
    return resposta_anexo(m["anexo_key"], f"manutencao_{manutencao_id}")


@router.post("", response_model=dict)
def abrir_manutencao(
    body: ManutencaoCreate,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    with _erros_de_negocio():
        manutencao_id = man.abrir(
            body.item_id, body.defeito, body.assistencia, body.observacoes, current_user.username
        )
    _auditar(ctx, man, "criar", manutencao_id, None)
    return {"detail": "Manutenção aberta.", "id": manutencao_id}


@router.post("/{manutencao_id}/enviar", response_model=dict)
def enviar_manutencao(
    manutencao_id: int,
    body: ManutencaoEnvio,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    antes = _obter_ou_404(man, manutencao_id)
    with _erros_de_negocio():
        man.enviar(
            manutencao_id, body.assistencia, data_br_para_date(body.previsao_retorno), current_user.username
        )
    _auditar(ctx, man, "enviar", manutencao_id, antes)
    return {"detail": "Envio à assistência registrado."}


@router.post("/{manutencao_id}/em-reparo", response_model=dict)
def marcar_em_reparo(
    manutencao_id: int,
    body: ManutencaoEmReparo,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    antes = _obter_ou_404(man, manutencao_id)
    with _erros_de_negocio():
        man.marcar_em_reparo(manutencao_id, data_br_para_date(body.previsao_retorno), current_user.username)
    _auditar(ctx, man, "em_reparo", manutencao_id, antes)
    return {"detail": "Manutenção marcada como em reparo."}


def _custo_do_formulario(texto: Optional[str]) -> Optional[Decimal]:
    texto = blank_to_none(texto)
    if texto is None:
        return None
    try:
        valor = Decimal(texto)
    except InvalidOperation:
        raise HTTPException(status_code=400, detail=MSG_CUSTO_INVALIDO)
    try:
        return validar_valor_monetario(valor, "custo")
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/{manutencao_id}/retorno", response_model=dict)
async def registrar_retorno(
    manutencao_id: int,
    resultado: Literal["retornada", "sem_conserto"] = Form(...),
    custo: Optional[str] = Form(default=None),
    observacoes: Optional[str] = Form(default=None),
    anexo: Optional[UploadFile] = File(default=None),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    antes = _obter_ou_404(man, manutencao_id)
    # Estado e custo conferidos ANTES de gravar o anexo: não deixa arquivo órfão no storage.
    if antes["status"] not in STATUS_EM_ANDAMENTO:
        raise HTTPException(
            status_code=400,
            detail=f"Ação não permitida para manutenção com status {ROTULOS_STATUS[antes['status']]}.",
        )
    valor = _custo_do_formulario(custo)
    anexo_key = None
    if anexo is not None and (anexo.filename or ""):
        anexo_key = await salvar_anexo(anexo, CATEGORIA_MANUTENCAO, f"manutencao_{manutencao_id}")
    with _erros_de_negocio():
        man.registrar_retorno(
            manutencao_id, resultado, valor, blank_to_none(observacoes), anexo_key, current_user.username
        )
    _auditar(ctx, man, "retornar" if resultado == "retornada" else "sem_conserto", manutencao_id, antes)
    if resultado == "sem_conserto":
        return {
            "detail": "Manutenção encerrada sem conserto. Avalie a baixa do equipamento.",
            "sugerir_baixa": True,
        }
    return {
        "detail": "Retorno registrado. O equipamento voltou a Disponível.",
        "sugerir_baixa": False,
    }


@router.post("/{manutencao_id}/cancelar", response_model=dict)
def cancelar_manutencao(
    manutencao_id: int,
    body: ManutencaoCancelamento,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    antes = _obter_ou_404(man, manutencao_id)
    with _erros_de_negocio():
        man.cancelar(manutencao_id, body.motivo, current_user.username)
    _auditar(ctx, man, "cancelar", manutencao_id, antes)
    return {"detail": "Manutenção cancelada. O equipamento voltou a Disponível."}
