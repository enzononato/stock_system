"""Eventos de acesso (F8): admissão, desligamento e mudança de função."""
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query

from app.db.acesso_db import AcessoErro
from app.db.acesso_evento_db import AcessoEventoDBManager
from app.dependencies import (
    CurrentUser,
    get_acesso_evento_db,
    get_audit_context,
    gestor_or_tecnico,
)
from app.routers.acessos import erro_http
from app.schemas.acessos import (
    EventoCancelar,
    EventoCreate,
    EventoDetalhe,
    EventoResumo,
    StatusEvento,
    TarefaResolver,
    TipoEvento,
)
from app.schemas.common import Paginated
from app.services.audit import AuditContext

router = APIRouter(prefix="/api/acessos", tags=["acessos"])

MSG_EVENTO_NAO_ENCONTRADO = "Evento não encontrado."


@router.get("/eventos", response_model=Paginated[EventoResumo])
def listar_eventos(
    status: Optional[StatusEvento] = None,
    tipo: Optional[TipoEvento] = None,
    colaborador_id: Optional[int] = None,
    busca: Optional[str] = None,
    limit: int = Query(25, ge=1, le=100),
    offset: int = Query(0, ge=0),
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoEventoDBManager = Depends(get_acesso_evento_db),
):
    itens, total = db.listar(
        status=status, tipo=tipo, colaborador_id=colaborador_id, busca=busca,
        limit=limit, offset=offset,
    )
    return {"items": itens, "total": total}


@router.post("/eventos", response_model=EventoDetalhe, status_code=201)
def abrir_evento(
    body: EventoCreate,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoEventoDBManager = Depends(get_acesso_evento_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        return db.abrir(
            ctx,
            colaborador_id=body.colaborador_id,
            tipo=body.tipo,
            cargo_anterior=body.cargo_anterior,
            cargo_novo=body.cargo_novo,
            observacao=body.observacao,
        )
    except AcessoErro as e:
        raise erro_http(e)


@router.get("/eventos/{evento_id}", response_model=EventoDetalhe)
def obter_evento(
    evento_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoEventoDBManager = Depends(get_acesso_evento_db),
):
    evento = db.obter(evento_id)
    if evento is None:
        raise HTTPException(status_code=404, detail=MSG_EVENTO_NAO_ENCONTRADO)
    return evento


@router.post("/tarefas/{tarefa_id}/resolver", response_model=EventoDetalhe)
def resolver_tarefa(
    tarefa_id: int,
    body: TarefaResolver,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoEventoDBManager = Depends(get_acesso_evento_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        return db.resolver_tarefa(
            ctx, tarefa_id,
            status=body.status, login=body.login,
            item_id=body.item_id, observacao=body.observacao,
        )
    except AcessoErro as e:
        raise erro_http(e)


@router.post("/eventos/{evento_id}/concluir", response_model=EventoDetalhe)
def concluir_evento(
    evento_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoEventoDBManager = Depends(get_acesso_evento_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        return db.concluir(ctx, evento_id)
    except AcessoErro as e:
        raise erro_http(e)


@router.post("/eventos/{evento_id}/cancelar", response_model=EventoDetalhe)
def cancelar_evento(
    evento_id: int,
    body: EventoCancelar,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoEventoDBManager = Depends(get_acesso_evento_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        return db.cancelar(ctx, evento_id, motivo=body.motivo_cancelamento)
    except AcessoErro as e:
        raise erro_http(e)
