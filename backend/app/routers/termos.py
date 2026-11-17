"""API dos termos: termo vigente e histórico do equipamento e do colaborador, pedido de
assinatura, cancelamento, reenvio e anexo do PDF assinado."""
from typing import Optional

from fastapi import APIRouter, Body, Depends, File, HTTPException, Query, UploadFile
from fastapi.concurrency import run_in_threadpool

from app.core.storage import storage
from app.core.tempo import agora_local
from app.db.colaborador_db import ColaboradorDBManager
from app.db.inventory_manager_db import InventoryDBManager
from app.db.termo_db import TermoDBManager
from app.dependencies import (
    CurrentUser,
    get_audit_context,
    get_colaborador_db,
    get_inventory_db,
    get_termo_db,
    gestor_or_tecnico,
)
from app.schemas.common import Paginated
from app.schemas.termos import (
    ArquivosTermoResponse,
    CancelarTermoRequest,
    SolicitarTermoRequest,
    TermoResponse,
    TermosDoItemResponse,
)
from app.services import termos
from app.services.audit import AuditContext
from app.services.termos_confirmacao import anexar_assinado
from app.services.termos_trava import MSG_PROCESSANDO

router = APIRouter(tags=["termos"])

MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20MB


def _obter_ou_404(db: TermoDBManager, termo_id: int) -> dict:
    termo = db.obter(termo_id)
    if termo is None:
        raise HTTPException(status_code=404, detail="Solicitação de termo não encontrada.")
    return termo


def _chamar(func, *args, **kwargs):
    try:
        return func(*args, **kwargs)
    except termos.TermoErro as e:
        # Trava ocupada (assinatura sendo processada): 409, como nas rotas de confirmação.
        status = 409 if str(e) == MSG_PROCESSANDO else 400
        raise HTTPException(status_code=status, detail=str(e)) from None


@router.get("/api/items/{item_id}/termos", response_model=TermosDoItemResponse)
def termos_do_item(
    item_id: int,
    limit: int = Query(default=25, ge=1, le=100),
    offset: int = Query(default=0, ge=0),
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    db: TermoDBManager = Depends(get_termo_db),
):
    if not inv.find(item_id):
        raise HTTPException(status_code=404, detail="Equipamento não encontrado.")
    linhas, total = db.listar_por_item(item_id, limit=limit, offset=offset)
    vigente = db.vigente(item_id)
    return {
        "revsign_ativo": termos.revsign_ligado(),
        "vigente": termos.para_resposta(vigente) if vigente else None,
        "historico": {"items": [termos.para_resposta(t) for t in linhas], "total": total},
    }


@router.post("/api/items/{item_id}/termos", response_model=TermoResponse, status_code=201)
def solicitar_termo(
    item_id: int,
    body: Optional[SolicitarTermoRequest] = Body(None),
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    item = inv.find(item_id)
    if not item:
        raise HTTPException(status_code=404, detail="Equipamento não encontrado.")
    tipo = termos.TIPO_POR_STATUS_ITEM.get(item["status"])
    if tipo is None:
        raise HTTPException(
            status_code=400, detail="Só é possível solicitar assinatura de equipamento com colaborador."
        )
    agendado_para = body.agendado_para if body else None
    if agendado_para is not None:
        if not termos.revsign_ligado():
            raise HTTPException(
                status_code=400, detail="Agendamento disponível só com a integração do RevSign ligada."
            )
        if agendado_para <= agora_local():
            raise HTTPException(status_code=400, detail="A data agendada precisa ser no futuro.")
    history_id = {
        "entrega": inv.ultimo_history_id(item_id, "Empréstimo"),
        "devolucao": inv.ultimo_history_id(item_id, "Devolução"),
    }.get(tipo)
    termo = _chamar(termos.solicitar, item_id, tipo, ctx, history_id=history_id, agendado_para=agendado_para)
    return termos.para_resposta(termo)


@router.get("/api/colaboradores/{colaborador_id}/termos", response_model=Paginated[TermoResponse])
def termos_do_colaborador(
    colaborador_id: int,
    limit: int = Query(default=25, ge=1, le=100),
    offset: int = Query(default=0, ge=0),
    _: CurrentUser = Depends(gestor_or_tecnico),
    colaboradores: ColaboradorDBManager = Depends(get_colaborador_db),
    db: TermoDBManager = Depends(get_termo_db),
):
    if not colaboradores.obter(colaborador_id):
        raise HTTPException(status_code=404, detail="Colaborador não encontrado.")
    linhas, total = db.listar_por_colaborador(colaborador_id, limit=limit, offset=offset)
    return {"items": [termos.para_resposta(t) for t in linhas], "total": total}


@router.post("/api/termos/{termo_id}/cancelar", response_model=TermoResponse)
def cancelar_termo(
    termo_id: int,
    body: Optional[CancelarTermoRequest] = Body(None),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    db: TermoDBManager = Depends(get_termo_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    _obter_ou_404(db, termo_id)
    motivo = (body.motivo.strip() if body and body.motivo else "") or f"Cancelada por {current_user.username}."
    return termos.para_resposta(_chamar(termos.cancelar, termo_id, ctx, motivo))


@router.post("/api/termos/{termo_id}/reenviar", response_model=TermoResponse)
def reenviar_termo(
    termo_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: TermoDBManager = Depends(get_termo_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    _obter_ou_404(db, termo_id)
    return termos.para_resposta(_chamar(termos.reenviar, termo_id, ctx))


@router.post("/api/termos/{termo_id}/assinado", response_model=TermoResponse)
async def anexar_termo_assinado(
    termo_id: int,
    signed_pdf: UploadFile = File(...),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    db: TermoDBManager = Depends(get_termo_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    _obter_ou_404(db, termo_id)
    conteudo = await signed_pdf.read()
    if len(conteudo) > MAX_UPLOAD_SIZE:
        raise HTTPException(status_code=413, detail="Arquivo excede o tamanho máximo de 20MB.")
    termo = await run_in_threadpool(
        _chamar, anexar_assinado, termo_id, conteudo, signed_pdf.filename, current_user.username, ctx
    )
    return termos.para_resposta(termo)


@router.get("/api/termos/{termo_id}/arquivos", response_model=ArquivosTermoResponse)
def arquivos_do_termo(
    termo_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: TermoDBManager = Depends(get_termo_db),
):
    termo = _obter_ou_404(db, termo_id)

    def _url(chave: Optional[str]) -> Optional[str]:
        return storage.get_url(chave) if chave and storage.exists(chave) else None

    return {
        "documento_url": _url(termo["documento_key"]),
        "assinado_url": _url(termo["arquivo_assinado_key"]),
    }
