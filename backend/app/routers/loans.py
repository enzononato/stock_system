import os
from fastapi import APIRouter, HTTPException, Depends, UploadFile, File
from datetime import datetime

from app.schemas.loans import LoanRequest
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import gestor_or_tecnico, get_inventory_db, get_audit_context, CurrentUser
from app.services import audit
from app.services.audit import AuditContext
from app.core.storage import storage

router = APIRouter(prefix="/api/loans", tags=["loans"])

MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20MB


def _safe_filename(name: str) -> str:
    return os.path.basename(name or "").replace("\\", "").replace("/", "") or "arquivo"


def _check_size(content: bytes) -> None:
    if len(content) > MAX_UPLOAD_SIZE:
        raise HTTPException(status_code=413, detail="Arquivo excede o tamanho máximo de 20MB.")


def _auditar_etapa(
    ctx: AuditContext, inv: InventoryDBManager, entidade: str, acao: str,
    item_id: int, antes_item: "dict | None", extra: "dict | None" = None,
) -> None:
    """Grava a mudança do item (só os campos alterados) numa etapa do fluxo de
    empréstimo/devolução. `extra` entra só no lado "depois"."""
    antes, depois = audit.diff(antes_item or {}, inv.find(item_id) or {})
    if extra:
        depois = {**(depois or {}), **extra}
    audit.record(ctx, entidade, item_id, acao, antes=antes, depois=depois)


@router.post("", response_model=dict)
def initiate_loan(
    body: LoanRequest,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    antes_item = inv.find(body.item_id)
    ok, msg = inv.issue(
        body.item_id,
        body.usuario,
        body.cpf,
        body.center_cost,
        body.cargo,
        body.setor,
        body.revenda,
        body.date_issue,
        current_user.username,
        pessoa_juridica=body.pessoa_juridica,
    )
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    _auditar_etapa(
        ctx, inv, "emprestimo", "criar", body.item_id, antes_item,
        extra={"cargo": body.cargo, "center_cost": body.center_cost, "setor": body.setor},
    )
    return {"detail": msg}


@router.post("/{item_id}/confirm", response_model=dict)
async def confirm_loan(
    item_id: int,
    signed_pdf: UploadFile = File(...),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    content = await signed_pdf.read()
    _check_size(content)
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"termo_{item_id}_{timestamp}_{_safe_filename(signed_pdf.filename)}"
    storage_key = storage.save("termos_assinados", filename, content)

    antes_item = inv.find(item_id)
    ok, msg = inv.confirm_loan(item_id, current_user.username, storage_key)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    _auditar_etapa(
        ctx, inv, "emprestimo", "confirmar", item_id, antes_item,
        extra={"termo_assinado": storage_key},
    )
    return {"detail": msg}


@router.get("/{item_id}/signed-term-url", response_model=dict)
def get_signed_term_url(
    item_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    """Alimenta o botão "Ver Termo" da tela de Termos: devolve a URL do PDF
    assinado do empréstimo ativo do item, sem expor a chave de storage crua."""
    item = inv.find(item_id)
    if not item:
        raise HTTPException(status_code=404, detail="Item não encontrado.")
    if item["status"] != "Indisponível":
        raise HTTPException(status_code=400, detail="Item não está com empréstimo ativo.")
    key = inv.get_active_signed_term_key(item_id)
    if not key or not storage.exists(key):
        raise HTTPException(status_code=404, detail="Termo assinado não encontrado.")
    return {"url": storage.get_url(key)}


@router.post("/{item_id}/return/initiate", response_model=dict)
def initiate_return(
    item_id: int,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    antes_item = inv.find(item_id)
    ok, result, filename = inv.generate_return_term_bytes(item_id, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=result)
    _auditar_etapa(ctx, inv, "devolucao", "criar", item_id, antes_item)

    # Salva o DOCX gerado no storage
    storage_key = storage.save("termos_devolucao", filename, result)
    download_url = storage.get_url(storage_key)

    return {"detail": "Termo de devolução gerado.", "download_url": download_url, "filename": filename}


@router.post("/{item_id}/return/confirm", response_model=dict)
async def confirm_return(
    item_id: int,
    signed_pdf: UploadFile = File(...),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    content = await signed_pdf.read()
    _check_size(content)
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"termo_devolucao_{item_id}_{timestamp}_{_safe_filename(signed_pdf.filename)}"
    storage_key = storage.save("termos_devolucao_assinados", filename, content)

    antes_item = inv.find(item_id)
    ok, msg = inv.confirm_return(item_id, current_user.username, storage_key)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    _auditar_etapa(
        ctx, inv, "devolucao", "confirmar", item_id, antes_item,
        extra={"termo_assinado": storage_key},
    )
    return {"detail": msg}
