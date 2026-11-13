from fastapi import APIRouter, HTTPException, Depends, UploadFile, File

from app.schemas.loans import LoanRequest
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import gestor_or_tecnico, get_inventory_db, get_audit_context, CurrentUser
from app.services import audit
from app.services.audit import AuditContext
from app.core.storage import storage
from app.services.termos_confirmacao import confirmar_devolucao, confirmar_emprestimo

router = APIRouter(prefix="/api/loans", tags=["loans"])

MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20MB


def _check_size(content: bytes) -> None:
    if len(content) > MAX_UPLOAD_SIZE:
        raise HTTPException(status_code=413, detail="Arquivo excede o tamanho máximo de 20MB.")


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
    audit.record_diff(
        ctx, "emprestimo", body.item_id, "criar", antes_item, obter_depois=lambda: inv.find(body.item_id),
        extra={"cargo": body.cargo, "center_cost": body.center_cost, "setor": body.setor},
    )
    return {"detail": msg}


@router.post("/{item_id}/confirm", response_model=dict)
async def confirm_loan(
    item_id: int,
    signed_pdf: UploadFile = File(...),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    ctx: AuditContext = Depends(get_audit_context),
):
    content = await signed_pdf.read()
    _check_size(content)
    ok, msg, _ = confirmar_emprestimo(item_id, content, signed_pdf.filename, current_user.username, ctx)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
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
    audit.record_diff(
        ctx, "devolucao", item_id, "criar", antes_item, obter_depois=lambda: inv.find(item_id)
    )

    # Salva o DOCX gerado no storage
    storage_key = storage.save("termos_devolucao", filename, result)
    download_url = storage.get_url(storage_key)

    return {"detail": "Termo de devolução gerado.", "download_url": download_url, "filename": filename}


@router.post("/{item_id}/return/confirm", response_model=dict)
async def confirm_return(
    item_id: int,
    signed_pdf: UploadFile = File(...),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    ctx: AuditContext = Depends(get_audit_context),
):
    content = await signed_pdf.read()
    _check_size(content)
    ok, msg, _ = confirmar_devolucao(item_id, content, signed_pdf.filename, current_user.username, ctx)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}
