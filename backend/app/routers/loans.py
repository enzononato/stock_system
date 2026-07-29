from fastapi import APIRouter, HTTPException, Depends, UploadFile, File
from datetime import datetime

from app.schemas.loans import LoanRequest
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import gestor_or_tecnico, get_inventory_db, CurrentUser
from app.core.storage import storage

router = APIRouter(prefix="/api/loans", tags=["loans"])


@router.post("", response_model=dict)
def initiate_loan(
    body: LoanRequest,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
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
    )
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


@router.post("/{item_id}/confirm", response_model=dict)
async def confirm_loan(
    item_id: int,
    signed_pdf: UploadFile = File(...),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    content = await signed_pdf.read()
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"termo_{item_id}_{timestamp}_{signed_pdf.filename}"
    storage_key = storage.save("termos_assinados", filename, content)

    ok, msg = inv.confirm_loan(item_id, current_user.username, storage_key)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


@router.post("/{item_id}/return/initiate", response_model=dict)
def initiate_return(
    item_id: int,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    ok, result, filename = inv.generate_return_term_bytes(item_id, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=result)

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
):
    content = await signed_pdf.read()
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"termo_devolucao_{item_id}_{timestamp}_{signed_pdf.filename}"
    storage_key = storage.save("termos_devolucao_assinados", filename, content)

    ok, msg = inv.confirm_return(item_id, current_user.username, storage_key)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}
