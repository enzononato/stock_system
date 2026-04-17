from fastapi import APIRouter, HTTPException, Depends, UploadFile, File, Form
from typing import Optional
from datetime import datetime

from app.schemas.items import ItemCreate, ItemUpdate, ItemResponse
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import get_current_user, gestor_or_tecnico, gestor_only, CurrentUser
from app.core.storage import storage

router = APIRouter(prefix="/api/items", tags=["items"])
inv = InventoryDBManager()


@router.get("", response_model=list[ItemResponse])
def list_items(
    tipo: Optional[str] = None,
    status: Optional[str] = None,
    revenda: Optional[str] = None,
    search: Optional[str] = None,
    _: CurrentUser = Depends(get_current_user),
):
    items = inv.list_items()
    if tipo:
        items = [i for i in items if i.get("tipo") == tipo]
    if status:
        items = [i for i in items if i.get("status") == status]
    if revenda:
        items = [i for i in items if i.get("revenda") == revenda]
    if search:
        s = search.lower()
        items = [
            i for i in items
            if s in (i.get("brand") or "").lower()
            or s in (i.get("model") or "").lower()
            or s in (i.get("identificador") or "").lower()
            or s in (i.get("assigned_to") or "").lower()
            or s in (i.get("host") or "").lower()
        ]
    return items


@router.get("/{item_id}", response_model=ItemResponse)
def get_item(item_id: int, _: CurrentUser = Depends(get_current_user)):
    item = inv.find(item_id)
    if not item:
        raise HTTPException(status_code=404, detail="Item não encontrado.")
    return item


@router.post("", response_model=dict)
def create_item(body: ItemCreate, current_user: CurrentUser = Depends(gestor_or_tecnico)):
    try:
        date_reg = datetime.strptime(body.date_registered, "%d/%m/%Y")
    except ValueError:
        raise HTTPException(status_code=400, detail="Formato de data inválido. Use dd/mm/yyyy.")

    item_data = body.model_dump(exclude_none=True, exclude={"date_registered"})
    item_data["date_registered"] = date_reg

    ok, result = inv.add_item(item_data, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=result)
    return {"detail": "Item cadastrado com sucesso.", "id": result}


@router.put("/{item_id}", response_model=dict)
def update_item(item_id: int, body: ItemUpdate, current_user: CurrentUser = Depends(gestor_or_tecnico)):
    existing = inv.find(item_id)
    if not existing:
        raise HTTPException(status_code=404, detail="Item não encontrado.")

    item_data = body.model_dump(exclude_none=True)
    if not item_data:
        raise HTTPException(status_code=400, detail="Nenhum campo para atualizar.")

    ok, msg = inv.update_item(item_id, item_data, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


@router.delete("/{item_id}", response_model=dict)
async def remove_item(
    item_id: int,
    reason: str = Form(...),
    attachment: Optional[UploadFile] = File(default=None),
    current_user: CurrentUser = Depends(gestor_only),
):
    storage_key = None
    if attachment:
        content = await attachment.read()
        storage_key = storage.save("notas_remocao", f"remocao_{item_id}_{attachment.filename}", content)

    ok, msg = inv.remove(item_id, current_user.username, reason, storage_key)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}
