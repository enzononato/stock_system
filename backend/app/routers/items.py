import os
from fastapi import APIRouter, HTTPException, Depends, UploadFile, File, Form, Query
from typing import Optional
from datetime import datetime

from app.schemas.items import ItemCreate, ItemUpdate, ItemResponse
from app.schemas.common import Paginated
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import get_current_user, gestor_or_tecnico, gestor_only, get_inventory_db, CurrentUser
from app.core.storage import storage
from app.core.config import settings

router = APIRouter(prefix="/api/items", tags=["items"])

MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20MB


def _safe_filename(name: str) -> str:
    return os.path.basename(name or "").replace("\\", "").replace("/", "") or "arquivo"


@router.get("", response_model=Paginated[ItemResponse])
def list_items(
    tipo: Optional[str] = None,
    status: Optional[str] = None,
    revenda: Optional[str] = None,
    search: Optional[str] = None,
    limit: int = Query(default=settings.DEFAULT_PAGE_SIZE, ge=1, le=settings.MAX_PAGE_SIZE),
    offset: int = Query(default=0, ge=0),
    _: CurrentUser = Depends(get_current_user),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    # T10: os filtros são repassados para a camada de dados (SQL), em vez de
    # filtrar em Python depois de carregar a tabela inteira em memória.
    items, total = inv.list_items(
        tipo=tipo, status=status, revenda=revenda, search=search, limit=limit, offset=offset,
    )
    return {"items": items, "total": total}


@router.get("/{item_id}", response_model=ItemResponse)
def get_item(
    item_id: int,
    _: CurrentUser = Depends(get_current_user),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    item = inv.find(item_id)
    if not item:
        raise HTTPException(status_code=404, detail="Item não encontrado.")
    return item


@router.post("", response_model=dict)
def create_item(
    body: ItemCreate,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    # O schema (ItemCreate) já validou o formato dd/mm/yyyy antes de chegar
    # aqui (T7); o strptime abaixo não deve mais estourar em uso normal.
    date_reg = datetime.strptime(body.date_registered, "%d/%m/%Y")

    item_data = body.model_dump(exclude_none=True, exclude={"date_registered"})
    item_data["date_registered"] = date_reg

    ok, result = inv.add_item(item_data, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=result)
    return {"detail": "Item cadastrado com sucesso.", "id": result}


@router.put("/{item_id}", response_model=dict)
def update_item(
    item_id: int,
    body: ItemUpdate,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
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
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    storage_key = None
    if attachment:
        content = await attachment.read()
        if len(content) > MAX_UPLOAD_SIZE:
            raise HTTPException(status_code=413, detail="Arquivo excede o tamanho máximo de 20MB.")
        storage_key = storage.save(
            "notas_remocao", f"remocao_{item_id}_{_safe_filename(attachment.filename)}", content
        )

    ok, msg = inv.remove(item_id, current_user.username, reason, storage_key)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}
