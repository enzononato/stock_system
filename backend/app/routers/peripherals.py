import os
import re
from fastapi import APIRouter, HTTPException, Depends, UploadFile, File, Form
from typing import Optional

from app.schemas.peripherals import PeripheralCreate, PeripheralResponse
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import (
    get_current_user, gestor_or_tecnico, get_inventory_db, get_audit_context, CurrentUser,
)
from app.services import audit
from app.services.audit import AuditContext
from app.core.storage import storage

router = APIRouter(tags=["peripherals"])

MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20MB


def _safe_filename(name: str) -> str:
    return os.path.basename(name or "").replace("\\", "").replace("/", "") or "arquivo"


@router.get("/api/peripherals", response_model=list[PeripheralResponse])
def list_peripherals(
    status: Optional[str] = None,
    tipo: Optional[str] = None,
    include_inactive: bool = False,
    _: CurrentUser = Depends(get_current_user),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    return inv.list_peripherals(
        status_filter=status or "",
        type_filter=tipo or "",
        include_inactive=include_inactive,
    )


@router.post("/api/peripherals", response_model=dict)
def create_peripheral(
    body: PeripheralCreate,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    data = body.model_dump(exclude_none=True)
    ok, msg = inv.add_peripheral(data, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    # add_peripheral só devolve o id dentro da mensagem ("... cadastrado com ID 12.").
    achado = re.search(r"ID (\d+)\.?$", msg)
    perif_id = int(achado.group(1)) if achado else None
    audit.record(
        ctx, "periferico", perif_id, "criar",
        depois=inv.get_peripheral(perif_id) if perif_id is not None else data,
    )
    return {"detail": msg}


@router.delete("/api/peripherals/{peripheral_id}", response_model=dict)
def deactivate_peripheral(
    peripheral_id: int,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    """T9: inativa um periférico (Gestor ou Técnico). A camada de dados
    recusa (400) caso o periférico esteja em uso/vinculado."""
    antes_perif = inv.get_peripheral(peripheral_id)
    ok, msg = inv.deactivate_peripheral(peripheral_id, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    antes, depois = audit.diff(antes_perif or {}, inv.get_peripheral(peripheral_id) or {})
    audit.record(ctx, "periferico", peripheral_id, "excluir", antes=antes, depois=depois)
    return {"detail": msg}


@router.get("/api/items/{item_id}/peripherals", response_model=list[PeripheralResponse])
def list_item_peripherals(
    item_id: int,
    _: CurrentUser = Depends(get_current_user),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    return inv.list_peripherals_for_equipment(item_id)


@router.post("/api/items/{item_id}/peripherals/{peripheral_id}", response_model=dict)
def link_peripheral(
    item_id: int,
    peripheral_id: int,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    antes_perif = inv.get_peripheral(peripheral_id) or {}
    ok, msg = inv.link_peripheral_to_equipment(item_id, peripheral_id, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    antes, depois = audit.diff(
        {**antes_perif, "equipamento_id": None},
        {**(inv.get_peripheral(peripheral_id) or {}), "equipamento_id": item_id},
    )
    audit.record(ctx, "periferico", peripheral_id, "vincular", antes=antes, depois=depois)
    return {"detail": msg}


@router.delete("/api/peripherals/links/{link_id}", response_model=dict)
def unlink_peripheral(
    link_id: int,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    vinculo = inv.get_peripheral_link(link_id)
    antes_perif = (inv.get_peripheral(vinculo["peripheral_id"]) if vinculo else None) or {}
    ok, msg = inv.unlink_peripheral_from_equipment(link_id, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    if vinculo:
        antes, depois = audit.diff(
            {**antes_perif, "equipamento_id": vinculo["equipment_id"]},
            {**(inv.get_peripheral(vinculo["peripheral_id"]) or {}), "equipamento_id": None},
        )
        audit.record(
            ctx, "periferico", vinculo["peripheral_id"], "desvincular", antes=antes, depois=depois
        )
    return {"detail": msg}


@router.post("/api/items/{item_id}/peripherals/{old_peripheral_id}/replace", response_model=dict)
async def replace_peripheral(
    item_id: int,
    old_peripheral_id: int,
    new_peripheral_id: int = Form(...),
    reason: str = Form(...),
    attachment: Optional[UploadFile] = File(default=None),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    if not reason or not reason.strip():
        raise HTTPException(status_code=400, detail="Informe o motivo da substituição.")

    storage_key = None
    if attachment:
        content = await attachment.read()
        if len(content) > MAX_UPLOAD_SIZE:
            raise HTTPException(status_code=413, detail="Arquivo excede o tamanho máximo de 20MB.")
        storage_key = storage.save(
            "notas_remocao",
            f"substituicao_{old_peripheral_id}_{_safe_filename(attachment.filename)}",
            content,
        )
    antes_perif = inv.get_peripheral(old_peripheral_id) or {}
    ok, msg = inv.replace_peripheral(
        item_id, old_peripheral_id, new_peripheral_id, reason, current_user.username, storage_key
    )
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    antes, depois = audit.diff(
        {**antes_perif, "equipamento_id": item_id, "substituido_por": None},
        {
            **(inv.get_peripheral(old_peripheral_id) or {}),
            "equipamento_id": None,
            "substituido_por": new_peripheral_id,
        },
    )
    audit.record(ctx, "periferico", old_peripheral_id, "substituir", antes=antes, depois=depois)
    return {"detail": msg}
