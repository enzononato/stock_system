from fastapi import APIRouter, HTTPException, Depends

from app.schemas.history import HistoryResponse
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import get_current_user, gestor_only, gestor_or_tecnico, CurrentUser

router = APIRouter(prefix="/api/history", tags=["history"])
inv = InventoryDBManager()


@router.get("", response_model=list[HistoryResponse])
def list_history(_: CurrentUser = Depends(gestor_or_tecnico)):
    return inv.list_history()


@router.post("/{history_id}/reverse", response_model=dict)
def reverse_entry(history_id: int, current_user: CurrentUser = Depends(gestor_only)):
    ok, msg = inv.reverse_history_entry(history_id, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}
