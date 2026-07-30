from typing import Optional

from fastapi import APIRouter, HTTPException, Depends, Request, Query

from app.schemas.history import HistoryResponse, ReverseRequest
from app.schemas.common import Paginated
from app.db.inventory_manager_db import InventoryDBManager
from app.db.user_manager_db import UserDBManager
from app.dependencies import (
    gestor_only,
    gestor_or_tecnico,
    get_inventory_db,
    get_user_db,
    limiter,
    CurrentUser,
)
from app.core.security import verify_password
from app.core.config import settings

router = APIRouter(prefix="/api/history", tags=["history"])


@router.get("", response_model=Paginated[HistoryResponse])
def list_history(
    search: Optional[str] = Query(
        default=None,
        description="Busca por operador, usuário, operação, revenda, tipo, marca, modelo ou identificador.",
    ),
    limit: int = Query(default=settings.DEFAULT_PAGE_SIZE, ge=1, le=settings.MAX_PAGE_SIZE),
    offset: int = Query(default=0, ge=0),
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    # A busca precisa acontecer no banco: com paginação server-side, filtrar no
    # cliente varreria apenas a página carregada e daria a impressão falsa de
    # ter pesquisado o histórico inteiro.
    rows, total = inv.list_history(search=search, limit=limit, offset=offset)
    return {"items": rows, "total": total}


@router.post("/{history_id}/reverse", response_model=dict)
@limiter.limit(settings.LOGIN_RATE_LIMIT)
def reverse_entry(
    request: Request,
    history_id: int,
    body: ReverseRequest,
    current_user: CurrentUser = Depends(gestor_only),
    inv: InventoryDBManager = Depends(get_inventory_db),
    user_db: UserDBManager = Depends(get_user_db),
):
    # T6: estorno é uma operação destrutiva sobre o histórico; exige
    # reconfirmação de senha do operador logado, não apenas o role.
    # Reaproveita o mesmo rate limit do login (T2) para não virar um
    # oráculo de senha (tentativas ilimitadas de adivinhação).
    # get_user_by_username e não get_user_by_id: apenas o primeiro traz a coluna
    # `password` no SELECT. Com get_user_by_id (que devolve id/username/role, sem
    # o hash, por design) o acesso a user["password"] levantava KeyError e todo
    # estorno respondia 500 — com senha certa ou errada.
    user = user_db.get_user_by_username(current_user.username)
    if not user or not verify_password(body.password, user["password"]):
        raise HTTPException(status_code=403, detail="Senha incorreta. Ação não autorizada.")

    ok, msg = inv.reverse_history_entry(history_id, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}
