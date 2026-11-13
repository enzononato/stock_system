"""Painel inicial (F7): indicadores calculados no servidor e alertas calculados na hora."""
from fastapi import APIRouter, Depends

from app.db.painel_db import PainelDBManager
from app.dependencies import CurrentUser, get_current_user, get_painel_db
from app.schemas.painel import ResumoPainel

router = APIRouter(prefix="/api/painel", tags=["painel"])


@router.get("/resumo", response_model=ResumoPainel)
def resumo_do_painel(
    _: CurrentUser = Depends(get_current_user),
    db: PainelDBManager = Depends(get_painel_db),
):
    return db.resumo()
