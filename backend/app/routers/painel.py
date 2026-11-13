"""Painel inicial (F7): indicadores calculados no servidor e alertas calculados na hora."""
from fastapi import APIRouter, Depends

from app.db.painel_db import PainelDBManager
from app.dependencies import CurrentUser, get_current_user, get_painel_db, gestor_or_tecnico
from app.schemas.painel import AlertasPainel, ResumoPainel
from app.services.painel import calcular_alertas

router = APIRouter(prefix="/api/painel", tags=["painel"])


@router.get("/resumo", response_model=ResumoPainel)
def resumo_do_painel(
    _: CurrentUser = Depends(get_current_user),
    db: PainelDBManager = Depends(get_painel_db),
):
    return db.resumo()


@router.get("/alertas", response_model=AlertasPainel)
def alertas_do_painel(
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: PainelDBManager = Depends(get_painel_db),
):
    return calcular_alertas(db)
