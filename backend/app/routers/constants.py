"""Endpoint que expõe as constantes de domínio para o frontend."""
from fastapi import APIRouter, Depends
from app.dependencies import get_current_user, get_unidade_db, CurrentUser
from app.db.unidade_db import UnidadeDBManager
from app.core.config import (
    CENTER_COST_OPTIONS, SETORES_OPTIONS,
    EQUIPMENT_TYPES, PERIPHERAL_TYPES, REMOVAL_REASONS,
)

router = APIRouter(prefix="/api/constants", tags=["constants"])


@router.get("")
def get_constants(
    _: CurrentUser = Depends(get_current_user),
    unidade_db: UnidadeDBManager = Depends(get_unidade_db),
):
    # "revendas" deixou de vir da constante fixa REVENDAS_OPTIONS (app/core/config.py)
    # e passa a vir das Unidades cadastradas no banco (só as ativas, só os nomes,
    # já ordenados por UnidadeDBManager.list_unidades). Banco vazio -> lista vazia,
    # sem estourar.
    unidades_ativas = unidade_db.list_unidades(include_inactive=False)
    revendas = [u["nome"] for u in unidades_ativas]
    return {
        "center_costs": CENTER_COST_OPTIONS,
        "revendas": revendas,
        "setores": SETORES_OPTIONS,
        "equipment_types": EQUIPMENT_TYPES,
        "peripheral_types": PERIPHERAL_TYPES,
        "removal_reasons": list(REMOVAL_REASONS.keys()),
        "removal_reasons_attachment": REMOVAL_REASONS,
    }
