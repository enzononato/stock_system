"""Endpoint que expõe as constantes de domínio para o frontend."""
from fastapi import APIRouter, Depends
from app.dependencies import get_current_user, CurrentUser
from app.core.config import (
    CENTER_COST_OPTIONS, REVENDAS_OPTIONS, SETORES_OPTIONS,
    EQUIPMENT_TYPES, PERIPHERAL_TYPES, REMOVAL_REASONS,
)

router = APIRouter(prefix="/api/constants", tags=["constants"])


@router.get("")
def get_constants(_: CurrentUser = Depends(get_current_user)):
    return {
        "center_costs": CENTER_COST_OPTIONS,
        "revendas": REVENDAS_OPTIONS,
        "setores": SETORES_OPTIONS,
        "equipment_types": EQUIPMENT_TYPES,
        "peripheral_types": PERIPHERAL_TYPES,
        "removal_reasons": list(REMOVAL_REASONS.keys()),
        "removal_reasons_attachment": REMOVAL_REASONS,
    }
