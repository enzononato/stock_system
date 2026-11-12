"""Consulta da trilha de auditoria (somente Gestor)."""
import logging
from datetime import date
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query

from app.db.audit_db import AuditDBManager
from app.dependencies import CurrentUser, get_audit_db, gestor_only
from app.schemas.audit import AuditEntry
from app.schemas.common import Paginated

logger = logging.getLogger(__name__)

# Faixa aceita em `de`/`ate`: dentro do que o DATETIME do MySQL suporta e longe
# do estouro de `date` ao somar um dia em `ate`.
_DATA_MINIMA = date(1000, 1, 1)
_DATA_MAXIMA = date(9998, 12, 31)

router = APIRouter(prefix="/api/audit", tags=["auditoria"])


@router.get("", response_model=Paginated[AuditEntry])
def listar_auditoria(
    entidade: Optional[str] = None,
    entidade_id: Optional[str] = None,
    username: Optional[str] = None,
    de: Optional[date] = Query(None, description="Data inicial (AAAA-MM-DD), em UTC."),
    ate: Optional[date] = Query(None, description="Data final (AAAA-MM-DD), em UTC."),
    limit: int = Query(50, ge=1, le=200),
    offset: int = Query(0, ge=0),
    _: CurrentUser = Depends(gestor_only),
    db: AuditDBManager = Depends(get_audit_db),
):
    for limite in (de, ate):
        if limite is not None and not (_DATA_MINIMA <= limite <= _DATA_MAXIMA):
            raise HTTPException(status_code=422, detail="Período inválido.")
    if de is not None and ate is not None and de > ate:
        raise HTTPException(status_code=422, detail="Período inválido.")
    try:
        rows, total = db.listar(
            entidade=entidade,
            entidade_id=entidade_id,
            username=username,
            de=de,
            ate=ate,
            limit=limit,
            offset=offset,
        )
    except Exception:
        logger.exception("Falha ao consultar a auditoria.")
        raise HTTPException(status_code=500, detail="Erro ao consultar a auditoria.")
    return {"items": rows, "total": total}
