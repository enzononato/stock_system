"""Inventário físico por unidade: abertura, leituras, encerramento e relatório."""
from typing import Literal, Optional

from fastapi import APIRouter, Depends, HTTPException, Query

from app.db.inventario_db import (
    InventarioDBManager,
    InventarioFechado,
    InventarioJaAberto,
    InventarioNaoEncontrado,
    UnidadeNaoEncontrada,
)
from app.dependencies import (
    CurrentUser,
    get_audit_context,
    get_inventario_db,
    gestor_only,
    gestor_or_tecnico,
)
from app.schemas.common import Paginated
from app.schemas.inventarios import (
    CancelarInventario,
    InventarioCreate,
    InventarioDetalhe,
    InventarioResponse,
)
from app.services import audit
from app.services.audit import AuditContext

router = APIRouter(prefix="/api/inventarios", tags=["inventarios"])

MSG_NAO_ENCONTRADO = "Inventário não encontrado."


def _detalhe(db: InventarioDBManager, inventario_id: int) -> dict:
    inventario = db.obter(inventario_id)
    if not inventario:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADO)
    return {**inventario, "resumo": db.resumo(inventario)}


@router.get("", response_model=Paginated[InventarioResponse])
def listar_inventarios(
    unidade_id: Optional[int] = None,
    status: Optional[Literal["aberto", "encerrado", "cancelado"]] = None,
    limit: int = Query(25, ge=1, le=100),
    offset: int = Query(0, ge=0),
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
):
    itens, total = db.listar(unidade_id=unidade_id, status=status, limit=limit, offset=offset)
    return {"items": itens, "total": total}


@router.post("", response_model=InventarioDetalhe, status_code=201)
def abrir_inventario(
    body: InventarioCreate,
    user: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        inventario_id = db.abrir(body.unidade_id, body.observacao, user.username)
    except UnidadeNaoEncontrada:
        raise HTTPException(status_code=404, detail="Unidade não encontrada.")
    except InventarioJaAberto as e:
        raise HTTPException(status_code=409, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    detalhe = _detalhe(db, inventario_id)
    audit.record(
        ctx, "inventario", inventario_id, "criar",
        depois={
            "unidade_id": detalhe["unidade_id"],
            "unidade_nome": detalhe["unidade_nome"],
            "observacao": detalhe["observacao"],
        },
    )
    return detalhe


@router.get("/{inventario_id}", response_model=InventarioDetalhe)
def obter_inventario(
    inventario_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
):
    return _detalhe(db, inventario_id)


@router.post("/{inventario_id}/cancelar", response_model=InventarioDetalhe)
def cancelar_inventario(
    inventario_id: int,
    body: CancelarInventario,
    user: CurrentUser = Depends(gestor_only),
    db: InventarioDBManager = Depends(get_inventario_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        db.cancelar(inventario_id, body.motivo, user.username)
    except InventarioNaoEncontrado:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADO)
    except InventarioFechado as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.record(
        ctx, "inventario", inventario_id, "cancelar",
        antes={"status": "aberto"}, depois={"status": "cancelado", "motivo": body.motivo},
    )
    return _detalhe(db, inventario_id)
