"""Etiquetas com QR code: dados para a folha A4 e seleção dos itens de uma unidade.

O QR carrega só `APP_PUBLIC_URL/e/{id}`; nenhum dado pessoal sai daqui.
"""
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException

from app.core import config
from app.db.inventario_db import InventarioDBManager
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import (
    CurrentUser,
    get_inventario_db,
    get_inventory_db,
    gestor_or_tecnico,
)
from app.schemas.inventarios import EtiquetasRequest, EtiquetasResponse, IdsEtiquetas

router = APIRouter(prefix="/api/etiquetas", tags=["etiquetas"])


@router.get("/ids", response_model=IdsEtiquetas)
def ids_da_unidade(
    revenda: Optional[str] = None,
    tipo: Optional[str] = None,
    status: Optional[str] = None,
    search: Optional[str] = None,
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    """Ids (crescentes) de todos os itens ativos da unidade que casam com os filtros."""
    if not revenda or not revenda.strip():
        raise HTTPException(status_code=422, detail="Escolha a unidade.")
    itens, _total = inv.list_items(
        tipo=tipo or None, status=status or None, revenda=revenda.strip(),
        search=search or None, limit=None,
    )
    ids = sorted(item["id"] for item in itens)
    return {"ids": ids, "total": len(ids)}


@router.post("", response_model=EtiquetasResponse)
def gerar_etiquetas(
    body: EtiquetasRequest,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
):
    # `settings` é lido na hora da chamada: os testes trocam o objeto em `app.core.config`.
    base_url = config.settings.APP_PUBLIC_URL.rstrip("/")
    dados = db.dados_etiquetas(body.item_ids)
    etiquetas = [{**d, "url": f"{base_url}/e/{d['id']}"} for d in dados]
    achados = {d["id"] for d in dados}
    return {
        "base_url": base_url,
        "etiquetas": etiquetas,
        "ignorados": sorted(i for i in body.item_ids if i not in achados),
    }
