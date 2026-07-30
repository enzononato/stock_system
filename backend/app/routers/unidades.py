"""Endpoints de CRUD e indicadores das Unidades (dados jurídicos por revenda,
usados no novo termo de responsabilidade)."""
from fastapi import APIRouter, Depends, HTTPException

from app.schemas.unidades import (
    UnidadeCreate,
    UnidadeUpdate,
    UnidadeResponse,
    IndicadoresResponse,
)
from app.db.unidade_db import UnidadeDBManager
from app.dependencies import (
    get_current_user,
    gestor_only,
    gestor_or_tecnico,
    get_unidade_db,
    CurrentUser,
)

router = APIRouter(prefix="/api/unidades", tags=["unidades"])


@router.get("", response_model=list[UnidadeResponse])
def list_unidades(
    include_inactive: bool = False,
    _: CurrentUser = Depends(get_current_user),
    db: UnidadeDBManager = Depends(get_unidade_db),
):
    """Alimenta os seletores de revenda do frontend — qualquer usuário
    autenticado pode listar."""
    return db.list_unidades(include_inactive=include_inactive)


@router.get("/{unidade_id}", response_model=UnidadeResponse)
def get_unidade(
    unidade_id: int,
    _: CurrentUser = Depends(get_current_user),
    db: UnidadeDBManager = Depends(get_unidade_db),
):
    unidade = db.get_unidade(unidade_id)
    if not unidade:
        raise HTTPException(status_code=404, detail="Unidade não encontrada.")
    return unidade


@router.get("/{unidade_id}/indicadores", response_model=IndicadoresResponse)
def get_indicadores(
    unidade_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: UnidadeDBManager = Depends(get_unidade_db),
):
    unidade = db.get_unidade(unidade_id)
    if not unidade:
        raise HTTPException(status_code=404, detail="Unidade não encontrada.")
    return db.get_indicadores(unidade_id)


@router.post("", response_model=dict)
def create_unidade(
    body: UnidadeCreate,
    current_user: CurrentUser = Depends(gestor_only),
    db: UnidadeDBManager = Depends(get_unidade_db),
):
    data = body.model_dump(exclude_none=True)
    ok, msg = db.add_unidade(data)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    criada = db.get_unidade_por_nome(data["nome"])
    return {"detail": msg, "id": criada["id"] if criada else None}


@router.put("/{unidade_id}", response_model=dict)
def update_unidade(
    unidade_id: int,
    body: UnidadeUpdate,
    current_user: CurrentUser = Depends(gestor_only),
    db: UnidadeDBManager = Depends(get_unidade_db),
):
    existente = db.get_unidade(unidade_id)
    if not existente:
        raise HTTPException(status_code=404, detail="Unidade não encontrada.")

    data = body.model_dump(exclude_none=True)
    if not data:
        raise HTTPException(status_code=400, detail="Nenhum campo para atualizar.")

    ok, msg = db.update_unidade(unidade_id, data)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


@router.delete("/{unidade_id}", response_model=dict)
def deactivate_unidade(
    unidade_id: int,
    current_user: CurrentUser = Depends(gestor_only),
    db: UnidadeDBManager = Depends(get_unidade_db),
):
    """Inativa (nunca apaga) uma unidade. A camada de dados recusa (400)
    caso ainda existam itens ativos associados."""
    existente = db.get_unidade(unidade_id)
    if not existente:
        raise HTTPException(status_code=404, detail="Unidade não encontrada.")

    ok, msg = db.deactivate_unidade(unidade_id)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


@router.post("/{unidade_id}/reativar", response_model=dict)
def reactivate_unidade(
    unidade_id: int,
    current_user: CurrentUser = Depends(gestor_only),
    db: UnidadeDBManager = Depends(get_unidade_db),
):
    """Reverte a inativação de uma unidade. Depois de reativada, o nome volta
    a aparecer em /api/constants automaticamente (a consulta lá já filtra só
    unidades ativas)."""
    existente = db.get_unidade(unidade_id)
    if not existente:
        raise HTTPException(status_code=404, detail="Unidade não encontrada.")

    ok, msg = db.reactivate_unidade(unidade_id)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}
