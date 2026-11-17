"""Locais físicos de cada unidade: seletor de local do item e cadastro em Configurações."""
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException

from app.db.local_db import LocalDBManager
from app.db.unidade_db import UnidadeDBManager
from app.dependencies import (
    CurrentUser,
    get_audit_context,
    get_current_user,
    get_local_db,
    get_unidade_db,
    gestor_or_tecnico,
)
from app.schemas.locais import LocalCreate, LocalResponse, LocalUpdate
from app.services import audit
from app.services.audit import AuditContext

router = APIRouter(prefix="/api/locais", tags=["locais"])

MSG_NAO_ENCONTRADO = "Local não encontrado."


@router.get("", response_model=list[LocalResponse])
def listar_locais(
    unidade_id: Optional[int] = None,
    include_inactive: bool = False,
    _: CurrentUser = Depends(get_current_user),
    db: LocalDBManager = Depends(get_local_db),
):
    return db.listar(unidade_id=unidade_id, include_inactive=include_inactive)


@router.get("/{local_id}", response_model=LocalResponse)
def obter_local(
    local_id: int,
    _: CurrentUser = Depends(get_current_user),
    db: LocalDBManager = Depends(get_local_db),
):
    local = db.obter(local_id)
    if not local:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADO)
    return local


@router.post("", response_model=LocalResponse, status_code=201)
def criar_local(
    body: LocalCreate,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: LocalDBManager = Depends(get_local_db),
    unidades: UnidadeDBManager = Depends(get_unidade_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    if not unidades.get_unidade(body.unidade_id):
        raise HTTPException(status_code=404, detail="Unidade não encontrada.")
    try:
        local_id = db.criar(body.unidade_id, body.nome, body.descricao)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    criado = db.obter(local_id)
    audit.record(ctx, "local", local_id, "criar", depois=criado)
    return criado


@router.put("/{local_id}", response_model=dict)
def atualizar_local(
    local_id: int,
    body: LocalUpdate,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: LocalDBManager = Depends(get_local_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    existente = db.obter(local_id)
    if not existente:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADO)

    # `exclude_unset`: campo ausente não muda; `descricao: null` limpa. `nome` e
    # `ativo` não são anuláveis, então um null explícito neles equivale a ausente.
    dados = body.model_dump(exclude_unset=True)
    for obrigatorio in ("nome", "ativo"):
        if obrigatorio in dados and dados[obrigatorio] is None:
            del dados[obrigatorio]
    if not dados:
        raise HTTPException(status_code=400, detail="Nenhum campo para atualizar.")
    if "ativo" in dados:
        dados["ativo"] = int(dados["ativo"])

    try:
        db.atualizar(local_id, dados)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    audit.record_diff(
        ctx, "local", local_id, "editar", existente,
        obter_depois=lambda: db.obter(local_id) or {},
    )
    return {"detail": "Local atualizado com sucesso."}
