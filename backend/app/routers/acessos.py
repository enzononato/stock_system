"""Módulo de Acessos (F8): catálogo de sistemas, pacotes por cargo e ficha de acessos."""
from fastapi import APIRouter, Depends, HTTPException

from app.db.acesso_db import AcessoDBManager, AcessoErro
from app.dependencies import (
    CurrentUser,
    get_acesso_db,
    get_audit_context,
    gestor_only,
    gestor_or_tecnico,
)
from app.schemas.acessos import (
    PacoteResponse,
    PacoteUpdate,
    SistemaCreate,
    SistemaResponse,
    SistemaUpdate,
)
from app.services.audit import AuditContext

router = APIRouter(prefix="/api/acessos", tags=["acessos"])

MSG_NADA_PARA_ATUALIZAR = "Nenhum campo para atualizar."


def erro_http(e: AcessoErro) -> HTTPException:
    """Erro de domínio do módulo como `{"detail": "<mensagem>"}` com o status dele."""
    return HTTPException(status_code=e.status_code, detail=str(e))


# ── Sistemas ────────────────────────────────────────────────────────────────

@router.get("/sistemas", response_model=list[SistemaResponse])
def listar_sistemas(
    include_inactive: bool = False,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoDBManager = Depends(get_acesso_db),
):
    return db.listar_sistemas(include_inactive=include_inactive)


@router.post("/sistemas", response_model=SistemaResponse, status_code=201)
def criar_sistema(
    body: SistemaCreate,
    _: CurrentUser = Depends(gestor_only),
    db: AcessoDBManager = Depends(get_acesso_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        return db.criar_sistema(ctx, nome=body.nome, descricao=body.descricao)
    except AcessoErro as e:
        raise erro_http(e)


@router.put("/sistemas/{sistema_id}", response_model=SistemaResponse)
def atualizar_sistema(
    sistema_id: int,
    body: SistemaUpdate,
    _: CurrentUser = Depends(gestor_only),
    db: AcessoDBManager = Depends(get_acesso_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    dados = body.model_dump(exclude_unset=True)
    # `nome` e `ativo` não são anuláveis: null explícito equivale a campo não enviado.
    for campo in ("nome", "ativo"):
        if dados.get(campo, 0) is None:
            del dados[campo]
    if not dados:
        raise HTTPException(status_code=400, detail=MSG_NADA_PARA_ATUALIZAR)
    try:
        return db.atualizar_sistema(ctx, sistema_id, dados)
    except AcessoErro as e:
        raise erro_http(e)


# ── Pacotes por função ──────────────────────────────────────────────────────

@router.get("/pacotes", response_model=list[PacoteResponse])
def listar_pacotes(
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoDBManager = Depends(get_acesso_db),
):
    return db.listar_pacotes()


@router.put("/pacotes", response_model=PacoteResponse)
def definir_pacote(
    body: PacoteUpdate,
    _: CurrentUser = Depends(gestor_only),
    db: AcessoDBManager = Depends(get_acesso_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        return db.definir_pacote(ctx, cargo=body.cargo, sistema_ids=body.sistema_ids)
    except AcessoErro as e:
        raise erro_http(e)
