"""Módulo de Acessos (F8): catálogo de sistemas, pacotes por cargo e ficha de acessos."""
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException

from app.db.acesso_db import AcessoDBManager, AcessoErro, obter_colaborador
from app.db.database_mysql import get_cursor
from app.dependencies import (
    CurrentUser,
    get_acesso_db,
    get_audit_context,
    gestor_only,
    gestor_or_tecnico,
)
from app.schemas.acessos import (
    AcessoCreate,
    AcessoResponse,
    AcessoUpdate,
    EquipamentoResumo,
    PacoteResponse,
    PacoteUpdate,
    SistemaCreate,
    SistemaResponse,
    SistemaUpdate,
    StatusAcesso,
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


# ── Ficha de acessos do colaborador ─────────────────────────────────────────

@router.get("/ficha", response_model=list[AcessoResponse])
def listar_ficha(
    colaborador_id: Optional[int] = None,
    item_id: Optional[int] = None,
    status: Optional[StatusAcesso] = None,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoDBManager = Depends(get_acesso_db),
):
    if colaborador_id is None and item_id is None:
        raise HTTPException(status_code=422, detail="Informe colaborador_id ou item_id.")
    return db.listar_acessos(colaborador_id=colaborador_id, item_id=item_id, status=status)


@router.post("/ficha", response_model=AcessoResponse, status_code=201)
def registrar_acesso(
    body: AcessoCreate,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoDBManager = Depends(get_acesso_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        return db.criar_acesso(
            ctx,
            colaborador_id=body.colaborador_id,
            sistema_id=body.sistema_id,
            login=body.login,
            item_id=body.item_id,
        )
    except AcessoErro as e:
        raise erro_http(e)


@router.put("/ficha/{acesso_id}", response_model=AcessoResponse)
def atualizar_acesso(
    acesso_id: int,
    body: AcessoUpdate,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoDBManager = Depends(get_acesso_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    dados = body.model_dump(exclude_unset=True)
    # `login` não é anulável: null (ou só espaços, que o schema reduz a null) não muda nada.
    # Só `item_id: null` é um pedido válido (limpa o equipamento).
    if "login" in dados and dados["login"] is None:
        del dados["login"]
    if not dados:
        raise HTTPException(status_code=400, detail=MSG_NADA_PARA_ATUALIZAR)
    try:
        return db.atualizar_acesso(ctx, acesso_id, dados)
    except AcessoErro as e:
        raise erro_http(e)


@router.post("/ficha/{acesso_id}/desativar", response_model=AcessoResponse)
def desativar_acesso(
    acesso_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoDBManager = Depends(get_acesso_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        return db.desativar_acesso(ctx, acesso_id)
    except AcessoErro as e:
        raise erro_http(e)


@router.get(
    "/colaboradores/{colaborador_id}/equipamentos", response_model=list[EquipamentoResumo]
)
def equipamentos_do_colaborador(
    colaborador_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: AcessoDBManager = Depends(get_acesso_db),
):
    with get_cursor(commit=False) as cur:
        if obter_colaborador(cur, colaborador_id) is None:
            raise HTTPException(status_code=404, detail="Colaborador não encontrado.")
        return db.equipamentos_com_colaborador(colaborador_id, cur=cur)
