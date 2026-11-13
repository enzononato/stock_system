"""Cadastro de colaboradores: espelho da base de RH com complemento manual."""
import logging
from dataclasses import asdict
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query

from app.core import config as _config
from app.db.colaborador_db import (
    ColaboradorDBManager,
    CpfDuplicado,
    MSG_SYNC_EM_ANDAMENTO,
    SincronizacaoEmAndamento,
    executar_sincronizacao,
)
from app.dependencies import (
    CurrentUser,
    get_audit_context,
    get_colaborador_db,
    get_current_user,
    gestor_only,
    gestor_or_tecnico,
)
from app.schemas.colaboradores import (
    ColaboradorCpfResponse,
    ColaboradorCreate,
    ColaboradorResponse,
    ColaboradorUpdate,
    ResumoSyncResponse,
    MSG_SITUACAO_INVALIDA,
    SITUACOES,
    SyncStatus,
)
from app.schemas.common import Paginated
from app.schemas.validators import normalizar_telefone, validar_email
from app.services import audit
from app.services.audit import AuditContext
from app.services.colaboradores_source import (
    MSG_INDISPONIVEL,
    ColaboradoresSourceError,
    buscar_por_cpf,
    cpf_valido,
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/colaboradores", tags=["colaboradores"])


def _base_indisponivel(e: ColaboradoresSourceError) -> HTTPException:
    """503 para base fora do ar OU mal configurada. A causa real (sem segredos: a
    fonte nunca põe a URL de conexão na mensagem) fica só no log."""
    logger.error("Base de colaboradores indisponível: %s", e)
    return HTTPException(status_code=503, detail=MSG_INDISPONIVEL)


@router.get("", response_model=Paginated[ColaboradorResponse])
def listar_colaboradores(
    busca: Optional[str] = None,
    situacao: Optional[str] = None,
    unidade_id: Optional[int] = None,
    limit: int = Query(25, ge=1, le=200),
    offset: int = Query(0, ge=0),
    _: CurrentUser = Depends(get_current_user),
    db: ColaboradorDBManager = Depends(get_colaborador_db),
):
    if situacao is not None and situacao not in SITUACOES:
        raise HTTPException(status_code=422, detail=MSG_SITUACAO_INVALIDA)
    itens, total = db.listar(
        busca=busca, situacao=situacao, unidade_id=unidade_id, limit=limit, offset=offset
    )
    return {"items": itens, "total": total}


@router.get("/cargos", response_model=list[str])
def listar_cargos(
    _: CurrentUser = Depends(get_current_user),
    db: ColaboradorDBManager = Depends(get_colaborador_db),
):
    return db.cargos()


@router.get("/sincronizacao", response_model=SyncStatus)
def status_sincronizacao(
    _: CurrentUser = Depends(get_current_user),
    db: ColaboradorDBManager = Depends(get_colaborador_db),
):
    configurado = bool((_config.settings.COLABORADORES_DATABASE_URL or "").strip())
    return {"ultima": db.ultima_sync(), "configurado": configurado}


@router.post("/sincronizar", response_model=ResumoSyncResponse)
def sincronizar_colaboradores(
    _: CurrentUser = Depends(gestor_only),
    db: ColaboradorDBManager = Depends(get_colaborador_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        return asdict(executar_sincronizacao(db, ctx))
    except SincronizacaoEmAndamento:
        raise HTTPException(status_code=409, detail=MSG_SYNC_EM_ANDAMENTO)
    except ColaboradoresSourceError as e:
        raise _base_indisponivel(e)


@router.get("/cpf/{cpf}", response_model=ColaboradorCpfResponse)
def consultar_por_cpf(
    cpf: str,
    _: CurrentUser = Depends(get_current_user),
    db: ColaboradorDBManager = Depends(get_colaborador_db),
):
    """Cadastro local pelo CPF; se não houver, o que a base de RH tem (sem gravar)."""
    if not cpf_valido(cpf):
        raise HTTPException(status_code=422, detail="CPF inválido.")
    local = db.obter_por_cpf(cpf)
    if local is not None:
        return {**local, "local": True}
    try:
        rh = buscar_por_cpf(cpf)
    except ColaboradoresSourceError as e:
        raise _base_indisponivel(e)
    if rh is None:
        raise HTTPException(status_code=404, detail="CPF não encontrado na base de colaboradores.")
    return {
        "id": None,
        "cpf": rh.cpf,
        "nome": rh.nome,
        "matricula": rh.matricula,
        "cargo": rh.cargo,
        "setor": rh.setor,
        "unidade_rh": rh.unidade,
        "unidade_id": db.resolver_unidade_id(rh.unidade),
        "situacao": rh.situacao,
        "email": _ou_none(validar_email, rh.email),
        "telefone": _ou_none(normalizar_telefone, rh.telefone),
        "origem": "rh",
        "local": False,
    }


def _ou_none(validar, valor):
    try:
        return validar(valor)
    except ValueError:
        return None


@router.get("/{colaborador_id}", response_model=ColaboradorResponse)
def obter_colaborador(
    colaborador_id: int,
    _: CurrentUser = Depends(get_current_user),
    db: ColaboradorDBManager = Depends(get_colaborador_db),
):
    colaborador = db.obter(colaborador_id)
    if colaborador is None:
        raise HTTPException(status_code=404, detail="Colaborador não encontrado.")
    return colaborador


@router.post("", response_model=ColaboradorResponse, status_code=201)
def criar_colaborador(
    body: ColaboradorCreate,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: ColaboradorDBManager = Depends(get_colaborador_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        novo_id = db.criar_manual(body.model_dump())
    except CpfDuplicado as e:
        raise HTTPException(status_code=409, detail=str(e))
    except ValueError as e:
        # Unidade/superior inexistente é dado inválido.
        raise HTTPException(status_code=422, detail=str(e))
    criado = db.obter(novo_id)
    audit.record(ctx, "colaborador", novo_id, "criar", depois=criado)
    return criado


@router.put("/{colaborador_id}", response_model=ColaboradorResponse)
def atualizar_colaborador(
    colaborador_id: int,
    body: ColaboradorUpdate,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: ColaboradorDBManager = Depends(get_colaborador_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    antes = db.obter(colaborador_id)
    if antes is None:
        raise HTTPException(status_code=404, detail="Colaborador não encontrado.")
    dados = body.model_dump(exclude_unset=True)
    if not dados:
        raise HTTPException(status_code=400, detail="Nenhum campo para atualizar.")
    try:
        db.atualizar(colaborador_id, dados)
    except LookupError:
        raise HTTPException(status_code=404, detail="Colaborador não encontrado.")
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    audit.record_diff(
        ctx, "colaborador", colaborador_id, "editar", antes,
        obter_depois=lambda: db.obter(colaborador_id) or {},
    )
    return db.obter(colaborador_id)
