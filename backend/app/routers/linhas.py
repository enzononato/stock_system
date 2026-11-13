"""Linhas telefônicas: cadastro, vínculo a aparelho e colaborador, eventos e relatório de custo."""
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query

from app.db.linha_db import (
    LinhaConflito,
    LinhaDBManager,
    LinhaInvalida,
    LinhaNaoEncontrada,
)
from app.dependencies import (
    CurrentUser,
    get_audit_context,
    get_current_user,
    get_linha_db,
    gestor_or_tecnico,
)
from app.schemas.common import Paginated
from app.schemas.linhas import AparelhoOpcao, LinhaCreate, LinhaResponse
from app.services.audit import AuditContext

router = APIRouter(prefix="/api/linhas", tags=["linhas"])

MSG_NAO_ENCONTRADA = "Linha não encontrada."


def _executar(func, *args, **kwargs):
    """Chama o manager e converte as exceções de domínio no status HTTP correspondente."""
    try:
        return func(*args, **kwargs)
    except LinhaNaoEncontrada:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADA)
    except LinhaInvalida as e:
        raise HTTPException(status_code=400, detail=str(e))
    except LinhaConflito as e:
        raise HTTPException(status_code=409, detail=str(e))


# As rotas estáticas ficam sempre acima de `/{linha_id}`.

@router.get("", response_model=Paginated[LinhaResponse])
def listar_linhas(
    busca: Optional[str] = None,
    status: Optional[str] = None,
    unidade_id: Optional[int] = None,
    operadora: Optional[str] = None,
    item_id: Optional[int] = None,
    colaborador_id: Optional[int] = None,
    limit: int = Query(25, ge=1, le=100),
    offset: int = Query(0, ge=0),
    _: CurrentUser = Depends(get_current_user),
    db: LinhaDBManager = Depends(get_linha_db),
):
    itens, total = db.listar(
        busca=busca, status=status, unidade_id=unidade_id, operadora=operadora,
        item_id=item_id, colaborador_id=colaborador_id, limit=limit, offset=offset,
    )
    return {"items": itens, "total": total}


@router.get("/aparelhos", response_model=list[AparelhoOpcao])
def listar_aparelhos(
    busca: Optional[str] = None,
    linha_id: Optional[int] = None,
    limit: int = Query(20, ge=1, le=50),
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: LinhaDBManager = Depends(get_linha_db),
):
    return db.aparelhos(busca=busca, linha_id=linha_id, limit=limit)


@router.get("/{linha_id}", response_model=LinhaResponse)
def obter_linha(
    linha_id: int,
    _: CurrentUser = Depends(get_current_user),
    db: LinhaDBManager = Depends(get_linha_db),
):
    linha = db.obter(linha_id)
    if not linha:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADA)
    return linha


@router.post("", response_model=LinhaResponse, status_code=201)
def criar_linha(
    body: LinhaCreate,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: LinhaDBManager = Depends(get_linha_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    linha_id = _executar(db.criar, body.model_dump(), ctx)
    return db.obter(linha_id)
