"""Exportações Excel e PDF das telas operacionais (F7).

Cada rota recebe os mesmos filtros da tela correspondente e exige `formato` (`xlsx` ou `pdf`).
Consultas acima do limite do formato respondem 400 com a orientação de refinar os filtros.
"""
from datetime import date, datetime
from typing import Callable, Literal, Optional

from fastapi import APIRouter, Depends, HTTPException, Query, Response

from app.db.colaborador_db import ColaboradorDBManager
from app.db.inventory_manager_db import InventoryDBManager, montar_order_by_items
from app.db.linha_db import LinhaDBManager
from app.db.manutencao_db import ManutencaoDBManager
from app.db.painel_db import PainelDBManager
from app.db.unidade_db import UnidadeDBManager
from app.dependencies import (
    CurrentUser,
    get_colaborador_db,
    get_inventory_db,
    get_linha_db,
    get_manutencao_db,
    get_painel_db,
    get_unidade_db,
    gestor_only,
    gestor_or_tecnico,
)
from app.services import exportacoes_telas as telas

router = APIRouter(prefix="/api/exportacoes", tags=["exportacoes"])

FormatoQuery = Literal["xlsx", "pdf"]


def _entregar(montar: Callable[[], telas.Planilha], formato: FormatoQuery) -> Response:
    try:
        planilha = montar()
    except telas.ExportacaoGrande as e:
        raise HTTPException(status_code=400, detail=e.detail)
    return telas.gerar(planilha, formato)


@router.get("/estoque")
def exportar_estoque(
    formato: FormatoQuery = Query(...),
    tipo: Optional[str] = None,
    status: Optional[str] = None,
    revenda: Optional[str] = None,
    search: Optional[str] = None,
    colaborador_id: Optional[int] = None,
    sort: str = "id",
    direcao: str = Query(default="desc", alias="dir"),
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    colaboradores: ColaboradorDBManager = Depends(get_colaborador_db),
):
    try:
        montar_order_by_items(sort, direcao)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    colaborador = colaboradores.obter(colaborador_id) if colaborador_id is not None else None
    return _entregar(
        lambda: telas.planilha_estoque(
            inv, tipo=tipo, status=status, revenda=revenda, search=search,
            colaborador_id=colaborador_id, sort=sort, direcao=direcao, formato=formato,
            colaborador_nome=colaborador["nome"] if colaborador else None,
        ),
        formato,
    )


@router.get("/historico")
def exportar_historico(
    formato: FormatoQuery = Query(...),
    search: Optional[str] = None,
    item_id: Optional[int] = None,
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    return _entregar(
        lambda: telas.planilha_historico(inv, search=search, item_id=item_id, formato=formato), formato
    )


@router.get("/manutencoes")
def exportar_manutencoes(
    formato: FormatoQuery = Query(...),
    status: Optional[Literal[
        "aberta", "enviada", "em_reparo", "retornada", "sem_conserto", "cancelada", "em_andamento"
    ]] = None,
    item_id: Optional[int] = None,
    revenda: Optional[str] = None,
    tipo: Optional[str] = None,
    de: Optional[date] = None,
    ate: Optional[date] = None,
    busca: Optional[str] = None,
    _: CurrentUser = Depends(gestor_or_tecnico),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
):
    return _entregar(
        lambda: telas.planilha_manutencoes(
            man, status=status, item_id=item_id, revenda=revenda, tipo=tipo, de=de, ate=ate,
            busca=busca, formato=formato,
        ),
        formato,
    )


@router.get("/linhas")
def exportar_linhas(
    formato: FormatoQuery = Query(...),
    busca: Optional[str] = None,
    status: Optional[str] = None,
    unidade_id: Optional[int] = None,
    operadora: Optional[str] = None,
    _: CurrentUser = Depends(gestor_or_tecnico),
    linhas: LinhaDBManager = Depends(get_linha_db),
    unidades: UnidadeDBManager = Depends(get_unidade_db),
):
    unidade = unidades.get_unidade(unidade_id) if unidade_id is not None else None
    return _entregar(
        lambda: telas.planilha_linhas(
            linhas, busca=busca, status=status, unidade_id=unidade_id,
            unidade_nome=unidade["nome"] if unidade else None, operadora=operadora, formato=formato,
        ),
        formato,
    )


@router.get("/termos-pendentes")
def exportar_termos_pendentes(
    formato: FormatoQuery = Query(...),
    busca: Optional[str] = None,
    _: CurrentUser = Depends(gestor_or_tecnico),
    painel: PainelDBManager = Depends(get_painel_db),
):
    return _entregar(
        lambda: telas.planilha_termos_pendentes(painel, busca=busca, formato=formato), formato
    )


@router.get("/relatorio-mensal")
def exportar_relatorio_mensal(
    formato: FormatoQuery = Query(...),
    year: Optional[int] = None,
    month: Optional[int] = None,
    _: CurrentUser = Depends(gestor_only),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    agora = datetime.now()
    ano = year or agora.year
    mes = month or agora.month
    if not (1 <= mes <= 12):
        raise HTTPException(status_code=400, detail="Mês inválido.")
    return _entregar(
        lambda: telas.planilha_relatorio_mensal(inv, ano=ano, mes=mes, formato=formato), formato
    )
