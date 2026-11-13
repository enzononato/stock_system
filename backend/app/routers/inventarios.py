"""Inventário físico por unidade: abertura, leituras, encerramento e relatório."""
from typing import Literal, Optional

from fastapi import APIRouter, Depends, HTTPException, Query, Response

from app.db.inventario_db import (
    InventarioDBManager,
    InventarioFechado,
    InventarioJaAberto,
    InventarioNaoEncontrado,
    LeituraNaoEncontrada,
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
    LeituraCreate,
    LeituraResponse,
    LeituraUpdate,
    Resultado,
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


# ── Leituras ────────────────────────────────────────────────────────────────

def _campos_auditados(leitura: dict) -> dict:
    return {
        k: leitura.get(k)
        for k in (
            "local_id_encontrado", "usuario_conferido", "colaborador_id_encontrado",
            "observacao", "resultado",
        )
    }


@router.get("/{inventario_id}/leituras", response_model=Paginated[LeituraResponse])
def listar_leituras(
    inventario_id: int,
    resultado: Optional[Resultado] = None,
    limit: int = Query(25, ge=1, le=100),
    offset: int = Query(0, ge=0),
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
):
    if not db.obter(inventario_id):
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADO)
    itens, total = db.listar_leituras(inventario_id, resultado=resultado, limit=limit, offset=offset)
    return {"items": itens, "total": total}


@router.post("/{inventario_id}/leituras", response_model=LeituraResponse, status_code=201)
def registrar_leitura(
    inventario_id: int,
    body: LeituraCreate,
    response: Response,
    user: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        leitura, duplicada = db.registrar_leitura(
            inventario_id,
            codigo=body.codigo,
            local_id_encontrado=body.local_id_encontrado,
            colaborador_id_encontrado=body.colaborador_id_encontrado,
            usuario_conferido=body.usuario_conferido,
            observacao=body.observacao,
            por=user.username,
        )
    except InventarioNaoEncontrado:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADO)
    except InventarioFechado as e:
        raise HTTPException(status_code=409, detail=str(e))
    except ValueError as e:  # inclui CodigoInvalido
        raise HTTPException(status_code=422, detail=str(e))
    if duplicada:
        response.status_code = 200
    else:
        audit.record(
            ctx, "inventario", inventario_id, "registrar_leitura",
            depois={
                "leitura_id": leitura["id"],
                "item_id": leitura["item_id"],
                "codigo_lido": leitura["codigo_lido"],
                "resultado": leitura["resultado"],
            },
        )
    return {**leitura, "duplicada": duplicada}


@router.patch("/{inventario_id}/leituras/{leitura_id}", response_model=LeituraResponse)
def corrigir_leitura(
    inventario_id: int,
    leitura_id: int,
    body: LeituraUpdate,
    user: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        antes, depois = db.atualizar_leitura(
            inventario_id, leitura_id, body.model_dump(exclude_unset=True), user.username
        )
    except InventarioNaoEncontrado:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADO)
    except LeituraNaoEncontrada:
        raise HTTPException(status_code=404, detail="Leitura não encontrada.")
    except InventarioFechado as e:
        raise HTTPException(status_code=409, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    diff_antes, diff_depois = audit.diff(_campos_auditados(antes), _campos_auditados(depois))
    audit.record(
        ctx, "inventario", inventario_id, "editar_leitura",
        antes={**(diff_antes or {}), "leitura_id": leitura_id},
        depois={**(diff_depois or {}), "leitura_id": leitura_id},
    )
    return depois


@router.delete("/{inventario_id}/leituras/{leitura_id}", response_model=dict)
def desfazer_leitura(
    inventario_id: int,
    leitura_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        leitura = db.excluir_leitura(inventario_id, leitura_id)
    except InventarioNaoEncontrado:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADO)
    except LeituraNaoEncontrada:
        raise HTTPException(status_code=404, detail="Leitura não encontrada.")
    except InventarioFechado as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.record(
        ctx, "inventario", inventario_id, "excluir_leitura",
        antes={
            "leitura_id": leitura_id,
            "item_id": leitura["item_id"],
            "codigo_lido": leitura["codigo_lido"],
            "resultado": leitura["resultado"],
        },
    )
    return {"detail": "Leitura excluída."}
