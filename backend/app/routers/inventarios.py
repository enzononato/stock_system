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
    RelatorioInventario,
    Resultado,
)
from app.services import audit, exports
from app.services.audit import AuditContext
from app.services.exports import Coluna
from app.services.inventario import ROTULO_AINDA_NAO_LIDO, ROTULOS_CATEGORIA

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


# ── Encerramento e relatório ────────────────────────────────────────────────

ROTULOS_SITUACAO = {"aberto": "Aberto", "encerrado": "Encerrado", "cancelado": "Cancelado"}

COLUNAS_EXCEL = [
    Coluna("categoria_rotulo", "Resultado", 18),
    Coluna("patrimonio", "Patrimônio", 12),
    Coluna("tipo", "Tipo", 16),
    Coluna("equipamento", "Marca/Modelo", 28),
    Coluna("identificador", "Nº de série", 20),
    Coluna("unidade_item", "Unidade do equipamento", 22, maiusculas=True),
    Coluna("local_esperado", "Local cadastrado", 20),
    Coluna("local_encontrado", "Local encontrado", 20),
    Coluna("usuario_esperado", "Usuário cadastrado", 26),
    Coluna("usuario_encontrado", "Usuário encontrado", 26),
    Coluna("lido_por", "Lido por", 16),
    Coluna("lido_em", "Lido em", 17),
    Coluna("codigo_lido", "Código lido", 26),
    Coluna("detalhe", "Detalhe", 48),
    Coluna("observacao", "Observação", 30),
]

COLUNAS_PDF = [
    Coluna("categoria_rotulo", "Resultado", 16),
    Coluna("patrimonio", "Patrimônio", 11),
    Coluna("tipo", "Tipo", 14),
    Coluna("equipamento", "Marca/Modelo", 24),
    Coluna("identificador", "Nº de série", 18),
    Coluna("local_esperado", "Local cadastrado", 18),
    Coluna("local_encontrado", "Local encontrado", 18),
    Coluna("usuario_esperado", "Usuário cadastrado", 22),
    Coluna("usuario_encontrado", "Usuário encontrado", 22),
    Coluna("detalhe", "Detalhe", 40),
]


def _data_hora(valor) -> str:
    return valor.strftime("%d/%m/%Y %H:%M") if valor else "—"


@router.post("/{inventario_id}/encerrar", response_model=InventarioDetalhe)
def encerrar_inventario(
    inventario_id: int,
    user: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        db.encerrar(inventario_id, user.username)
    except InventarioNaoEncontrado:
        raise HTTPException(status_code=404, detail=MSG_NAO_ENCONTRADO)
    except InventarioFechado as e:
        raise HTTPException(status_code=409, detail=str(e))
    detalhe = _detalhe(db, inventario_id)
    audit.record(
        ctx, "inventario", inventario_id, "encerrar",
        antes={"status": "aberto"}, depois={"status": "encerrado", **detalhe["resumo"]},
    )
    return detalhe


@router.get("/{inventario_id}/relatorio", response_model=RelatorioInventario)
def relatorio_inventario(
    inventario_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
):
    detalhe = _detalhe(db, inventario_id)
    return {"inventario": detalhe, "linhas": db.linhas_relatorio(detalhe)}


@router.get("/{inventario_id}/relatorio/exportar")
def exportar_relatorio(
    inventario_id: int,
    formato: Literal["xlsx", "pdf"],
    _: CurrentUser = Depends(gestor_or_tecnico),
    db: InventarioDBManager = Depends(get_inventario_db),
):
    detalhe = _detalhe(db, inventario_id)
    aberto = detalhe["status"] == "aberto"
    resumo = detalhe["resumo"]
    linhas = []
    for linha in db.linhas_relatorio(detalhe):
        rotulo = (
            ROTULO_AINDA_NAO_LIDO
            if aberto and linha["categoria"] == "nao_encontrado"
            else ROTULOS_CATEGORIA[linha["categoria"]]
        )
        linhas.append({
            **linha,
            "categoria_rotulo": rotulo,
            "patrimonio": f"#{linha['item_id']}" if linha["item_id"] else "—",
            "equipamento": " ".join(p for p in (linha["brand"], linha["model"]) if p),
            "lido_em": _data_hora(linha["lido_em"]) if linha["lido_em"] else "",
        })
    titulo = f"Inventário #{inventario_id} — {detalhe['unidade_nome'].upper()}"
    if aberto:
        titulo += " (parcial)"
    filtros = {
        "Situação": ROTULOS_SITUACAO[detalhe["status"]],
        "Aberto em": _data_hora(detalhe["iniciado_em"]),
        "Encerrado em": _data_hora(detalhe["encerrado_em"]),
        "Resumo": (
            f"Conferidos {resumo['ok']} · Não encontrados {resumo['nao_encontrados']} · "
            f"Local errado {resumo['local_errado']} · Usuário diferente {resumo['usuario_diferente']} · "
            f"De outra unidade {resumo['outra_unidade']} · Não cadastrados {resumo['nao_cadastrado']}"
        ),
    }
    if formato == "xlsx":
        conteudo = exports.gerar_excel(titulo, COLUNAS_EXCEL, linhas, filtros)
    else:
        conteudo = exports.gerar_pdf(titulo, COLUNAS_PDF, linhas, filtros, paisagem=True)
    return exports.resposta_arquivo(conteudo, f"inventario_{inventario_id}", formato)
