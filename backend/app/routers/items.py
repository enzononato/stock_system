import os
from fastapi import APIRouter, HTTPException, Depends, UploadFile, File, Form, Query
from typing import Optional
from datetime import datetime

from app.schemas.items import ItemCreate, ItemUpdate, ItemResponse, TransferenciaCreate
from app.schemas.common import Paginated
from app.db.inventory_manager_db import InventoryDBManager
from app.db.manutencao_db import ManutencaoDBManager
from app.db.transferencia_db import erro_local, TransferenciaDBManager, TransferenciaErro
from app.schemas.validators import data_br_para_date
from app.services.anexos import (
    CATEGORIA_APAGAMENTO, CATEGORIA_NF, resposta_anexo, salvar_anexo, validar_anexo,
)
from app.services.baixa import motivo_bloqueio_baixa, validar_apagamento
from app.dependencies import (
    get_current_user, gestor_or_tecnico, gestor_only, get_inventory_db, get_audit_context, get_transferencia_db, get_manutencao_db, CurrentUser,
)
from app.services import audit
from app.services.audit import AuditContext
from app.core.storage import storage
from app.core.config import settings

router = APIRouter(prefix="/api/items", tags=["items"])

MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20MB

# Campos que a edição pode limpar: `null` explícito grava NULL; chave ausente não muda.
CAMPOS_ANULAVEIS = ("valor_compra", "data_compra", "garantia_ate", "local_id")
CAMPOS_DE_DATA = ("data_compra", "garantia_ate")
MSG_GARANTIA_ANTES_DA_COMPRA = "A garantia não pode terminar antes da data da compra."


def _safe_filename(name: str) -> str:
    return os.path.basename(name or "").replace("\\", "").replace("/", "") or "arquivo"


@router.get("", response_model=Paginated[ItemResponse])
def list_items(
    tipo: Optional[str] = None,
    status: Optional[str] = None,
    revenda: Optional[str] = None,
    search: Optional[str] = None,
    colaborador_id: Optional[int] = None,
    limit: int = Query(default=settings.DEFAULT_PAGE_SIZE, ge=1, le=settings.MAX_PAGE_SIZE),
    offset: int = Query(default=0, ge=0),
    _: CurrentUser = Depends(get_current_user),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    # T10: os filtros são repassados para a camada de dados (SQL), em vez de
    # filtrar em Python depois de carregar a tabela inteira em memória.
    items, total = inv.list_items(
        tipo=tipo, status=status, revenda=revenda, search=search,
        colaborador_id=colaborador_id, limit=limit, offset=offset,
    )
    return {"items": items, "total": total}


@router.get("/{item_id}", response_model=ItemResponse)
def get_item(
    item_id: int,
    _: CurrentUser = Depends(get_current_user),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    item = inv.get_item(item_id)
    if not item:
        raise HTTPException(status_code=404, detail="Item não encontrado.")
    return item


@router.post("", response_model=dict)
def create_item(
    body: ItemCreate,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    # O schema (ItemCreate) já validou o formato dd/mm/yyyy antes de chegar
    # aqui (T7); o strptime abaixo não deve mais estourar em uso normal.
    date_reg = datetime.strptime(body.date_registered, "%d/%m/%Y")

    item_data = body.model_dump(exclude_none=True, exclude={"date_registered"})
    item_data["date_registered"] = date_reg
    for campo in CAMPOS_DE_DATA:
        if campo in item_data:
            item_data[campo] = data_br_para_date(item_data[campo])

    compra, garantia = item_data.get("data_compra"), item_data.get("garantia_ate")
    if compra and garantia and garantia < compra:
        raise HTTPException(status_code=400, detail=MSG_GARANTIA_ANTES_DA_COMPRA)
    if body.local_id is not None:
        erro = erro_local(body.local_id, body.revenda, rotulo_unidade="do equipamento")
        if erro:
            raise HTTPException(status_code=400, detail=erro)

    ok, result = inv.add_item(item_data, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=result)
    audit.record_diff(ctx, "item", result, "criar", None, obter_depois=lambda: inv.find(result))
    return {"detail": "Item cadastrado com sucesso.", "id": result}


@router.put("/{item_id}", response_model=dict)
def update_item(
    item_id: int,
    body: ItemUpdate,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    existing = inv.find(item_id)
    if not existing:
        raise HTTPException(status_code=404, detail="Item não encontrado.")

    item_data = body.model_dump(exclude_none=True)
    for campo in CAMPOS_ANULAVEIS:
        if campo in body.model_fields_set and getattr(body, campo) is None:
            item_data[campo] = None
    if not item_data:
        raise HTTPException(status_code=400, detail="Nenhum campo para atualizar.")

    # A unidade só muda pela transferência (fica registrada e muda o local junto).
    if "revenda" in item_data and item_data["revenda"] != existing.get("revenda"):
        raise HTTPException(
            status_code=400, detail="Para mudar a unidade do equipamento, use a transferência."
        )
    for campo in CAMPOS_DE_DATA:
        if item_data.get(campo) is not None:
            item_data[campo] = data_br_para_date(item_data[campo])
    compra = item_data["data_compra"] if "data_compra" in item_data else existing.get("data_compra")
    garantia = item_data["garantia_ate"] if "garantia_ate" in item_data else existing.get("garantia_ate")
    if compra and garantia and garantia < compra:
        raise HTTPException(status_code=400, detail=MSG_GARANTIA_ANTES_DA_COMPRA)
    # Só valida o local quando muda: manter um local desativado depois não trava a edição.
    novo_local = item_data.get("local_id")
    if novo_local is not None and novo_local != existing.get("local_id"):
        erro = erro_local(novo_local, existing.get("revenda"), rotulo_unidade="do equipamento")
        if erro:
            raise HTTPException(status_code=400, detail=erro)

    ok, msg = inv.update_item(item_id, item_data, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    audit.record_diff(
        ctx, "item", item_id, "editar", existing, obter_depois=lambda: inv.find(item_id)
    )
    return {"detail": msg}


@router.post("/{item_id}/nota-fiscal", response_model=dict)
async def anexar_nota_fiscal(
    item_id: int,
    arquivo: UploadFile = File(...),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    existing = inv.find(item_id)
    if not existing:
        raise HTTPException(status_code=404, detail="Item não encontrado.")
    key = await salvar_anexo(arquivo, CATEGORIA_NF, f"nf_{item_id}")
    ok, msg = inv.update_item(item_id, {"nf_anexo_key": key}, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    audit.record_diff(
        ctx, "item", item_id, "editar",
        antes={"nf_anexo_key": existing.get("nf_anexo_key")}, depois={"nf_anexo_key": key},
    )
    return {"detail": "Nota fiscal anexada.", "nf_anexo_key": key}


@router.get("/{item_id}/nota-fiscal")
def baixar_nota_fiscal(
    item_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    item = inv.find(item_id)
    if not item:
        raise HTTPException(status_code=404, detail="Item não encontrado.")
    if not item.get("nf_anexo_key"):
        raise HTTPException(status_code=404, detail="Nenhuma nota fiscal anexada a este equipamento.")
    return resposta_anexo(item["nf_anexo_key"], f"nota_fiscal_item_{item_id}")


@router.post("/{item_id}/transferir", response_model=dict)
def transferir_item(
    item_id: int,
    body: TransferenciaCreate,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    transferencias: TransferenciaDBManager = Depends(get_transferencia_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    try:
        r = transferencias.transferir(
            item_id, body.para_unidade_id, body.para_local_id, body.motivo, current_user.username
        )
    except TransferenciaErro as e:
        raise HTTPException(status_code=e.status_code, detail=e.detail)
    audit.record(
        ctx, "item", item_id, "transferir",
        antes={"revenda": r["de_unidade"], "local_id": r["de_local_id"]},
        depois={"revenda": r["para_unidade"], "local_id": r["para_local_id"], "motivo": body.motivo},
    )
    de = r["de_unidade"] or "sem unidade"
    return {"detail": f"Equipamento transferido de {de} para {r['para_unidade']}.", "id": r["id"]}


@router.delete("/{item_id}", response_model=dict)
async def remove_item(
    item_id: int,
    reason: str = Form(...),
    attachment: Optional[UploadFile] = File(default=None),
    dados_apagados: Optional[str] = Form(default=None),
    metodo_apagamento: Optional[str] = Form(default=None),
    apagado_por: Optional[str] = Form(default=None),
    apagado_em: Optional[str] = Form(default=None),
    comprovante_apagamento: Optional[UploadFile] = File(default=None),
    current_user: CurrentUser = Depends(gestor_only),
    inv: InventoryDBManager = Depends(get_inventory_db),
    man: ManutencaoDBManager = Depends(get_manutencao_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    # Checagem de existência antes de gravar o anexo, por dois motivos: manter o
    # 404 coerente com GET e PUT do mesmo recurso (a remoção respondia 400 "ID não
    # encontrado", vindo da camada de dados), e não deixar comprovante órfão no
    # storage para um item que nunca existiu.
    existing = inv.find(item_id)
    if not existing:
        raise HTTPException(status_code=404, detail="Item não encontrado.")

    bloqueio = motivo_bloqueio_baixa(existing, man.tem_em_andamento(item_id))
    if bloqueio:
        raise HTTPException(status_code=400, detail=bloqueio)
    # Os campos do apagamento são Form (não passam pelo handler de 422): erro = 400 com a mensagem pura.
    try:
        apagamento = validar_apagamento(
            tipo=existing.get("tipo"), motivo=reason, dados_apagados=dados_apagados,
            metodo=metodo_apagamento, apagado_por=apagado_por, apagado_em=apagado_em,
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    # Todos os arquivos são lidos e validados ANTES de gravar qualquer um no storage.
    conteudo_anexo = None
    if attachment:
        conteudo_anexo = await attachment.read()
        if len(conteudo_anexo) > MAX_UPLOAD_SIZE:
            raise HTTPException(status_code=413, detail="Arquivo excede o tamanho máximo de 20MB.")
    enviar_comprovante = (
        apagamento.dados_apagados == "sim"
        and comprovante_apagamento is not None
        and bool(comprovante_apagamento.filename)
    )
    if enviar_comprovante:
        validar_anexo(await comprovante_apagamento.read())
        await comprovante_apagamento.seek(0)

    storage_key = None
    if conteudo_anexo is not None:
        storage_key = storage.save(
            "notas_remocao", f"remocao_{item_id}_{_safe_filename(attachment.filename)}", conteudo_anexo
        )
    comprovante_key = None
    if enviar_comprovante:
        comprovante_key = await salvar_anexo(
            comprovante_apagamento, CATEGORIA_APAGAMENTO, f"apagamento_{item_id}"
        )

    ok, msg = inv.remove(
        item_id, current_user.username, reason, storage_key,
        apagamento=apagamento, comprovante_key=comprovante_key,
    )
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    audit.record(
        ctx, "item", item_id, "excluir",
        antes=existing,
        depois={
            "is_active": 0,
            "motivo": reason,
            "apagamento": {
                "dados_apagados": apagamento.dados_apagados,
                "metodo": apagamento.metodo,
                "apagado_por": apagamento.apagado_por,
                "apagado_em": apagamento.apagado_em,
            },
        },
    )
    return {"detail": msg}
