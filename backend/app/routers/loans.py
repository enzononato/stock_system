import logging
import re
from typing import Optional

from fastapi import APIRouter, Body, HTTPException, Depends, UploadFile, File

from app.schemas.loans import LoanRequest, ReturnInitiateRequest
from app.db.colaborador_db import ColaboradorDBManager
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import (
    gestor_or_tecnico, get_colaborador_db, get_inventory_db, get_audit_context, CurrentUser,
)
from app.services import audit, termos
from app.services.audit import AuditContext
from app.services.termos import DocumentoGerado
from app.core.storage import storage
from app.services.termos_confirmacao import MSG_JA_CONFIRMADO, confirmar_devolucao, confirmar_emprestimo
from app.services.termos_trava import MSG_PROCESSANDO

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/loans", tags=["loans"])

MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20MB

MSG_SEM_EMAIL = "O colaborador precisa ter e-mail cadastrado para receber o termo."
MSG_NAO_ENCONTRADO = "Colaborador não encontrado."
MSG_CPF_NAO_CONFERE = "O CPF do colaborador escolhido não confere com o do responsável pelo equipamento."


def _check_size(content: bytes) -> None:
    if len(content) > MAX_UPLOAD_SIZE:
        raise HTTPException(status_code=413, detail="Arquivo excede o tamanho máximo de 20MB.")


def _status_da_confirmacao(msg: str) -> int:
    """409 quando outra confirmação ganhou a corrida (ou está em andamento); 400 nos demais casos."""
    return 409 if msg in (*MSG_JA_CONFIRMADO.values(), MSG_PROCESSANDO) else 400


def _solicitar_termo(item_id: int, tipo: str, ctx: AuditContext, **kwargs) -> Optional[dict]:
    """Cria a solicitação de assinatura sem nunca derrubar a operação principal (já gravada)."""
    try:
        return termos.para_resposta(termos.solicitar(item_id, tipo, ctx, **kwargs))
    except Exception:
        logger.exception("Falha ao criar a solicitação de termo (item %s, tipo %s).", item_id, tipo)
        return None


@router.post("", response_model=dict)
def initiate_loan(
    body: LoanRequest,
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    colaboradores: ColaboradorDBManager = Depends(get_colaborador_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    usuario, cpf, cargo, setor = body.usuario, body.cpf, body.cargo, body.setor
    if body.colaborador_id is not None:
        colaborador = colaboradores.obter(body.colaborador_id)
        if not colaborador:
            raise HTTPException(status_code=400, detail=MSG_NAO_ENCONTRADO)
        if colaborador["situacao"] == "desligado":
            raise HTTPException(status_code=400, detail="Colaborador desligado não pode receber equipamento.")
        if not colaborador.get("email"):
            raise HTTPException(status_code=400, detail=MSG_SEM_EMAIL)
        # Os dados do cadastro valem sobre o corpo; cargo/setor só complementam o que o cadastro não tem.
        usuario, cpf = colaborador["nome"], colaborador["cpf"]
        cargo = colaborador.get("cargo") or body.cargo or ""
        setor = colaborador.get("setor") or body.setor or ""

    antes_item = inv.find(body.item_id)
    ok, msg = inv.issue(
        body.item_id,
        usuario,
        cpf,
        body.center_cost,
        cargo,
        setor,
        body.revenda,
        body.date_issue,
        current_user.username,
        pessoa_juridica=body.pessoa_juridica,
        colaborador_id=body.colaborador_id,
    )
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    audit.record_diff(
        ctx, "emprestimo", body.item_id, "criar", antes_item, obter_depois=lambda: inv.find(body.item_id),
        extra={"cargo": cargo, "center_cost": body.center_cost, "setor": setor},
    )
    termo = _solicitar_termo(
        body.item_id, "entrega", ctx, history_id=inv.ultimo_history_id(body.item_id, "Empréstimo")
    )
    return {"detail": msg, "termo": termo}


@router.post("/{item_id}/confirm", response_model=dict)
async def confirm_loan(
    item_id: int,
    signed_pdf: UploadFile = File(...),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    ctx: AuditContext = Depends(get_audit_context),
):
    content = await signed_pdf.read()
    _check_size(content)
    ok, msg, _ = confirmar_emprestimo(item_id, content, signed_pdf.filename, current_user.username, ctx)
    if not ok:
        raise HTTPException(status_code=_status_da_confirmacao(msg), detail=msg)
    return {"detail": msg}


@router.get("/{item_id}/signed-term-url", response_model=dict)
def get_signed_term_url(
    item_id: int,
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    """Alimenta o botão "Ver Termo" da tela de Termos: devolve a URL do PDF
    assinado do empréstimo ativo do item, sem expor a chave de storage crua."""
    item = inv.find(item_id)
    if not item:
        raise HTTPException(status_code=404, detail="Item não encontrado.")
    if item["status"] != "Indisponível":
        raise HTTPException(status_code=400, detail="Item não está com empréstimo ativo.")
    key = inv.get_active_signed_term_key(item_id)
    if not key or not storage.exists(key):
        raise HTTPException(status_code=404, detail="Termo assinado não encontrado.")
    return {"url": storage.get_url(key)}


@router.post("/{item_id}/return/initiate", response_model=dict)
def initiate_return(
    item_id: int,
    body: Optional[ReturnInitiateRequest] = Body(None),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
    colaboradores: ColaboradorDBManager = Depends(get_colaborador_db),
    ctx: AuditContext = Depends(get_audit_context),
):
    colaborador_id = body.colaborador_id if body else None
    vincular = False
    if colaborador_id is not None:
        # Colaborador desligado pode devolver; só precisa existir e ter e-mail para o termo.
        colaborador = colaboradores.obter(colaborador_id)
        if not colaborador:
            raise HTTPException(status_code=400, detail=MSG_NAO_ENCONTRADO)
        if not colaborador.get("email"):
            raise HTTPException(status_code=400, detail=MSG_SEM_EMAIL)
        vinculado = (inv.find(item_id) or {}).get("colaborador_id")
        if vinculado is not None and vinculado != colaborador_id:
            raise HTTPException(status_code=400, detail="O equipamento está vinculado a outro colaborador.")
        vincular = vinculado is None
        # Equipamento antigo (sem vínculo de cadastro): o termo sai com o nome e o CPF registrados
        # nele, então quem o recebe precisa ser essa pessoa. Com vínculo, já é o colaborador dele.
        if vincular:
            cpf_do_item = re.sub(r"[^0-9]", "", (inv.find(item_id) or {}).get("cpf") or "")
            if cpf_do_item and cpf_do_item != re.sub(r"[^0-9]", "", colaborador.get("cpf") or ""):
                raise HTTPException(status_code=422, detail=MSG_CPF_NAO_CONFERE)

    antes_item = inv.find(item_id)
    ok, result, filename = inv.generate_return_term_bytes(item_id, current_user.username)
    if not ok:
        raise HTTPException(status_code=400, detail=result)
    audit.record_diff(
        ctx, "devolucao", item_id, "criar", antes_item, obter_depois=lambda: inv.find(item_id)
    )
    if vincular:
        inv.vincular_colaborador(item_id, colaborador_id)

    # Salva o DOCX gerado no storage
    storage_key = storage.save("termos_devolucao", filename, result)
    download_url = storage.get_url(storage_key)

    termo = _solicitar_termo(
        item_id, "devolucao", ctx,
        history_id=inv.ultimo_history_id(item_id, "Devolução"),
        documento=DocumentoGerado(result, filename, storage_key),
    )
    return {
        "detail": "Termo de devolução gerado.",
        "download_url": download_url,
        "filename": filename,
        "termo": termo,
    }


@router.post("/{item_id}/return/confirm", response_model=dict)
async def confirm_return(
    item_id: int,
    signed_pdf: UploadFile = File(...),
    current_user: CurrentUser = Depends(gestor_or_tecnico),
    ctx: AuditContext = Depends(get_audit_context),
):
    content = await signed_pdf.read()
    _check_size(content)
    ok, msg, _ = confirmar_devolucao(item_id, content, signed_pdf.filename, current_user.username, ctx)
    if not ok:
        raise HTTPException(status_code=_status_da_confirmacao(msg), detail=msg)
    return {"detail": msg}
