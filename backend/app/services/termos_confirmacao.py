"""Confirmação da assinatura de um termo, compartilhada pelo upload manual e pelo RevSign.

Os dois caminhos guardam o PDF assinado no storage, chamam a mesma confirmação do
`InventoryDBManager` (`confirm_loan`/`confirm_return`), gravam a mesma auditoria e marcam a
solicitação como `assinado`. A diferença é só a `origem` ("manual" ou "revsign").
"""
import logging
import os
from datetime import datetime
from typing import Literal, Optional

from app.core.storage import storage
from app.core.tempo import agora_local
from app.db.termo_db import PENDENTES
from app.dependencies import get_inventory_db, get_termo_db
from app.services import audit
from app.services.audit import AuditContext
from app.services.revsign_client import get_revsign_client
from app.services.termos import TermoErro, revsign_ligado

logger = logging.getLogger(__name__)

Origem = Literal["manual", "revsign"]


def nome_seguro(nome: str) -> str:
    return os.path.basename(nome or "").replace("\\", "").replace("/", "") or "arquivo"


def auditar_etapa(
    ctx: AuditContext, entidade: str, acao: str, item_id: int,
    antes_item: Optional[dict], extra: Optional[dict] = None,
) -> None:
    inv = get_inventory_db()
    audit.record_diff(
        ctx, entidade, item_id, acao, antes_item, obter_depois=lambda: inv.find(item_id), extra=extra
    )


def _carimbo() -> str:
    return datetime.now().strftime("%Y%m%d%H%M%S")


def _termo_a_marcar(item_id: int, tipo: str, termo_id: Optional[int]) -> Optional[dict]:
    termos = get_termo_db()
    if termo_id is not None:
        return termos.obter(termo_id)
    pendentes = termos.pendentes_do_item(item_id, tipo=tipo)
    return pendentes[0] if pendentes else None


def _marcar_assinada(termo: Optional[dict], chave: str, origem: Origem, ctx: AuditContext) -> None:
    """Marca a solicitação como `assinado` (condicional a ainda estar pendente)."""
    if termo is None or termo["status"] not in PENDENTES:
        return
    marcou = get_termo_db().atualizar(
        termo["id"], somente_se_status=PENDENTES,
        status="assinado", assinado_em=agora_local(), arquivo_assinado_key=chave, motivo=None,
    )
    if not marcou:
        return
    audit.record(
        ctx, "termo", termo["id"], "assinar",
        {"status": termo["status"]},
        {"status": "assinado", "origem": origem, "tipo": termo["tipo"]},
    )
    if (
        termo["status"] == "enviado"
        and origem == "manual"
        and termo["revsign_documento_id"]
        and revsign_ligado()
    ):
        # Assinado à mão: o pedido que seguia aberto no RevSign perde o sentido.
        try:
            get_revsign_client().cancelar(termo["revsign_documento_id"])
        except Exception:
            logger.warning(
                "Não foi possível cancelar o documento %s no RevSign.", termo["revsign_documento_id"]
            )


def confirmar_emprestimo(
    item_id: int, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext,
    *, termo_id: Optional[int] = None, origem: Origem = "manual",
) -> tuple[bool, str, str]:
    inv = get_inventory_db()
    chave = storage.save(
        "termos_assinados", f"termo_{item_id}_{_carimbo()}_{nome_seguro(nome_arquivo)}", conteudo
    )
    antes_item = inv.find(item_id)
    ok, msg = inv.confirm_loan(item_id, operador, chave)
    if not ok:
        return False, msg, chave
    auditar_etapa(ctx, "emprestimo", "confirmar", item_id, antes_item, extra={"termo_assinado": chave})
    _marcar_assinada(_termo_a_marcar(item_id, "entrega", termo_id), chave, origem, ctx)
    return True, msg, chave


def confirmar_devolucao(
    item_id: int, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext,
    *, termo_id: Optional[int] = None, origem: Origem = "manual",
) -> tuple[bool, str, str]:
    inv = get_inventory_db()
    chave = storage.save(
        "termos_devolucao_assinados",
        f"termo_devolucao_{item_id}_{_carimbo()}_{nome_seguro(nome_arquivo)}",
        conteudo,
    )
    antes_item = inv.find(item_id)
    ok, msg = inv.confirm_return(item_id, operador, chave)
    if not ok:
        return False, msg, chave
    auditar_etapa(ctx, "devolucao", "confirmar", item_id, antes_item, extra={"termo_assinado": chave})
    _marcar_assinada(_termo_a_marcar(item_id, "devolucao", termo_id), chave, origem, ctx)
    return True, msg, chave


def registrar_uso_assinado(
    termo_id: int, conteudo: bytes, nome_arquivo: str, ctx: AuditContext,
    *, origem: Origem = "manual",
) -> tuple[bool, str, str]:
    """Termo de uso: só guarda o PDF e marca a solicitação; o item não muda de situação."""
    termo = get_termo_db().obter(termo_id)
    if termo is None:
        return False, "Solicitação de termo não encontrada.", ""
    chave = storage.save(
        "termos_assinados",
        f"termo_uso_{termo['item_id']}_{_carimbo()}_{nome_seguro(nome_arquivo)}",
        conteudo,
    )
    _marcar_assinada(termo, chave, origem, ctx)
    return True, "Termo de uso assinado registrado.", chave


def concluir_assinatura(
    termo: dict, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext,
    *, origem: Origem,
) -> tuple[bool, str, str]:
    tipo = termo["tipo"]
    if tipo == "entrega":
        return confirmar_emprestimo(
            termo["item_id"], conteudo, nome_arquivo, operador, ctx, termo_id=termo["id"], origem=origem
        )
    if tipo == "devolucao":
        return confirmar_devolucao(
            termo["item_id"], conteudo, nome_arquivo, operador, ctx, termo_id=termo["id"], origem=origem
        )
    return registrar_uso_assinado(termo["id"], conteudo, nome_arquivo, ctx, origem=origem)


def anexar_assinado(
    termo_id: int, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext
) -> dict:
    """Upload manual do PDF assinado de uma solicitação.

    Serializado com o retorno do RevSign pela mesma trava por solicitação: assinatura manual e
    webhook ao mesmo tempo não confirmam o empréstimo duas vezes.
    """
    from app.services.termos_retorno import trava  # import tardio: termos_retorno importa este módulo

    termos = get_termo_db()
    with trava(termo_id) as pegou:
        if not pegou:
            raise TermoErro("Esta solicitação está sendo processada. Tente novamente em instantes.")
        termo = termos.obter(termo_id)
        if termo is None:
            raise TermoErro("Solicitação de termo não encontrada.")
        if termo["status"] not in ("enviado", "manual", "erro"):
            raise TermoErro("Esta solicitação não aceita anexo.")
        ok, msg, _ = concluir_assinatura(termo, conteudo, nome_arquivo, operador, ctx, origem="manual")
        if not ok:
            raise TermoErro(msg)
        return termos.obter(termo_id)
