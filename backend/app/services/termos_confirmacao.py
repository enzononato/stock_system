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
from app.db.inventory_manager_db import MSG_DEVOLUCAO_JA_CONFIRMADA, MSG_EMPRESTIMO_JA_CONFIRMADO
from app.db.termo_db import PENDENTES
from app.dependencies import get_inventory_db, get_termo_db
from app.services import audit
from app.services.audit import AuditContext
from app.services.revsign_client import get_revsign_client
from app.services.termos import TermoErro, revsign_ligado
from app.services.termos_trava import MSG_PROCESSANDO, trava

logger = logging.getLogger(__name__)

Origem = Literal["manual", "revsign"]

# Resposta de `confirmar_*` quando outra confirmação já concluiu a operação, por tipo de termo.
MSG_JA_CONFIRMADO = {
    "entrega": MSG_EMPRESTIMO_JA_CONFIRMADO,
    "devolucao": MSG_DEVOLUCAO_JA_CONFIRMADA,
}


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


_CONFIRMACAO = {
    "entrega": {
        "categoria": "termos_assinados",
        "prefixo": "termo_",
        "entidade": "emprestimo",
        "confirmar": "confirm_loan",
    },
    "devolucao": {
        "categoria": "termos_devolucao_assinados",
        "prefixo": "termo_devolucao_",
        "entidade": "devolucao",
        "confirmar": "confirm_return",
    },
}


def _confirmar(
    tipo: str, item_id: int, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext,
    termo_id: Optional[int], origem: Origem, travado: bool,
) -> tuple[bool, str, str]:
    cfg = _CONFIRMACAO[tipo]

    def executar() -> tuple[bool, str, str]:
        inv = get_inventory_db()
        chave = storage.save(
            cfg["categoria"], f"{cfg['prefixo']}{item_id}_{_carimbo()}_{nome_seguro(nome_arquivo)}", conteudo
        )
        antes_item = inv.find(item_id)
        ok, msg = getattr(inv, cfg["confirmar"])(item_id, operador, chave)
        if not ok:
            return False, msg, chave
        auditar_etapa(ctx, cfg["entidade"], "confirmar", item_id, antes_item, extra={"termo_assinado": chave})
        _marcar_assinada(_termo_a_marcar(item_id, tipo, termo_id), chave, origem, ctx)
        return True, msg, chave

    if travado:
        return executar()
    # Chamada de fora de quem já segura a trava (upload manual direto): com solicitação pendente,
    # serializa com o RevSign e com o anexo pela solicitação. Sem solicitação, quem garante uma
    # única confirmação é a transição atômica do item (`confirm_loan`/`confirm_return`).
    alvo = _termo_a_marcar(item_id, tipo, termo_id)
    if alvo is None:
        return executar()
    with trava(alvo["id"]) as pegou:
        if not pegou:
            return False, MSG_PROCESSANDO, ""
        return executar()


def confirmar_emprestimo(
    item_id: int, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext,
    *, termo_id: Optional[int] = None, origem: Origem = "manual", travado: bool = False,
) -> tuple[bool, str, str]:
    """`travado=True`: o chamador já segura a trava da solicitação (ver `termos_trava`)."""
    return _confirmar("entrega", item_id, conteudo, nome_arquivo, operador, ctx, termo_id, origem, travado)


def confirmar_devolucao(
    item_id: int, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext,
    *, termo_id: Optional[int] = None, origem: Origem = "manual", travado: bool = False,
) -> tuple[bool, str, str]:
    return _confirmar("devolucao", item_id, conteudo, nome_arquivo, operador, ctx, termo_id, origem, travado)


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
    *, origem: Origem, travado: bool = False,
) -> tuple[bool, str, str]:
    tipo = termo["tipo"]
    if tipo == "entrega":
        return confirmar_emprestimo(
            termo["item_id"], conteudo, nome_arquivo, operador, ctx,
            termo_id=termo["id"], origem=origem, travado=travado,
        )
    if tipo == "devolucao":
        return confirmar_devolucao(
            termo["item_id"], conteudo, nome_arquivo, operador, ctx,
            termo_id=termo["id"], origem=origem, travado=travado,
        )
    return registrar_uso_assinado(termo["id"], conteudo, nome_arquivo, ctx, origem=origem)


def anexar_assinado(
    termo_id: int, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext
) -> dict:
    """Upload manual do PDF assinado de uma solicitação.

    Serializado com o retorno do RevSign pela mesma trava por solicitação: assinatura manual e
    webhook ao mesmo tempo não confirmam o empréstimo duas vezes.
    """
    termos = get_termo_db()
    with trava(termo_id) as pegou:
        if not pegou:
            raise TermoErro(MSG_PROCESSANDO)
        termo = termos.obter(termo_id)
        if termo is None:
            raise TermoErro("Solicitação de termo não encontrada.")
        if termo["status"] not in ("enviado", "manual", "erro"):
            raise TermoErro("Esta solicitação não aceita anexo.")
        ok, msg, _ = concluir_assinatura(
            termo, conteudo, nome_arquivo, operador, ctx, origem="manual", travado=True
        )
        if not ok:
            raise TermoErro(msg)
        return termos.obter(termo_id)
