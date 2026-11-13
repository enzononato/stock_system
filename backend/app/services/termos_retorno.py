"""Retorno do RevSign: sincronização das assinaturas (webhook e consulta periódica).

O webhook é só um gatilho: ele identifica a solicitação e chama `sincronizar`, que consulta o
RevSign (a fonte da verdade). O status que vier no corpo do webhook nunca é usado.

`sincronizar` é idempotente e serializado por solicitação (`GET_LOCK`), então o mesmo evento
repetido, ou webhook e consulta periódica ao mesmo tempo, geram uma única confirmação.
"""
import logging
from contextlib import contextmanager
from datetime import datetime
from typing import Iterator, Optional

from app.core import config as _config
from app.core.tempo import FUSO_BRASILIA, agora_local
from app.db.database_mysql import get_cursor
from app.dependencies import get_termo_db
from app.services import audit
from app.services.audit import AuditContext
from app.services.revsign_client import EventoWebhook, RevSignError, get_revsign_client
from app.services.termos import OPERADOR_REVSIGN, PREFIXO_REFERENCIA
from app.services.termos_confirmacao import concluir_assinatura

logger = logging.getLogger(__name__)

MOTIVO_FALHA_DOWNLOAD = "Falha ao baixar o PDF assinado do RevSign; nova tentativa na próxima consulta."
_PREFIXO_FALHA_CONSULTA = "Falha ao consultar o RevSign"
_MOTIVO_PADRAO = {
    "recusado": "Recusado pelo destinatário no RevSign.",
    "expirado": "Prazo de assinatura expirado no RevSign.",
    "cancelado": "Cancelado no RevSign.",
}
_RESULTADOS_SEM_MUDANCA = {"pendente", "em_processamento", "falha_download", "falha_consulta", "ja_processado"}


def nome_da_trava(termo_id: int) -> str:
    """Nome do `GET_LOCK` da solicitação. Vale para o servidor MySQL inteiro, então leva o
    nome do database. Limite do MySQL: 64 caracteres."""
    return f"patrimonio_termo_{termo_id}:{_config.settings.DB_NAME}"[:64]


@contextmanager
def trava(termo_id: int) -> Iterator[bool]:
    """Trava por solicitação, sem espera. Entrega True se pegou o lock (e o segura, numa
    conexão própria do pool, durante o bloco) e False se outra rotina está processando."""
    nome = nome_da_trava(termo_id)
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT GET_LOCK(%s, 0) AS ok", (nome,))
        pegou = cur.fetchone()["ok"] == 1
        try:
            yield pegou
        finally:
            if pegou:
                try:
                    cur.execute("SELECT RELEASE_LOCK(%s)", (nome,))
                except Exception:
                    logger.exception("Falha ao liberar a trava da solicitação de termo %s.", termo_id)


def _para_local(dt: datetime) -> datetime:
    return dt.astimezone(FUSO_BRASILIA).replace(tzinfo=None) if dt.tzinfo else dt


def sincronizar(termo_id: int) -> str:
    """Consulta o RevSign e aplica o resultado. Ver o módulo para os resultados possíveis."""
    with trava(termo_id) as pegou:
        if not pegou:
            return "em_processamento"
        return _sincronizar_com_trava(termo_id)


def _sincronizar_com_trava(termo_id: int) -> str:
    termos = get_termo_db()
    termo = termos.obter(termo_id)
    if termo is None or termo["status"] != "enviado" or not termo["revsign_documento_id"]:
        return "ja_processado"  # inclusive: cancelada enquanto o RevSign ainda ia assinar
    documento_id = termo["revsign_documento_id"]
    cliente = get_revsign_client()
    ctx = AuditContext.sistema()

    try:
        status = cliente.consultar(documento_id)
    except RevSignError as e:
        termos.atualizar(
            termo_id, somente_se_status=("enviado",),
            motivo=f"{_PREFIXO_FALHA_CONSULTA} ({e}); nova tentativa na próxima consulta."[:500],
        )
        return "falha_consulta"

    if status.situacao == "pendente":
        if (termo["motivo"] or "").startswith((_PREFIXO_FALHA_CONSULTA, MOTIVO_FALHA_DOWNLOAD[:20])):
            termos.atualizar(termo_id, somente_se_status=("enviado",), motivo=None)
        return "pendente"

    if status.situacao == "assinado":
        try:
            pdf = cliente.baixar_assinado(documento_id)
        except RevSignError:
            termos.atualizar(termo_id, somente_se_status=("enviado",), motivo=MOTIVO_FALHA_DOWNLOAD)
            return "falha_download"
        ok, msg, chave = concluir_assinatura(
            termo, pdf, f"revsign_{documento_id}.pdf", OPERADOR_REVSIGN, ctx, origem="revsign"
        )
        if ok:
            if status.assinado_em is not None:
                termos.atualizar(termo_id, assinado_em=_para_local(status.assinado_em))
            return "assinado"
        # Assinado no RevSign, mas o Patrimônio recusou a confirmação (ex.: o item mudou de
        # situação): o PDF fica guardado para quem for resolver à mão.
        campos = {"status": "erro", "motivo": msg[:500]}
        if chave:
            campos["arquivo_assinado_key"] = chave
        if termos.atualizar(termo_id, somente_se_status=("enviado",), **campos):
            audit.record(
                ctx, "termo", termo_id, "atualizar", {"status": "enviado"},
                {"status": "erro", "motivo": campos["motivo"]},
            )
        return "erro"

    # recusado, expirado ou cancelado no RevSign
    campos = {"status": status.situacao, "motivo": (status.motivo or _MOTIVO_PADRAO[status.situacao])[:500]}
    if status.situacao == "cancelado":
        campos["cancelado_em"] = agora_local()
    if termos.atualizar(termo_id, somente_se_status=("enviado",), **campos):
        audit.record(
            ctx, "termo", termo_id, "atualizar", {"status": "enviado"},
            {"status": status.situacao, "motivo": campos["motivo"]},
        )
    return status.situacao


def _termo_do_evento(evento: EventoWebhook) -> Optional[dict]:
    termos = get_termo_db()
    referencia = evento.referencia_externa or ""
    if referencia.startswith(PREFIXO_REFERENCIA):
        sufixo = referencia[len(PREFIXO_REFERENCIA):]
        if sufixo.isascii() and sufixo.isdigit():
            termo = termos.obter(int(sufixo))
            if termo is not None:
                bate = evento.documento_id is None or termo["revsign_documento_id"] == evento.documento_id
                return termo if bate else None
    if evento.documento_id:
        return termos.obter_por_documento(evento.documento_id)
    return None


def processar_evento(evento: EventoWebhook) -> str:
    """Acha a solicitação do evento e sincroniza. Sem correspondência: `"nao_encontrado"`."""
    termo = _termo_do_evento(evento)
    if termo is None:
        return "nao_encontrado"
    return sincronizar(termo["id"])
