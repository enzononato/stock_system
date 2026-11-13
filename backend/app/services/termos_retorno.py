"""Retorno do RevSign: sincronização das assinaturas (webhook e consulta periódica).

O webhook é só um gatilho: ele identifica a solicitação e chama `sincronizar`, que consulta o
RevSign (a fonte da verdade). O status que vier no corpo do webhook nunca é usado.

`sincronizar` é idempotente e serializado por solicitação (`GET_LOCK`), então o mesmo evento
repetido, ou webhook e consulta periódica ao mesmo tempo, geram uma única confirmação.
"""
import logging
from datetime import datetime
from typing import Optional

from app.core.tempo import FUSO_BRASILIA, agora_local
from app.dependencies import get_termo_db
from app.services import audit
from app.services.audit import AuditContext
from app.services.revsign_client import EventoWebhook, RevSignError, get_revsign_client
from app.services.termos import OPERADOR_REVSIGN, PREFIXO_REFERENCIA
from app.services.termos_confirmacao import MSG_JA_CONFIRMADO, concluir_assinatura
from app.services.termos_trava import nome_da_trava, trava  # noqa: F401 (reexportados)

logger = logging.getLogger(__name__)

MOTIVO_FALHA_DOWNLOAD = "Falha ao baixar o PDF assinado do RevSign; nova tentativa na próxima consulta."
_PREFIXO_FALHA_CONSULTA = "Falha ao consultar o RevSign"
_MOTIVO_PADRAO = {
    "recusado": "Recusado pelo destinatário no RevSign.",
    "expirado": "Prazo de assinatura expirado no RevSign.",
    "cancelado": "Cancelado no RevSign.",
}
_RESULTADOS_SEM_MUDANCA = {"pendente", "em_processamento", "falha_download", "falha_consulta", "ja_processado"}


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
            termo, pdf, f"revsign_{documento_id}.pdf", OPERADOR_REVSIGN, ctx, origem="revsign",
            travado=True,
        )
        if not ok and msg == MSG_JA_CONFIRMADO.get(termo["tipo"]):
            # Outra confirmação (upload manual, sem solicitação ligada) já concluiu a operação:
            # nada a fazer aqui, e nenhuma linha extra no histórico ou na auditoria.
            return "ja_processado"
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


# ── Jobs do agendador ─────────────────────────────────────────────────────────


def enviar_agendados(agora: Optional[datetime] = None) -> int:
    """Envia as solicitações `agendado` cuja data já passou. Devolve quantas saíram de `agendado`.

    Idempotente: cada solicitação é tratada dentro da trava dela e relida antes de enviar, então
    se o job rodar de novo (estado não gravado) nada é enviado em dobro. O erro de uma
    solicitação é logado e não impede as outras.
    """
    from app.services import termos  # import tardio: evita ciclo com termos_confirmacao

    if not termos.revsign_ligado():
        return 0
    enviados = 0
    for vencido in get_termo_db().agendados_vencidos(agora or agora_local()):
        try:
            with trava(vencido["id"]) as pegou:
                if not pegou:
                    continue  # outro worker está cuidando dela
                atual = get_termo_db().obter(vencido["id"])
                if atual is None or atual["status"] != "agendado":
                    continue
                termos.enviar(vencido["id"], AuditContext.sistema())
                depois = get_termo_db().obter(vencido["id"])
                if depois is not None and depois["status"] != "agendado":
                    enviados += 1
        except Exception:
            logger.exception("Falha ao enviar o termo agendado %s.", vencido["id"])
    return enviados


def consultar_enviados() -> int:
    """Consulta ao RevSign as solicitações `enviado` (reserva para o webhook perdido). Devolve
    quantas mudaram de status."""
    from app.services import termos

    if not termos.revsign_ligado():
        return 0
    mudaram = 0
    for enviado in get_termo_db().enviados():
        try:
            if sincronizar(enviado["id"]) not in _RESULTADOS_SEM_MUDANCA:
                mudaram += 1
        except Exception:
            logger.exception("Falha ao consultar o termo %s no RevSign.", enviado["id"])
    return mudaram
