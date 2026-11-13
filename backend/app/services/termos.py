"""Solicitações de assinatura de termos: solicitar, enviar, cancelar e reenviar.

Cada entrega, devolução ou pedido avulso ("uso") vira uma linha de `termo_solicitacoes`.
Com `REVSIGN_ENABLED=true` o termo (.docx) vai ao RevSign e o colaborador assina pelo e-mail;
com a integração desligada, ou se algo falhar, a solicitação fica `manual` com o motivo e o
fluxo de sempre (baixar o termo e subir o PDF assinado) segue valendo.

A confirmação da assinatura (que muda a situação do item) está em `termos_confirmacao`; o
retorno do RevSign (webhook, consulta, jobs) em `termos_retorno`.
"""
import hashlib
import logging
import os
from dataclasses import dataclass
from datetime import datetime
from typing import Literal, Optional

from app.core import config as _config
from app.core.storage import storage
from app.core.tempo import agora_local
from app.db.termo_db import PENDENTES
from app.dependencies import get_colaborador_db, get_inventory_db, get_termo_db
from app.services import audit
from app.services.audit import AuditContext
from app.services.revsign_client import NovoDocumento, RevSignError, get_revsign_client
from app.services.termos_trava import MSG_PROCESSANDO, trava

logger = logging.getLogger(__name__)

TipoTermo = Literal["entrega", "devolucao", "uso"]

STATUS_ITEM_POR_TIPO = {"entrega": "Pendente", "uso": "Indisponível", "devolucao": "Pendente Devolução"}
TIPO_POR_STATUS_ITEM = {status: tipo for tipo, status in STATUS_ITEM_POR_TIPO.items()}
ROTULO_TIPO = {"entrega": "entrega", "devolucao": "devolução", "uso": "uso"}
PREFIXO_REFERENCIA = "patrimonio-termo-"
OPERADOR_REVSIGN = "RevSign"

MOTIVO_DESLIGADO = "Integração com o RevSign desligada: baixe o termo e anexe o PDF assinado."
MOTIVO_SEM_COLABORADOR = "Colaborador não encontrado no cadastro: assinatura manual."
MOTIVO_SEM_EMAIL = "Colaborador sem e-mail cadastrado: assinatura manual."
MOTIVO_SITUACAO_MUDOU = "O equipamento mudou de situação antes do envio."
MOTIVO_ESTORNO = "Operação estornada no histórico."
MOTIVO_SEM_CANCELAR_REMOTO = " (não foi possível cancelar no RevSign)"

_TAMANHO_MAX_MOTIVO = 500


class TermoErro(Exception):
    """Erro de regra com mensagem própria para o usuário (os routers respondem 400)."""


@dataclass(frozen=True)
class DocumentoGerado:
    conteudo: bytes
    nome_arquivo: str
    chave: str


def revsign_ligado() -> bool:
    # Lido no momento da chamada: testes trocam `app.core.config.settings`.
    return bool(_config.settings.REVSIGN_ENABLED)


def referencia_externa(termo_id: int) -> str:
    return f"{PREFIXO_REFERENCIA}{termo_id}"


def versao_modelo(caminho: Optional[str]) -> Optional[str]:
    """`"<arquivo>@<8 primeiros hex do sha256>"`, ou None se o modelo não existe."""
    if not caminho or not os.path.isfile(caminho):
        return None
    with open(caminho, "rb") as f:
        resumo = hashlib.sha256(f.read()).hexdigest()[:8]
    return f"{os.path.basename(caminho)}@{resumo}"


def caminho_modelo(tipo: TipoTermo, revenda: Optional[str]) -> Optional[str]:
    if tipo == "devolucao":
        return _config.TERMO_DEVOLUCAO_MODELOS.get(revenda or "")
    return _config.TERMO_MODELOS


def _cortar(motivo: Optional[str]) -> Optional[str]:
    return motivo[:_TAMANHO_MAX_MOTIVO] if motivo else motivo


def _resolver_colaborador(item: dict, termo: Optional[dict] = None) -> Optional[dict]:
    """Colaborador do item; sem vínculo, o da solicitação; sem nenhum, o do CPF (payload antigo)."""
    colaboradores = get_colaborador_db()
    colaborador_id = item.get("colaborador_id") or (termo or {}).get("colaborador_id")
    if colaborador_id:
        achado = colaboradores.obter(colaborador_id)
        if achado:
            return achado
    if item.get("cpf"):
        return colaboradores.obter_por_cpf(item["cpf"])
    return None


def _descricao_do_item(item: dict) -> str:
    return " ".join(p for p in (item.get("tipo"), item.get("brand"), item.get("model")) if p)


def solicitar(
    item_id: int,
    tipo: TipoTermo,
    ctx: AuditContext,
    *,
    history_id: Optional[int] = None,
    agendado_para: Optional[datetime] = None,
    documento: Optional[DocumentoGerado] = None,
) -> dict:
    """Cria a solicitação, cancela a pendente anterior do item e (sem agendamento) envia."""
    item = get_inventory_db().find(item_id)
    if not item:
        raise TermoErro("Equipamento não encontrado.")
    termos = get_termo_db()
    colaborador = _resolver_colaborador(item)
    novo_id = termos.criar(
        item_id=item_id,
        tipo=tipo,
        status="agendado" if agendado_para else "manual",
        solicitado_por=ctx.username or "sistema",
        solicitado_em=agora_local(),
        colaborador_id=colaborador["id"] if colaborador else None,
        history_id=history_id,
        agendado_para=agendado_para,
        destinatario_nome=colaborador["nome"] if colaborador else None,
        destinatario_email=colaborador.get("email") if colaborador else None,
    )
    cancelar_pendentes_do_item(
        item_id, ctx, f"Substituída pela solicitação #{novo_id}.", exceto=novo_id
    )
    audit.record(
        ctx, "termo", novo_id, "criar", None,
        {
            "tipo": tipo,
            "item_id": item_id,
            "colaborador_id": colaborador["id"] if colaborador else None,
            "history_id": history_id,
            "status": "agendado" if agendado_para else "manual",
            "agendado_para": agendado_para,
        },
    )
    if agendado_para is None:
        return enviar(novo_id, ctx, documento=documento)
    return termos.obter(novo_id)


def _gerar_documento(item_id: int, tipo: TipoTermo) -> tuple[Optional[DocumentoGerado], Optional[str]]:
    """Gera o .docx e guarda no storage. Devolve (documento, None) ou (None, mensagem de erro)."""
    inv = get_inventory_db()
    permitidos = (STATUS_ITEM_POR_TIPO[tipo],)
    if tipo == "devolucao":
        ok, resultado, nome = inv.render_return_term(item_id, status_permitidos=permitidos)
        categoria = "termos_devolucao"
    else:
        ok, resultado, nome = inv.render_loan_term(item_id, status_permitidos=permitidos)
        categoria = "terms"
    if not ok:
        return None, str(resultado)
    try:
        chave = storage.save(categoria, nome, resultado)
    except Exception:
        logger.exception("Falha ao guardar o termo gerado (item %s).", item_id)
        return None, "falha ao guardar o arquivo do termo."
    return DocumentoGerado(resultado, nome, chave), None


def enviar(termo_id: int, ctx: AuditContext, *, documento: Optional[DocumentoGerado] = None) -> dict:
    """Prepara o termo e envia ao RevSign (ou deixa `manual`). Nunca propaga erro do RevSign.

    Só age em `agendado`/`manual`. Se o item já não está na situação do tipo da solicitação,
    cancela com `MOTIVO_SITUACAO_MUDOU`.
    """
    termos = get_termo_db()
    termo = termos.obter(termo_id)
    if termo is None or termo["status"] not in ("agendado", "manual"):
        return termo
    item = get_inventory_db().find(termo["item_id"])
    tipo = termo["tipo"]
    if item is None or item["status"] != STATUS_ITEM_POR_TIPO[tipo]:
        return cancelar(termo_id, ctx, MOTIVO_SITUACAO_MUDOU, travado=True)

    colaborador = _resolver_colaborador(item, termo)
    campos = {
        "colaborador_id": colaborador["id"] if colaborador else termo["colaborador_id"],
        "destinatario_nome": colaborador["nome"] if colaborador else termo["destinatario_nome"],
        "destinatario_email": colaborador.get("email") if colaborador else termo["destinatario_email"],
        "modelo_versao": versao_modelo(caminho_modelo(tipo, item.get("revenda"))),
    }

    erro_geracao = None
    if documento is None:
        documento, erro_geracao = _gerar_documento(item["id"], tipo)
    if documento is not None:
        campos["documento_key"] = documento.chave

    email = campos["destinatario_email"]
    status, motivo, revsign_id = "manual", None, None
    if not revsign_ligado():
        motivo = MOTIVO_DESLIGADO
    elif erro_geracao:
        motivo = f"Não foi possível gerar o termo: {erro_geracao}"
    elif colaborador is None:
        motivo = MOTIVO_SEM_COLABORADOR
    elif not email:
        motivo = MOTIVO_SEM_EMAIL
    else:
        try:
            criado = get_revsign_client().criar_documento(
                NovoDocumento(
                    titulo=(
                        f"Termo de {ROTULO_TIPO[tipo]} — {_descricao_do_item(item)} "
                        f"(patrimônio #{item['id']})"
                    ),
                    arquivo=documento.conteudo,
                    nome_arquivo=documento.nome_arquivo,
                    signatario_nome=colaborador["nome"],
                    signatario_email=email,
                    referencia_externa=referencia_externa(termo_id),
                )
            )
            status, revsign_id = "enviado", criado.documento_id
        except RevSignError as e:
            motivo = f"Falha ao enviar ao RevSign ({e}): assinatura manual."

    if status == "enviado":
        campos.update(status="enviado", enviado_em=agora_local(), revsign_documento_id=revsign_id, motivo=None)
    else:
        campos.update(status="manual", motivo=_cortar(motivo))
    try:
        gravou = termos.atualizar(termo_id, somente_se_status=("agendado", "manual"), **campos)
    except Exception:
        _cancelar_remoto(revsign_id)  # o documento não pode ficar órfão no RevSign
        raise
    if not gravou:
        # Cancelada (ou enviada por outra rotina) enquanto o termo era preparado.
        _cancelar_remoto(revsign_id)
        return termos.obter(termo_id)

    audit.record(
        ctx, "termo", termo_id, "enviar",
        {"status": termo["status"]},
        {"status": campos["status"], "motivo": campos.get("motivo")},
    )
    return termos.obter(termo_id)


def _cancelar_remoto(revsign_id: Optional[str]) -> bool:
    """Cancela o documento no RevSign em melhor esforço. True se deu certo (ou nada a fazer)."""
    if not revsign_id:
        return True
    try:
        get_revsign_client().cancelar(revsign_id)
        return True
    except Exception:
        logger.warning("Não foi possível cancelar o documento %s no RevSign.", revsign_id)
        return False


def _cancelar(termo_id: int, ctx: AuditContext, motivo: str) -> dict:
    termos = get_termo_db()
    termo = termos.obter(termo_id)
    if termo is None:
        raise TermoErro("Solicitação de termo não encontrada.")
    if termo["status"] not in PENDENTES:
        raise TermoErro("Esta solicitação não está pendente.")
    if (
        termo["status"] == "enviado"
        and termo["revsign_documento_id"]
        and revsign_ligado()
        and not _cancelar_remoto(termo["revsign_documento_id"])
    ):
        motivo = f"{motivo}{MOTIVO_SEM_CANCELAR_REMOTO}"
    motivo = _cortar(motivo)
    if not termos.atualizar(
        termo_id, somente_se_status=PENDENTES,
        status="cancelado", cancelado_em=agora_local(), motivo=motivo,
    ):
        raise TermoErro("Esta solicitação não está pendente.")
    audit.record(
        ctx, "termo", termo_id, "cancelar",
        {"status": termo["status"]}, {"status": "cancelado", "motivo": motivo},
    )
    return termos.obter(termo_id)


def cancelar(termo_id: int, ctx: AuditContext, motivo: str, *, travado: bool = False) -> dict:
    """Cancela a solicitação pendente (e o documento no RevSign, em melhor esforço).

    Roda sob a mesma trava por solicitação do retorno do RevSign, para não cancelar uma
    solicitação que está sendo confirmada naquele instante. `travado=True`: o chamador já a segura.
    """
    if travado:
        return _cancelar(termo_id, ctx, motivo)
    with trava(termo_id) as pegou:
        if not pegou:
            raise TermoErro(MSG_PROCESSANDO)
        return _cancelar(termo_id, ctx, motivo)


def cancelar_pendentes_do_item(
    item_id: int, ctx: AuditContext, motivo: str, *, exceto: Optional[int] = None
) -> list[int]:
    cancelados: list[int] = []
    for termo in get_termo_db().pendentes_do_item(item_id):
        if termo["id"] == exceto:
            continue
        try:
            cancelar(termo["id"], ctx, motivo)
            cancelados.append(termo["id"])
        except TermoErro:
            continue  # já deixou de ser pendente
    return cancelados


def cancelar_por_estorno(history_id: int, ctx: AuditContext) -> list[int]:
    cancelados: list[int] = []
    for termo in get_termo_db().por_history(history_id):
        if termo["status"] not in PENDENTES:
            continue
        try:
            cancelar(termo["id"], ctx, MOTIVO_ESTORNO)
            cancelados.append(termo["id"])
        except TermoErro:
            continue
    return cancelados


def reenviar(termo_id: int, ctx: AuditContext) -> dict:
    termos = get_termo_db()
    termo = termos.obter(termo_id)
    if termo is None:
        raise TermoErro("Solicitação de termo não encontrada.")
    status = termo["status"]
    if status == "enviado":
        if not revsign_ligado():
            raise TermoErro("Integração com o RevSign desligada: use a assinatura manual.")
        try:
            get_revsign_client().reenviar(termo["revsign_documento_id"])
        except RevSignError as e:
            raise TermoErro(f"Não foi possível reenviar: {e}") from None
    elif status == "manual":
        if not revsign_ligado():
            raise TermoErro("Integração com o RevSign desligada: use a assinatura manual.")
        enviar(termo_id, ctx)
    else:
        raise TermoErro("Só é possível reenviar uma solicitação enviada ou em assinatura manual.")
    atual = termos.obter(termo_id)
    audit.record(
        ctx, "termo", termo_id, "reenviar", {"status": status}, {"status": atual["status"]},
    )
    return atual


_CAMPOS_INTERNOS = (
    "documento_key", "arquivo_assinado_key", "revsign_documento_id", "destinatario_nome", "atualizado_em",
)


def para_resposta(linha: dict) -> dict:
    """Linha de `termo_solicitacoes` no formato de `TermoResponse` (sem chaves internas)."""
    saida = {k: v for k, v in linha.items() if k not in _CAMPOS_INTERNOS}
    saida["tem_documento"] = bool(linha.get("documento_key"))
    saida["tem_assinado"] = bool(linha.get("arquivo_assinado_key"))
    saida["legado"] = bool(linha.get("legado", False))
    return saida


MOTIVO_CONFIRMACAO_ESTORNADA = "Confirmação estornada."


def reabrir_por_estorno(entrada: dict, ctx: AuditContext) -> Optional[int]:
    """Estorno de uma `Confirmação Empréstimo`/`Confirmação Devolução`: a solicitação que foi
    assinada por ela volta a `manual`, para o card voltar a oferecer "Anexar assinado".

    Acha a solicitação pelo PDF guardado na confirmação (ou, na falta dele, a última assinada
    do mesmo tipo do item). Devolve o id reaberto, ou None.
    """
    tipo = {"Confirmação Empréstimo": "entrega", "Confirmação Devolução": "devolucao"}.get(
        entrada.get("operation")
    )
    if tipo is None or not entrada.get("item_id"):
        return None
    termos = get_termo_db()
    assinadas = termos.assinadas_do_item(entrada["item_id"], tipo)
    chave = entrada.get("termo_assinado_anexo")
    alvo = next((t for t in assinadas if chave and t["arquivo_assinado_key"] == chave), None)
    alvo = alvo or (assinadas[0] if assinadas else None)
    if alvo is None:
        return None
    if not termos.atualizar(
        alvo["id"], somente_se_status=("assinado",),
        status="manual", motivo=MOTIVO_CONFIRMACAO_ESTORNADA, assinado_em=None, arquivo_assinado_key=None,
    ):
        return None
    audit.record(
        ctx, "termo", alvo["id"], "atualizar",
        {"status": "assinado"}, {"status": "manual", "motivo": MOTIVO_CONFIRMACAO_ESTORNADA},
    )
    return alvo["id"]
