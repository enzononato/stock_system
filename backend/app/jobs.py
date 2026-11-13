"""Tarefas agendadas da aplicação. As fases seguintes registram seus jobs em `registrar_jobs`."""
import logging

from app.core import config as _config
from app.db.colaborador_db import (
    ColaboradorDBManager,
    SincronizacaoEmAndamento,
    executar_sincronizacao,
)
from app.services import scheduler
from app.services.audit import AuditContext
from app.services.colaboradores_source import ColaboradoresSourceError

logger = logging.getLogger(__name__)


def sincronizar_colaboradores() -> None:
    """Sincroniza o cadastro de colaboradores com a base de RH."""
    try:
        resumo = executar_sincronizacao(ColaboradorDBManager(), AuditContext.sistema())
    except SincronizacaoEmAndamento:
        logger.info("Sincronização de colaboradores já em andamento; tarefa agendada ignorada.")
        return
    except ColaboradoresSourceError as e:
        # Base de RH indisponível ou mal configurada: já está no log de sincronização.
        raise scheduler.JobFalhou(str(e)) from None
    logger.info("Sincronização agendada de colaboradores concluída: %s", resumo)


def registrar_jobs() -> None:
    """Registra os jobs da aplicação. Chamada uma vez no startup."""
    cfg = _config.settings
    if cfg.COLABORADORES_DATABASE_URL:
        scheduler.registrar_job(
            "sincronizar_colaboradores",
            sincronizar_colaboradores,
            diariamente_as=cfg.COLABORADORES_SYNC_HORA,
        )
