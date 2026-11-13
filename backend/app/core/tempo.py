"""Horário de Brasília sem fuso, usado nas datas de `termo_solicitacoes`."""
from datetime import datetime
from zoneinfo import ZoneInfo

FUSO_BRASILIA = ZoneInfo("America/Sao_Paulo")


def agora_local() -> datetime:
    """Agora em Brasília, sem `tzinfo` (como o MySQL guarda o DATETIME)."""
    return datetime.now(FUSO_BRASILIA).replace(tzinfo=None)
