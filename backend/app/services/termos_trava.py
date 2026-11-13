"""Trava por solicitação de termo (`GET_LOCK` do MySQL).

Serializa tudo que muda uma solicitação ou confirma o item por causa dela: consulta/webhook do
RevSign, upload manual do PDF assinado, cancelamento e envio agendado. Quem já segura a trava
chama as funções de serviço com `travado=True` (o lock é por conexão: pegá-lo de novo a partir de
outra conexão do pool falharia).
"""
import logging
from contextlib import contextmanager
from typing import Iterator

from app.core import config as _config
from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)

MSG_PROCESSANDO = "Esta solicitação está sendo processada. Tente novamente em instantes."


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
