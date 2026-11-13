"""Transferência de equipamentos entre unidades (F4)."""
import logging
from typing import Optional

from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)


def erro_local(local_id: int, unidade_nome: Optional[str], *, rotulo_unidade: str) -> Optional[str]:
    """Motivo pelo qual `local_id` não serve para um equipamento da unidade
    `unidade_nome`, ou `None` se serve. `rotulo_unidade` entra na mensagem
    ("do equipamento", "de destino")."""
    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT l.ativo, u.nome FROM locais l JOIN unidades u ON u.id = l.unidade_id WHERE l.id=%s",
            (local_id,),
        )
        linha = cur.fetchone()
    if not linha:
        return "Local não encontrado."
    if not linha["ativo"]:
        return "Este local está inativo."
    if (linha["nome"] or "").strip().casefold() != (unidade_nome or "").strip().casefold():
        return f"O local informado não pertence à unidade {rotulo_unidade}."
    return None
