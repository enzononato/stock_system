"""Transferência de equipamentos entre unidades (F4)."""
import logging
from datetime import datetime
from typing import Optional

import pymysql

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


class TransferenciaErro(Exception):
    """Regra de negócio violada na transferência; o router converte em HTTPException."""

    def __init__(self, status_code: int, detail: str):
        super().__init__(detail)
        self.status_code = status_code
        self.detail = detail


class TransferenciaDBManager:
    """Transferências entre unidades. A tabela `transferencias` vem da migração 0404."""

    def transferir(
        self, item_id: int, para_unidade_id: int, para_local_id: Optional[int], motivo: str, usuario: str
    ) -> dict:
        """Muda `items.revenda`/`local_id`, grava `transferencias` e uma linha de
        `history` (`Edição`) numa transação, com a linha do item travada."""
        agora = datetime.now()
        try:
            with get_cursor() as cur:
                cur.execute(
                    "SELECT id, status, revenda, local_id FROM items WHERE id=%s AND is_active=1 FOR UPDATE",
                    (item_id,),
                )
                item = cur.fetchone()
                if not item:
                    raise TransferenciaErro(404, "Item não encontrado.")
                if item["status"] in ("Pendente", "Pendente Devolução"):
                    raise TransferenciaErro(
                        400, "Conclua ou estorne o empréstimo ou a devolução pendente antes de transferir."
                    )
                cur.execute("SELECT nome, is_active FROM unidades WHERE id=%s", (para_unidade_id,))
                destino = cur.fetchone()
                if not destino or not destino["is_active"]:
                    raise TransferenciaErro(400, "Unidade de destino não encontrada ou inativa.")
                de_unidade = item["revenda"]
                para_unidade = destino["nome"]
                if (de_unidade or "").strip().casefold() == para_unidade.strip().casefold():
                    raise TransferenciaErro(400, "O equipamento já está nesta unidade.")
                if para_local_id is not None:
                    erro = erro_local(para_local_id, para_unidade, rotulo_unidade="de destino")
                    if erro:
                        raise TransferenciaErro(400, erro)

                cur.execute(
                    "UPDATE items SET revenda=%s, local_id=%s WHERE id=%s",
                    (para_unidade, para_local_id, item_id),
                )
                cur.execute(
                    "INSERT INTO transferencias (item_id, de_unidade, para_unidade, de_local_id, "
                    "para_local_id, motivo, transferido_por, transferido_em) "
                    "VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                    (item_id, de_unidade, para_unidade, item["local_id"], para_local_id, motivo, usuario, agora),
                )
                transferencia_id = cur.lastrowid
                cur.execute(
                    "INSERT INTO history (item_id, operador, data_operacao, operation, details, revenda) "
                    "VALUES (%s,%s,%s,'Edição',%s,%s)",
                    (
                        item_id, usuario, agora,
                        f"Transferência {de_unidade or 'sem unidade'} → {para_unidade}", para_unidade,
                    ),
                )
        except TransferenciaErro:
            raise
        except pymysql.MySQLError:
            logger.exception("Erro de banco de dados")
            raise TransferenciaErro(400, "Erro ao processar a operação. Tente novamente.")
        return {
            "id": transferencia_id,
            "de_unidade": de_unidade,
            "para_unidade": para_unidade,
            "de_local_id": item["local_id"],
            "para_local_id": para_local_id,
        }

    def listar_por_item(self, item_id: int) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT t.*, ld.nome AS de_local_nome, lp.nome AS para_local_nome "
                "FROM transferencias t "
                "LEFT JOIN locais ld ON ld.id = t.de_local_id "
                "LEFT JOIN locais lp ON lp.id = t.para_local_id "
                "WHERE t.item_id=%s ORDER BY t.transferido_em, t.id",
                (item_id,),
            )
            return cur.fetchall()
