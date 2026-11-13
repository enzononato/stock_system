"""
Camada de acesso ao inventário físico por unidade.

As tabelas (`inventarios`, `inventario_leituras`, `inventario_nao_encontrados`) vêm
da migração 0600; este módulo não cria schema. Leitura, correção, encerramento e
cancelamento travam a linha do inventário com `SELECT ... FOR UPDATE`, o que serializa
técnicos que leem ao mesmo tempo e impede escrita depois do encerramento.

Itens de uma unidade = `items.is_active = 1 AND items.revenda = unidades.nome` (o
sistema liga item e unidade pelo nome).
"""
import logging
from datetime import datetime
from typing import Optional

import pymysql

from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)


class InventarioNaoEncontrado(Exception):
    pass


class InventarioFechado(Exception):
    pass


class InventarioJaAberto(Exception):
    def __init__(self, inventario_id: int):
        super().__init__(f"Já existe um inventário aberto para esta unidade (#{inventario_id}).")
        self.inventario_id = inventario_id


class UnidadeNaoEncontrada(Exception):
    pass


class LeituraNaoEncontrada(Exception):
    pass


def _erro_de_duplicado(e: pymysql.MySQLError) -> bool:
    return bool(e.args) and e.args[0] == 1062


_SELECT_INVENTARIO = (
    "SELECT i.*, u.nome AS unidade_nome, "
    "(SELECT COUNT(*) FROM inventario_leituras l WHERE l.inventario_id = i.id) AS total_leituras "
    "FROM inventarios i JOIN unidades u ON u.id = i.unidade_id"
)

_RESULTADOS = ("ok", "local_errado", "usuario_diferente", "outra_unidade", "nao_cadastrado")


class InventarioDBManager:
    def _travar(self, cur, inventario_id: int) -> dict:
        cur.execute(
            "SELECT i.*, u.nome AS unidade_nome FROM inventarios i "
            "JOIN unidades u ON u.id = i.unidade_id WHERE i.id = %s FOR UPDATE",
            (inventario_id,),
        )
        inventario = cur.fetchone()
        if not inventario:
            raise InventarioNaoEncontrado()
        return inventario

    def abrir(self, unidade_id: int, observacao: Optional[str], por: str) -> int:
        try:
            with get_cursor() as cur:
                cur.execute("SELECT id, nome, is_active FROM unidades WHERE id = %s", (unidade_id,))
                unidade = cur.fetchone()
                if not unidade:
                    raise UnidadeNaoEncontrada()
                if not unidade["is_active"]:
                    raise ValueError("Unidade inativa não pode ser inventariada.")
                cur.execute(
                    "INSERT INTO inventarios (unidade_id, status, iniciado_por, iniciado_em, observacao) "
                    "VALUES (%s, 'aberto', %s, %s, %s)",
                    (unidade_id, por, datetime.now(), observacao),
                )
                return cur.lastrowid
        except pymysql.MySQLError as e:
            if not _erro_de_duplicado(e):
                raise
        # Outro inventário aberto na mesma unidade (índice único da coluna gerada).
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT id FROM inventarios WHERE unidade_id = %s AND status = 'aberto'", (unidade_id,)
            )
            aberto = cur.fetchone()
        raise InventarioJaAberto(aberto["id"] if aberto else 0)

    def obter(self, inventario_id: int) -> Optional[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(_SELECT_INVENTARIO + " WHERE i.id = %s", (inventario_id,))
            return cur.fetchone()

    def listar(
        self, *, unidade_id: Optional[int], status: Optional[str], limit: int, offset: int
    ) -> tuple[list[dict], int]:
        where, params = ["1 = 1"], []
        if unidade_id is not None:
            where.append("i.unidade_id = %s")
            params.append(unidade_id)
        if status:
            where.append("i.status = %s")
            params.append(status)
        where_sql = " AND ".join(where)
        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT COUNT(*) AS total FROM inventarios i WHERE {where_sql}", params)
            total = cur.fetchone()["total"]
            cur.execute(
                f"{_SELECT_INVENTARIO} WHERE {where_sql} "
                "ORDER BY i.iniciado_em DESC, i.id DESC LIMIT %s OFFSET %s",
                [*params, limit, offset],
            )
            return cur.fetchall(), total

    def resumo(self, inventario: dict) -> dict:
        inventario_id = inventario["id"]
        status = inventario["status"]
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT resultado, COUNT(*) AS n FROM inventario_leituras "
                "WHERE inventario_id = %s GROUP BY resultado",
                (inventario_id,),
            )
            contagem = {r["resultado"]: r["n"] for r in cur.fetchall()}
            resumo = {r: int(contagem.get(r, 0)) for r in _RESULTADOS}
            resumo["lidos"] = sum(resumo.values())
            resumo["encontrados"] = resumo["ok"] + resumo["local_errado"] + resumo["usuario_diferente"]

            if status == "aberto":
                unidade_nome = inventario["unidade_nome"]
                cur.execute(
                    "SELECT COUNT(*) AS n FROM items WHERE is_active = 1 AND revenda = %s",
                    (unidade_nome,),
                )
                resumo["esperados"] = cur.fetchone()["n"]
                cur.execute(
                    "SELECT COUNT(*) AS n FROM items i WHERE i.is_active = 1 AND i.revenda = %s "
                    "AND NOT EXISTS (SELECT 1 FROM inventario_leituras l "
                    "WHERE l.inventario_id = %s AND l.item_id = i.id)",
                    (unidade_nome, inventario_id),
                )
                resumo["nao_encontrados"] = cur.fetchone()["n"]
            elif status == "encerrado":
                cur.execute(
                    "SELECT COUNT(*) AS n FROM inventario_nao_encontrados WHERE inventario_id = %s",
                    (inventario_id,),
                )
                resumo["nao_encontrados"] = cur.fetchone()["n"]
                resumo["esperados"] = resumo["encontrados"] + resumo["nao_encontrados"]
            else:
                resumo["nao_encontrados"] = 0
                resumo["esperados"] = resumo["encontrados"]
        return resumo

    def cancelar(self, inventario_id: int, motivo: str, por: str) -> None:
        with get_cursor() as cur:
            inventario = self._travar(cur, inventario_id)
            if inventario["status"] != "aberto":
                raise InventarioFechado("Só inventários abertos podem ser cancelados.")
            cur.execute(
                "UPDATE inventarios SET status = 'cancelado', encerrado_por = %s, "
                "encerrado_em = %s, motivo_cancelamento = %s WHERE id = %s",
                (por, datetime.now(), motivo, inventario_id),
            )
