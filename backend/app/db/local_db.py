"""
Camada de acesso aos locais físicos de cada unidade (sala, almoxarifado, andar...).

A tabela `locais` é criada pela migração 0003 (e `items.local_id` pela 0004); este
módulo não cria schema. O nome é único dentro da unidade. Locais nunca são apagados
pela API: ficam inativos (`ativo = 0`) e somem dos seletores.
"""
import logging
from typing import Optional

import pymysql

from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)

MSG_DUPLICADO = "Já existe um local com esse nome nesta unidade."

_CAMPOS_EDITAVEIS = ("nome", "descricao", "ativo")

_SELECT = (
    "SELECT l.id, l.unidade_id, u.nome AS unidade_nome, l.nome, l.descricao, l.ativo "
    "FROM locais l JOIN unidades u ON u.id = l.unidade_id"
)


def _erro_de_duplicado(e: pymysql.MySQLError) -> bool:
    return bool(e.args) and e.args[0] == 1062


class LocalDBManager:
    def listar(self, unidade_id: Optional[int] = None, include_inactive: bool = False) -> list[dict]:
        """Locais ordenados por nome. `include_inactive` diz respeito ao `ativo` do
        próprio local; os locais de uma unidade inativa continuam listados."""
        where: list[str] = []
        params: list = []
        if unidade_id is not None:
            where.append("l.unidade_id = %s")
            params.append(unidade_id)
        if not include_inactive:
            where.append("l.ativo = 1")
        sql = _SELECT + (" WHERE " + " AND ".join(where) if where else "")
        sql += " ORDER BY l.nome, l.id"
        with get_cursor(commit=False) as cur:
            cur.execute(sql, params)
            return cur.fetchall()

    def obter(self, local_id: int) -> Optional[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(_SELECT + " WHERE l.id = %s", (local_id,))
            return cur.fetchone()

    def criar(self, unidade_id: int, nome: str, descricao: Optional[str]) -> int:
        """Devolve o id. Nome já usado na unidade -> `ValueError(MSG_DUPLICADO)`."""
        try:
            with get_cursor(dict_cursor=False) as cur:
                cur.execute(
                    "INSERT INTO locais (unidade_id, nome, descricao) VALUES (%s, %s, %s)",
                    (unidade_id, nome, descricao),
                )
                return cur.lastrowid
        except pymysql.MySQLError as e:
            if _erro_de_duplicado(e):
                raise ValueError(MSG_DUPLICADO) from e
            raise

    def atualizar(self, local_id: int, dados: dict) -> None:
        """Atualiza só as chaves de `dados` (nome, descricao, ativo). A unidade de um
        local não muda. Nome já usado na unidade -> `ValueError(MSG_DUPLICADO)`."""
        campos = {k: v for k, v in dados.items() if k in _CAMPOS_EDITAVEIS}
        if not campos:
            return
        sets = ", ".join(f"{k} = %s" for k in campos)
        try:
            with get_cursor(dict_cursor=False) as cur:
                cur.execute(f"UPDATE locais SET {sets} WHERE id = %s", [*campos.values(), local_id])
        except pymysql.MySQLError as e:
            if _erro_de_duplicado(e):
                raise ValueError(MSG_DUPLICADO) from e
            raise
