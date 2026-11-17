"""Consulta da trilha de auditoria (`audit_log`). A escrita fica em
`app/services/audit.py`; a tabela é criada pela migração 0005."""
import json
from datetime import date, datetime, time, timedelta, timezone
from typing import Optional

from app.db.database_mysql import get_cursor


def _json_ou_none(valor):
    if valor is None:
        return None
    if isinstance(valor, (bytes, bytearray)):
        valor = valor.decode("utf-8")
    if isinstance(valor, str):
        return json.loads(valor)
    return valor


class AuditDBManager:
    def listar(
        self,
        *,
        entidade: Optional[str] = None,
        entidade_id: Optional[str] = None,
        username: Optional[str] = None,
        de: Optional[date] = None,
        ate: Optional[date] = None,
        limit: int = 50,
        offset: int = 0,
    ) -> tuple[list[dict], int]:
        """Registros mais recentes primeiro. `de`/`ate` (inclusivos) comparam
        com `ocorrido_em`, que está em UTC."""
        condicoes: list[str] = []
        params: list = []
        if entidade:
            condicoes.append("entidade = %s")
            params.append(entidade)
        if entidade_id:
            condicoes.append("entidade_id = %s")
            params.append(entidade_id)
        if username:
            condicoes.append("username = %s")
            params.append(username)
        if de:
            condicoes.append("ocorrido_em >= %s")
            params.append(datetime.combine(de, time.min))
        if ate:
            condicoes.append("ocorrido_em < %s")
            params.append(datetime.combine(ate + timedelta(days=1), time.min))
        where = f"WHERE {' AND '.join(condicoes)}" if condicoes else ""

        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT COUNT(*) AS total FROM audit_log {where}", params)
            total = cur.fetchone()["total"]
            cur.execute(
                f"SELECT * FROM audit_log {where} "
                "ORDER BY ocorrido_em DESC, id DESC LIMIT %s OFFSET %s",
                [*params, limit, offset],
            )
            linhas = cur.fetchall()

        for linha in linhas:
            linha["ocorrido_em"] = linha["ocorrido_em"].replace(tzinfo=timezone.utc)
            linha["antes"] = _json_ou_none(linha["antes"])
            linha["depois"] = _json_ou_none(linha["depois"])
        return linhas, total
