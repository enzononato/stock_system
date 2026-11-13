"""Ordens de manutenção de equipamentos (F4).

A tabela `manutencoes` vem da migração 0403. Toda escrita roda numa transação
com `SELECT ... FOR UPDATE` na linha do item (sempre primeiro) e na da
manutenção, para duas ações simultâneas não furarem a máquina de estados.
Horários usam `datetime.now()` (hora local do servidor, igual a
`history.data_operacao`) para a linha do tempo ordenar fontes diferentes na
mesma escala.
"""
import logging
from datetime import date, datetime, timedelta
from decimal import Decimal
from typing import Literal, Optional

import pymysql

from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)

STATUS_MANUTENCAO = ("aberta", "enviada", "em_reparo", "retornada", "sem_conserto", "cancelada")
STATUS_EM_ANDAMENTO = ("aberta", "enviada", "em_reparo")
ROTULOS_STATUS = {
    "aberta": "Aberta",
    "enviada": "Enviada",
    "em_reparo": "Em reparo",
    "retornada": "Retornada",
    "sem_conserto": "Sem conserto",
    "cancelada": "Cancelada",
}
TRANSICOES = {
    "aberta": {"enviada", "em_reparo", "retornada", "sem_conserto", "cancelada"},
    "enviada": {"em_reparo", "retornada", "sem_conserto", "cancelada"},
    "em_reparo": {"retornada", "sem_conserto", "cancelada"},
    "retornada": set(),
    "sem_conserto": set(),
    "cancelada": set(),
}

_EM_ANDAMENTO_SQL = ", ".join(f"'{s}'" for s in STATUS_EM_ANDAMENTO)

_SELECT_COM_ITEM = (
    "SELECT m.*, i.tipo AS item_tipo, i.brand AS item_brand, i.model AS item_model, "
    "i.identificador AS item_identificador, i.revenda AS item_revenda, i.status AS item_status, "
    "(SELECT COALESCE(SUM(x.custo), 0) FROM manutencoes x WHERE x.item_id = m.item_id) "
    "AS custo_acumulado_item "
    "FROM manutencoes m JOIN items i ON i.id = m.item_id"
)


class ManutencaoErro(Exception):
    """Regra de negócio violada; o router converte em HTTPException."""

    def __init__(self, status_code: int, detail: str):
        super().__init__(detail)
        self.status_code = status_code
        self.detail = detail


def _acrescentar_observacao(texto: Optional[str]):
    """(fragmento SQL, parâmetros) que acrescenta `texto` às observações (máx. 1000)."""
    if not texto:
        return "", []
    return ", observacoes = LEFT(CONCAT_WS(CHAR(10), observacoes, %s), 1000)", [texto]


class ManutencaoDBManager:
    # ── escrita ────────────────────────────────────────────────────────────

    def abrir(
        self, item_id: int, defeito: str, assistencia: Optional[str], observacoes: Optional[str], usuario: str
    ) -> int:
        agora = datetime.now()
        try:
            with get_cursor() as cur:
                cur.execute(
                    "SELECT id, status, colaborador_id, assigned_to FROM items "
                    "WHERE id=%s AND is_active=1 FOR UPDATE",
                    (item_id,),
                )
                item = cur.fetchone()
                if not item:
                    raise ManutencaoErro(404, "Item não encontrado.")
                if (
                    item["colaborador_id"] is not None
                    or item["assigned_to"]
                    or item["status"] not in ("Disponível", "Em Manutenção")
                ):
                    raise ManutencaoErro(
                        400,
                        "Só é possível abrir manutenção para equipamento sem colaborador. "
                        "Faça a devolução antes.",
                    )
                cur.execute(
                    f"SELECT COUNT(*) AS total FROM manutencoes "
                    f"WHERE item_id=%s AND status IN ({_EM_ANDAMENTO_SQL})",
                    (item_id,),
                )
                if cur.fetchone()["total"] > 0:
                    raise ManutencaoErro(400, "Este equipamento já tem uma manutenção em andamento.")
                # Status "Em Manutenção" sem ordem em andamento = depois de "sem conserto": reabrir é permitido.
                cur.execute(
                    "INSERT INTO manutencoes (item_id, status, defeito, assistencia, aberta_por, aberta_em, observacoes) "
                    "VALUES (%s,'aberta',%s,%s,%s,%s,%s)",
                    (item_id, defeito, assistencia, usuario, agora, observacoes),
                )
                manutencao_id = cur.lastrowid
                cur.execute("UPDATE items SET status='Em Manutenção' WHERE id=%s", (item_id,))
            return manutencao_id
        except ManutencaoErro:
            raise
        except pymysql.MySQLError:
            logger.exception("Erro de banco de dados")
            raise ManutencaoErro(400, "Erro ao processar a operação. Tente novamente.")

    def _travar(self, cur, manutencao_id: int) -> dict:
        """Trava item e manutenção (nessa ordem) e devolve a manutenção."""
        cur.execute("SELECT item_id FROM manutencoes WHERE id=%s", (manutencao_id,))
        achada = cur.fetchone()
        if not achada:
            raise ManutencaoErro(404, "Manutenção não encontrada.")
        cur.execute("SELECT id FROM items WHERE id=%s FOR UPDATE", (achada["item_id"],))
        cur.execute("SELECT * FROM manutencoes WHERE id=%s FOR UPDATE", (manutencao_id,))
        manutencao = cur.fetchone()
        if not manutencao:
            raise ManutencaoErro(404, "Manutenção não encontrada.")
        return manutencao

    @staticmethod
    def _checar_transicao(manutencao: dict, novo: str) -> None:
        atual = manutencao["status"]
        if novo not in TRANSICOES[atual]:
            raise ManutencaoErro(
                400, f"Ação não permitida para manutenção com status {ROTULOS_STATUS[atual]}."
            )

    @staticmethod
    def _checar_previsao(manutencao: dict, previsao: Optional[date]) -> None:
        if previsao is not None and previsao < manutencao["aberta_em"].date():
            raise ManutencaoErro(400, "A previsão de retorno não pode ser anterior à abertura.")

    def enviar(self, manutencao_id: int, assistencia: str, previsao_retorno: Optional[date], usuario: str) -> None:
        try:
            with get_cursor() as cur:
                m = self._travar(cur, manutencao_id)
                self._checar_transicao(m, "enviada")
                self._checar_previsao(m, previsao_retorno)
                cur.execute(
                    "UPDATE manutencoes SET status='enviada', assistencia=%s, enviada_em=%s, "
                    "previsao_retorno=COALESCE(%s, previsao_retorno) WHERE id=%s",
                    (assistencia, datetime.now(), previsao_retorno, manutencao_id),
                )
        except ManutencaoErro:
            raise
        except pymysql.MySQLError:
            logger.exception("Erro de banco de dados")
            raise ManutencaoErro(400, "Erro ao processar a operação. Tente novamente.")

    def marcar_em_reparo(self, manutencao_id: int, previsao_retorno: Optional[date], usuario: str) -> None:
        try:
            with get_cursor() as cur:
                m = self._travar(cur, manutencao_id)
                self._checar_transicao(m, "em_reparo")
                self._checar_previsao(m, previsao_retorno)
                cur.execute(
                    "UPDATE manutencoes SET status='em_reparo', em_reparo_em=%s, "
                    "previsao_retorno=COALESCE(%s, previsao_retorno) WHERE id=%s",
                    (datetime.now(), previsao_retorno, manutencao_id),
                )
        except ManutencaoErro:
            raise
        except pymysql.MySQLError:
            logger.exception("Erro de banco de dados")
            raise ManutencaoErro(400, "Erro ao processar a operação. Tente novamente.")

    def registrar_retorno(
        self,
        manutencao_id: int,
        resultado: Literal["retornada", "sem_conserto"],
        custo: Optional[Decimal],
        observacoes: Optional[str],
        anexo_key: Optional[str],
        usuario: str,
    ) -> None:
        try:
            with get_cursor() as cur:
                m = self._travar(cur, manutencao_id)
                self._checar_transicao(m, resultado)
                obs_sql, obs_params = _acrescentar_observacao(observacoes)
                cur.execute(
                    "UPDATE manutencoes SET status=%s, retornada_em=%s, encerrada_por=%s, "
                    "custo=%s, anexo_key=COALESCE(%s, anexo_key)" + obs_sql + " WHERE id=%s",
                    [resultado, datetime.now(), usuario, custo, anexo_key, *obs_params, manutencao_id],
                )
                if resultado == "retornada":
                    cur.execute("UPDATE items SET status='Disponível' WHERE id=%s", (m["item_id"],))
                # sem_conserto: o item continua "Em Manutenção" até a baixa (ou uma nova ordem).
        except ManutencaoErro:
            raise
        except pymysql.MySQLError:
            logger.exception("Erro de banco de dados")
            raise ManutencaoErro(400, "Erro ao processar a operação. Tente novamente.")

    def cancelar(self, manutencao_id: int, motivo: str, usuario: str) -> None:
        try:
            with get_cursor() as cur:
                m = self._travar(cur, manutencao_id)
                self._checar_transicao(m, "cancelada")
                obs_sql, obs_params = _acrescentar_observacao(f"Cancelada: {motivo}")
                cur.execute(
                    "UPDATE manutencoes SET status='cancelada', cancelada_em=%s, encerrada_por=%s"
                    + obs_sql + " WHERE id=%s",
                    [datetime.now(), usuario, *obs_params, manutencao_id],
                )
                cur.execute("UPDATE items SET status='Disponível' WHERE id=%s", (m["item_id"],))
        except ManutencaoErro:
            raise
        except pymysql.MySQLError:
            logger.exception("Erro de banco de dados")
            raise ManutencaoErro(400, "Erro ao processar a operação. Tente novamente.")

    # ── leitura ────────────────────────────────────────────────────────────

    def obter(self, manutencao_id: int) -> Optional[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(_SELECT_COM_ITEM + " WHERE m.id=%s", (manutencao_id,))
            return cur.fetchone()

    def tem_em_andamento(self, item_id: int) -> bool:
        with get_cursor(commit=False) as cur:
            cur.execute(
                f"SELECT COUNT(*) AS total FROM manutencoes "
                f"WHERE item_id=%s AND status IN ({_EM_ANDAMENTO_SQL})",
                (item_id,),
            )
            return cur.fetchone()["total"] > 0

    def listar(
        self,
        *,
        status: Optional[str] = None,
        item_id: Optional[int] = None,
        revenda: Optional[str] = None,
        tipo: Optional[str] = None,
        de: Optional[date] = None,
        ate: Optional[date] = None,
        busca: Optional[str] = None,
        limit: Optional[int] = 25,
        offset: int = 0,
    ) -> tuple[list[dict], int, Decimal]:
        """(linhas da página, total filtrado, soma de `custo` de TODAS as linhas filtradas).

        O JOIN com `items` não filtra `is_active`: itens baixados mantêm o histórico
        de custo. `limit=None` devolve tudo (exportação)."""
        where = ["1=1"]
        params: list = []
        if status == "em_andamento":
            where.append(f"m.status IN ({_EM_ANDAMENTO_SQL})")
        elif status:
            where.append("m.status = %s")
            params.append(status)
        if item_id is not None:
            where.append("m.item_id = %s")
            params.append(item_id)
        if revenda:
            where.append("i.revenda = %s")
            params.append(revenda)
        if tipo:
            where.append("i.tipo = %s")
            params.append(tipo)
        if de is not None:
            where.append("m.aberta_em >= %s")
            params.append(datetime.combine(de, datetime.min.time()))
        if ate is not None:
            where.append("m.aberta_em < %s")
            params.append(datetime.combine(ate + timedelta(days=1), datetime.min.time()))
        if busca:
            like = f"%{busca}%"
            where.append(
                "(i.brand LIKE %s OR i.model LIKE %s OR i.identificador LIKE %s "
                "OR m.assistencia LIKE %s OR m.defeito LIKE %s)"
            )
            params.extend([like] * 5)
        where_sql = " AND ".join(where)

        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT COUNT(*) AS total, COALESCE(SUM(m.custo), 0) AS custo_total "
                "FROM manutencoes m JOIN items i ON i.id = m.item_id WHERE " + where_sql,
                params,
            )
            resumo = cur.fetchone()
            consulta = _SELECT_COM_ITEM + " WHERE " + where_sql + " ORDER BY m.aberta_em DESC, m.id DESC"
            consulta_params = list(params)
            if limit is not None:
                consulta += " LIMIT %s OFFSET %s"
                consulta_params.extend([limit, offset])
            cur.execute(consulta, consulta_params)
            linhas = cur.fetchall()
        return linhas, resumo["total"], Decimal(resumo["custo_total"])
