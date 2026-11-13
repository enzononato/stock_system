"""Camada de acesso a `termo_solicitacoes` (migração 0300).

Cada linha é um pedido de assinatura de termo (entrega, devolução ou uso) de um equipamento.
Todas as datas são horário de Brasília sem fuso (`app.core.tempo.agora_local`). Este módulo
não decide regra de negócio (isso é de `app.services.termos*`): só grava e lê, com
atualizações condicionais ao status para que duas rotinas concorrentes não se atropelem.
"""
from datetime import datetime
from typing import Optional

from app.core.tempo import agora_local
from app.db.database_mysql import get_cursor

PENDENTES = ("agendado", "enviado", "manual", "erro")

COLUNAS_ATUALIZAVEIS = {
    "status", "colaborador_id", "modelo_versao", "destinatario_nome", "destinatario_email",
    "agendado_para", "enviado_em", "assinado_em", "cancelado_em", "revsign_documento_id",
    "documento_key", "arquivo_assinado_key", "motivo",
}

# Situações do item em que existe um termo "vigente".
_STATUS_ITEM_COM_TERMO = ("Pendente", "Indisponível", "Pendente Devolução")

_SELECT = """
    SELECT t.*, c.nome AS colaborador_nome,
           CONCAT_WS(' ', i.tipo, i.brand, i.model) AS item_descricao
    FROM termo_solicitacoes t
    LEFT JOIN colaboradores c ON c.id = t.colaborador_id
    LEFT JOIN items i ON i.id = t.item_id
"""

_CHAVES_LEGADO = (
    "id", "item_id", "colaborador_id", "colaborador_nome", "item_descricao", "tipo", "history_id",
    "modelo_versao", "status", "destinatario_nome", "destinatario_email", "solicitado_por",
    "solicitado_em", "agendado_para", "enviado_em", "assinado_em", "cancelado_em",
    "revsign_documento_id", "documento_key", "arquivo_assinado_key", "motivo", "atualizado_em",
)


def _marcadores(valores) -> str:
    return ", ".join(["%s"] * len(valores))


class TermoDBManager:
    def criar(
        self,
        *,
        item_id: int,
        tipo: str,
        status: str,
        solicitado_por: str,
        solicitado_em: datetime,
        colaborador_id: Optional[int] = None,
        history_id: Optional[int] = None,
        agendado_para: Optional[datetime] = None,
        destinatario_nome: Optional[str] = None,
        destinatario_email: Optional[str] = None,
    ) -> int:
        with get_cursor() as cur:
            cur.execute(
                """INSERT INTO termo_solicitacoes
                (item_id, colaborador_id, tipo, history_id, status, destinatario_nome,
                 destinatario_email, solicitado_por, solicitado_em, agendado_para, atualizado_em)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (item_id, colaborador_id, tipo, history_id, status, destinatario_nome,
                 destinatario_email, solicitado_por, solicitado_em, agendado_para, agora_local()),
            )
            return cur.lastrowid

    def obter(self, termo_id: int) -> Optional[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(f"{_SELECT} WHERE t.id = %s", (termo_id,))
            return cur.fetchone()

    def obter_por_documento(self, revsign_documento_id: str) -> Optional[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(f"{_SELECT} WHERE t.revsign_documento_id = %s", (str(revsign_documento_id),))
            return cur.fetchone()

    def atualizar(
        self, termo_id: int, *, somente_se_status: Optional[tuple[str, ...]] = None, **campos
    ) -> bool:
        invalidas = set(campos) - COLUNAS_ATUALIZAVEIS
        if invalidas:
            raise ValueError(f"Coluna(s) não atualizável(is): {', '.join(sorted(invalidas))}")
        sets = [f"{coluna} = %s" for coluna in campos] + ["atualizado_em = %s"]
        params: list = [*campos.values(), agora_local(), termo_id]
        sql = f"UPDATE termo_solicitacoes SET {', '.join(sets)} WHERE id = %s"
        if somente_se_status:
            sql += f" AND status IN ({_marcadores(somente_se_status)})"
            params.extend(somente_se_status)
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(sql, params)
            return cur.rowcount > 0

    def pendentes_do_item(self, item_id: int, *, tipo: Optional[str] = None) -> list[dict]:
        sql = f"{_SELECT} WHERE t.item_id = %s AND t.status IN ({_marcadores(PENDENTES)})"
        params: list = [item_id, *PENDENTES]
        if tipo is not None:
            sql += " AND t.tipo = %s"
            params.append(tipo)
        with get_cursor(commit=False) as cur:
            cur.execute(f"{sql} ORDER BY t.id DESC", params)
            return list(cur.fetchall())

    def assinadas_do_item(self, item_id: int, tipo: str) -> list[dict]:
        """Solicitações `assinado` do item para o tipo, da mais recente para a mais antiga."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                f"{_SELECT} WHERE t.item_id = %s AND t.tipo = %s AND t.status = 'assinado' ORDER BY t.id DESC",
                (item_id, tipo),
            )
            return list(cur.fetchall())

    def por_history(self, history_id: int) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(f"{_SELECT} WHERE t.history_id = %s ORDER BY t.id DESC", (history_id,))
            return list(cur.fetchall())

    def _listar(self, coluna: str, valor: int, limit: int, offset: int) -> tuple[list[dict], int]:
        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT COUNT(*) AS total FROM termo_solicitacoes t WHERE t.{coluna} = %s", (valor,))
            total = cur.fetchone()["total"]
            cur.execute(
                f"{_SELECT} WHERE t.{coluna} = %s ORDER BY t.id DESC LIMIT %s OFFSET %s",
                (valor, limit, offset),
            )
            return list(cur.fetchall()), total

    def listar_por_item(self, item_id: int, *, limit: int, offset: int) -> tuple[list[dict], int]:
        return self._listar("item_id", item_id, limit, offset)

    def listar_por_colaborador(
        self, colaborador_id: int, *, limit: int, offset: int
    ) -> tuple[list[dict], int]:
        return self._listar("colaborador_id", colaborador_id, limit, offset)

    def vigente(self, item_id: int) -> Optional[dict]:
        """Termo vigente do equipamento.

        É a última solicitação não cancelada, só enquanto o item está `Pendente`,
        `Indisponível` ou `Pendente Devolução`. Sem solicitação e com um
        `Confirmação Empréstimo` ativo (posterior ao último `Empréstimo`) no histórico, devolve
        o termo "legado" (anterior à F3), já assinado.
        """
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT status FROM items WHERE id = %s AND is_active = 1", (item_id,))
            item = cur.fetchone()
            if not item or item["status"] not in _STATUS_ITEM_COM_TERMO:
                return None
            cur.execute(
                f"{_SELECT} WHERE t.item_id = %s AND t.status <> 'cancelado' ORDER BY t.id DESC LIMIT 1",
                (item_id,),
            )
            linha = cur.fetchone()
            if linha:
                linha["legado"] = False
                return linha
            cur.execute(
                """SELECT id, operador, data_operacao, termo_assinado_anexo FROM history
                WHERE item_id = %s AND operation = 'Confirmação Empréstimo' AND is_reversed = 0
                ORDER BY data_operacao DESC, id DESC LIMIT 1""",
                (item_id,),
            )
            confirmacao = cur.fetchone()
            if not confirmacao:
                return None
            cur.execute(
                """SELECT MAX(id) AS ultimo FROM history
                WHERE item_id = %s AND operation = 'Empréstimo' AND is_reversed = 0""",
                (item_id,),
            )
            emprestimo = cur.fetchone()
            if emprestimo and emprestimo["ultimo"] and confirmacao["id"] < emprestimo["ultimo"]:
                return None  # a confirmação é de um empréstimo anterior
        legado = {chave: None for chave in _CHAVES_LEGADO}
        legado.update(
            item_id=item_id,
            tipo="entrega",
            status="assinado",
            solicitado_por=confirmacao["operador"],
            solicitado_em=confirmacao["data_operacao"],
            assinado_em=confirmacao["data_operacao"],
            arquivo_assinado_key=confirmacao["termo_assinado_anexo"],
            legado=True,
        )
        return legado

    def agendados_vencidos(self, agora: datetime) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                f"{_SELECT} WHERE t.status = 'agendado' AND t.agendado_para <= %s ORDER BY t.id",
                (agora,),
            )
            return list(cur.fetchall())

    def enviados(self) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                f"{_SELECT} WHERE t.status = 'enviado' AND t.revsign_documento_id IS NOT NULL ORDER BY t.id"
            )
            return list(cur.fetchall())
