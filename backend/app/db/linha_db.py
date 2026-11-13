"""
Camada de acesso às linhas telefônicas (F5).

As tabelas `linhas` e `linha_eventos` vêm da migração 0500; este módulo não cria
schema. Cada escrita roda num único `get_cursor()`: a alteração da linha, os eventos
em `linha_eventos` e a auditoria entram na mesma transação (falhou um, desfaz tudo).
Datas em hora local do servidor (`datetime.now()`).
"""
import logging
import re
from datetime import datetime
from decimal import Decimal
from typing import Optional

import pymysql

from app.db.database_mysql import get_cursor
from app.schemas.linhas import TIPOS_APARELHO, formatar_numero_linha
from app.services import audit
from app.services.audit import AuditContext

logger = logging.getLogger(__name__)


class LinhaNaoEncontrada(Exception):
    pass


class LinhaInvalida(Exception):
    pass


class LinhaConflito(Exception):
    pass


MSG_NUMERO_DUPLICADO = "Já existe uma linha com este número."

_FROM_LINHA = (
    " FROM linhas l JOIN unidades u ON u.id = l.unidade_id"
    " LEFT JOIN items i ON i.id = l.item_id"
    " LEFT JOIN colaboradores c ON c.id = l.colaborador_id"
)
_SELECT_LINHA = (
    "SELECT l.*, u.nome AS unidade_nome, i.tipo AS item_tipo, i.brand AS item_brand, "
    "i.model AS item_model, i.identificador AS item_identificador, i.is_active AS item_ativo, "
    "c.nome AS colaborador_nome, c.situacao AS colaborador_situacao"
    + _FROM_LINHA
)

_CAMPOS_EDITAVEIS = (
    "numero", "operadora", "plano", "custo_mensal", "unidade_id",
    "item_id", "colaborador_id", "iccid", "observacoes",
)


def _erro_de_duplicado(e: pymysql.MySQLError) -> bool:
    return bool(e.args) and e.args[0] == 1062


def _escapar_like(texto: str) -> str:
    return texto.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def _registrar_evento(
    cur,
    linha_id: int,
    tipo: str,
    *,
    por_usuario: Optional[str],
    de_valor: Optional[str] = None,
    para_valor: Optional[str] = None,
    campo: Optional[str] = None,
    item_id: Optional[int] = None,
    item_id_anterior: Optional[int] = None,
    observacao: Optional[str] = None,
) -> None:
    cur.execute(
        "INSERT INTO linha_eventos (linha_id, tipo, campo, de_valor, para_valor, item_id, "
        "item_id_anterior, ocorrido_em, por_usuario, observacao) "
        "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
        (linha_id, tipo, campo, de_valor, para_valor, item_id, item_id_anterior,
         datetime.now(), por_usuario, observacao),
    )


def _descrever_item(item: Optional[dict]) -> Optional[str]:
    """`Celular Samsung A15 · CEL-A` (tipo, marca e modelo; identificador quando houver)."""
    if not item:
        return None
    partes = [str(item[k]) for k in ("tipo", "brand", "model") if item.get(k)]
    texto = " ".join(partes)
    if item.get("identificador"):
        texto = f"{texto} · {item['identificador']}" if texto else str(item["identificador"])
    return texto or None


def _validar_unidade(cur, unidade_id: int) -> dict:
    cur.execute("SELECT id, nome, is_active FROM unidades WHERE id = %s", (unidade_id,))
    unidade = cur.fetchone()
    if not unidade or not unidade["is_active"]:
        raise LinhaInvalida("Unidade não encontrada ou inativa.")
    return unidade


def _validar_aparelho(cur, item_id: int, linha_id: Optional[int]) -> dict:
    cur.execute(
        "SELECT id, tipo, brand, model, identificador, is_active FROM items WHERE id = %s FOR UPDATE",
        (item_id,),
    )
    item = cur.fetchone()
    if not item:
        raise LinhaInvalida("Aparelho não encontrado.")
    if not item["is_active"]:
        raise LinhaInvalida("Aparelho baixado não pode receber linha.")
    if item["tipo"] not in TIPOS_APARELHO:
        raise LinhaInvalida("Só Celular ou Tablet podem receber uma linha.")
    sql = "SELECT numero FROM linhas WHERE item_id = %s AND status <> 'cancelada'"
    params: list = [item_id]
    if linha_id is not None:
        sql += " AND id <> %s"
        params.append(linha_id)
    cur.execute(sql + " FOR UPDATE", params)
    outra = cur.fetchone()
    if outra:
        raise LinhaConflito(
            f"O aparelho já está vinculado à linha {formatar_numero_linha(outra['numero'])}."
        )
    return item


def _validar_colaborador(cur, colaborador_id: int) -> dict:
    cur.execute("SELECT id, nome, situacao FROM colaboradores WHERE id = %s", (colaborador_id,))
    colaborador = cur.fetchone()
    if not colaborador:
        raise LinhaInvalida("Colaborador não encontrado.")
    if colaborador["situacao"] == "desligado":
        raise LinhaInvalida("Colaborador desligado não pode receber linha.")
    return colaborador


def _checar_numero_livre(cur, numero: str, linha_id: Optional[int]) -> None:
    sql = "SELECT id, status FROM linhas WHERE numero = %s"
    params: list = [numero]
    if linha_id is not None:
        sql += " AND id <> %s"
        params.append(linha_id)
    cur.execute(sql, params)
    existente = cur.fetchone()
    if not existente:
        return
    if existente["status"] == "cancelada":
        raise LinhaConflito(
            "Já existe uma linha cancelada com este número: reative-a em vez de cadastrar de novo."
        )
    raise LinhaConflito(MSG_NUMERO_DUPLICADO)


def _ler_linha_crua(cur, linha_id: int) -> Optional[dict]:
    cur.execute("SELECT * FROM linhas WHERE id = %s", (linha_id,))
    return cur.fetchone()


class LinhaDBManager:
    def listar(
        self,
        *,
        busca: Optional[str] = None,
        status: Optional[str] = None,
        unidade_id: Optional[int] = None,
        operadora: Optional[str] = None,
        item_id: Optional[int] = None,
        colaborador_id: Optional[int] = None,
        limit: int = 25,
        offset: int = 0,
    ) -> tuple[list[dict], int]:
        where: list[str] = []
        params: list = []
        if status:
            where.append("l.status = %s")
            params.append(status)
        if unidade_id is not None:
            where.append("l.unidade_id = %s")
            params.append(unidade_id)
        if operadora:
            where.append("l.operadora = %s")
            params.append(operadora)
        if item_id is not None:
            where.append("l.item_id = %s")
            params.append(item_id)
        if colaborador_id is not None:
            where.append("l.colaborador_id = %s")
            params.append(colaborador_id)
        busca = (busca or "").strip()
        if busca:
            like = f"%{_escapar_like(busca)}%"
            condicoes = [
                "l.plano LIKE %s", "l.operadora LIKE %s", "c.nome LIKE %s", "i.identificador LIKE %s",
            ]
            params.extend([like] * 4)
            digitos = re.sub(r"\D", "", busca)
            if digitos:
                condicoes += ["l.numero LIKE %s", "l.iccid LIKE %s"]
                params.extend([f"%{digitos}%"] * 2)
            where.append("(" + " OR ".join(condicoes) + ")")
        filtro = (" WHERE " + " AND ".join(where)) if where else ""
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT COUNT(*) AS n" + _FROM_LINHA + filtro, params)
            total = cur.fetchone()["n"]
            cur.execute(
                _SELECT_LINHA + filtro
                + " ORDER BY FIELD(l.status,'ativa','suspensa','cancelada'), l.numero, l.id LIMIT %s OFFSET %s",
                [*params, limit, offset],
            )
            return cur.fetchall(), total

    def obter(self, linha_id: int) -> Optional[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(_SELECT_LINHA + " WHERE l.id = %s", (linha_id,))
            return cur.fetchone()

    def criar(self, dados: dict, ctx: AuditContext) -> int:
        """Cadastra a linha (sempre `ativa`) com o evento `ativacao` e a auditoria na mesma transação."""
        numero = dados["numero"]
        item_id = dados.get("item_id")
        colaborador_id = dados.get("colaborador_id")
        try:
            with get_cursor() as cur:
                _checar_numero_livre(cur, numero, None)
                _validar_unidade(cur, dados["unidade_id"])
                if item_id is not None:
                    _validar_aparelho(cur, item_id, None)
                if colaborador_id is not None:
                    _validar_colaborador(cur, colaborador_id)
                agora = datetime.now()
                cur.execute(
                    "INSERT INTO linhas (numero, operadora, plano, custo_mensal, status, unidade_id, "
                    "item_id, colaborador_id, iccid, observacoes, criada_em, atualizada_em) "
                    "VALUES (%s, %s, %s, %s, 'ativa', %s, %s, %s, %s, %s, %s, %s)",
                    (numero, dados["operadora"], dados.get("plano"), dados["custo_mensal"],
                     dados["unidade_id"], item_id, colaborador_id, dados.get("iccid"),
                     dados.get("observacoes"), agora, agora),
                )
                linha_id = cur.lastrowid
                _registrar_evento(cur, linha_id, "ativacao", por_usuario=ctx.username, item_id=item_id)
                audit.record(
                    ctx, "linha", linha_id, "criar",
                    depois=_ler_linha_crua(cur, linha_id), cur=cur,
                )
                return linha_id
        except pymysql.MySQLError as e:
            if _erro_de_duplicado(e):
                raise LinhaConflito(MSG_NUMERO_DUPLICADO) from e
            raise

    def aparelhos(
        self, *, busca: Optional[str] = None, linha_id: Optional[int] = None, limit: int = 20
    ) -> list[dict]:
        """Celulares e tablets ativos sem linha (a linha `linha_id` não conta como ocupação)."""
        placeholders = ", ".join(["%s"] * len(TIPOS_APARELHO))
        params: list = list(TIPOS_APARELHO)
        sql = (
            "SELECT i.id, i.tipo, i.brand, i.model, i.identificador, i.codigo_patrimonial, i.revenda, "
            "i.status, COALESCE(c.nome, i.assigned_to) AS responsavel_nome "
            "FROM items i LEFT JOIN colaboradores c ON c.id = i.colaborador_id "
            f"WHERE i.is_active = 1 AND i.tipo IN ({placeholders}) "
            "AND NOT EXISTS (SELECT 1 FROM linhas l WHERE l.item_id = i.id AND l.status <> 'cancelada'"
        )
        if linha_id is not None:
            sql += " AND l.id <> %s"
            params.append(linha_id)
        sql += ")"
        busca = (busca or "").strip()
        if busca:
            like = f"%{_escapar_like(busca)}%"
            sql += (
                " AND (i.brand LIKE %s OR i.model LIKE %s OR i.identificador LIKE %s "
                "OR i.codigo_patrimonial LIKE %s OR COALESCE(c.nome, i.assigned_to) LIKE %s)"
            )
            params.extend([like] * 5)
        sql += " ORDER BY i.tipo, i.brand, i.model, i.id LIMIT %s"
        params.append(limit)
        with get_cursor(commit=False) as cur:
            cur.execute(sql, params)
            return cur.fetchall()

    def atualizar(
        self, linha_id: int, dados: dict, ctx: AuditContext, motivo: Optional[str] = None
    ) -> dict:
        """Altera a linha e grava, na mesma transação, os eventos de cada mudança e a auditoria.

        `dados` só traz as chaves enviadas (`None` limpa plano, aparelho, colaborador, iccid e
        observações). Sem diferença real não há evento nem auditoria. Devolve a linha relida.
        """
        try:
            with get_cursor() as cur:
                cur.execute("SELECT * FROM linhas WHERE id = %s FOR UPDATE", (linha_id,))
                atual = cur.fetchone()
                if not atual:
                    raise LinhaNaoEncontrada()
                if atual["status"] == "cancelada":
                    raise LinhaConflito("Linha cancelada: reative-a antes de alterar.")
                mudancas = {
                    k: v for k, v in dados.items() if k in _CAMPOS_EDITAVEIS and v != atual[k]
                }
                if mudancas:
                    self._aplicar_mudancas(cur, linha_id, atual, mudancas, ctx, motivo)
        except pymysql.MySQLError as e:
            if _erro_de_duplicado(e):
                raise LinhaConflito(MSG_NUMERO_DUPLICADO) from e
            raise
        return self.obter(linha_id)

    @staticmethod
    def _aplicar_mudancas(
        cur, linha_id: int, atual: dict, mudancas: dict, ctx: AuditContext, motivo: Optional[str]
    ) -> None:
        if "numero" in mudancas:
            _checar_numero_livre(cur, mudancas["numero"], linha_id)
        unidade_nova = None
        if "unidade_id" in mudancas:
            unidade_nova = _validar_unidade(cur, mudancas["unidade_id"])
        item_novo = None
        if mudancas.get("item_id") is not None:
            item_novo = _validar_aparelho(cur, mudancas["item_id"], linha_id)
        colaborador_novo = None
        if mudancas.get("colaborador_id") is not None:
            colaborador_novo = _validar_colaborador(cur, mudancas["colaborador_id"])

        sets = ", ".join(f"{k} = %s" for k in mudancas)
        cur.execute(
            f"UPDATE linhas SET {sets}, atualizada_em = %s WHERE id = %s",
            [*mudancas.values(), datetime.now(), linha_id],
        )

        item_depois = mudancas["item_id"] if "item_id" in mudancas else atual["item_id"]
        base = {"por_usuario": ctx.username, "observacao": motivo, "item_id": item_depois}

        if "item_id" in mudancas:
            antigo = None
            if atual["item_id"] is not None:
                cur.execute(
                    "SELECT tipo, brand, model, identificador FROM items WHERE id = %s",
                    (atual["item_id"],),
                )
                antigo = cur.fetchone()
            _registrar_evento(
                cur, linha_id, "troca_aparelho",
                de_valor=_descrever_item(antigo), para_valor=_descrever_item(item_novo),
                item_id_anterior=atual["item_id"], **base,
            )
        if "colaborador_id" in mudancas:
            antigo_nome = None
            if atual["colaborador_id"] is not None:
                cur.execute("SELECT nome FROM colaboradores WHERE id = %s", (atual["colaborador_id"],))
                encontrado = cur.fetchone()
                antigo_nome = encontrado["nome"] if encontrado else None
            _registrar_evento(
                cur, linha_id, "troca_usuario", de_valor=antigo_nome,
                para_valor=colaborador_novo["nome"] if colaborador_novo else None, **base,
            )
        if "operadora" in mudancas:
            _registrar_evento(
                cur, linha_id, "portabilidade",
                de_valor=atual["operadora"], para_valor=mudancas["operadora"], **base,
            )
        for campo in ("numero", "plano", "custo_mensal", "unidade_id", "iccid", "observacoes"):
            if campo not in mudancas:
                continue
            de, para = atual[campo], mudancas[campo]
            if campo == "unidade_id":
                cur.execute("SELECT nome FROM unidades WHERE id = %s", (de,))
                encontrada = cur.fetchone()
                de = encontrada["nome"] if encontrada else de
                para = unidade_nova["nome"]
            _registrar_evento(
                cur, linha_id, "edicao",
                campo="unidade" if campo == "unidade_id" else campo,
                de_valor=None if de is None else str(de),
                para_valor=None if para is None else str(para), **base,
            )

        depois = _ler_linha_crua(cur, linha_id)
        antes_diff, depois_diff = audit.diff(
            {k: v for k, v in atual.items() if k != "atualizada_em"},
            {k: v for k, v in depois.items() if k != "atualizada_em"},
        )
        audit.record(ctx, "linha", linha_id, "editar", antes_diff, depois_diff, cur=cur)

    def alterar_status(
        self, linha_id: int, status: str, ctx: AuditContext, observacao: Optional[str] = None
    ) -> dict:
        """Suspende, reativa ou cancela (cancelar libera o aparelho). Devolve a linha relida."""
        with get_cursor() as cur:
            cur.execute("SELECT * FROM linhas WHERE id = %s FOR UPDATE", (linha_id,))
            atual = cur.fetchone()
            if not atual:
                raise LinhaNaoEncontrada()
            antigo = atual["status"]
            if status == "suspensa" and antigo != "ativa":
                raise LinhaConflito("Só linhas ativas podem ser suspensas.")
            if status == "ativa" and antigo == "ativa":
                raise LinhaConflito("A linha já está ativa.")
            if status == "cancelada" and antigo == "cancelada":
                raise LinhaConflito("A linha já está cancelada.")

            cancelando = status == "cancelada"
            item_depois = None if cancelando else atual["item_id"]
            cur.execute(
                "UPDATE linhas SET status = %s, item_id = %s, atualizada_em = %s WHERE id = %s",
                (status, item_depois, datetime.now(), linha_id),
            )
            tipo, acao = {
                "suspensa": ("suspensao", "suspender"),
                "ativa": ("reativacao", "reativar"),
                "cancelada": ("cancelamento", "cancelar"),
            }[status]
            _registrar_evento(
                cur, linha_id, tipo, por_usuario=ctx.username, de_valor=antigo, para_valor=status,
                item_id=item_depois,
                item_id_anterior=atual["item_id"] if cancelando else None,
                observacao=observacao,
            )
            antes = {"status": antigo}
            depois = {"status": status}
            if cancelando:
                antes["item_id"] = atual["item_id"]
                depois["item_id"] = None
            audit.record(ctx, "linha", linha_id, acao, antes, depois, cur=cur)
        return self.obter(linha_id)

    def eventos(self, linha_id: int) -> list[dict]:
        return self._eventos("e.linha_id = %s", (linha_id,))

    def eventos_do_aparelho(self, item_id: int) -> list[dict]:
        """Eventos das linhas que passaram pelo aparelho (fonte da linha do tempo do equipamento, F7)."""
        return self._eventos("(e.item_id = %s OR e.item_id_anterior = %s)", (item_id, item_id))

    @staticmethod
    def _eventos(condicao: str, params: tuple) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT e.*, l.numero AS linha_numero FROM linha_eventos e "
                "JOIN linhas l ON l.id = e.linha_id "
                f"WHERE {condicao} ORDER BY e.ocorrido_em DESC, e.id DESC LIMIT 200",
                params,
            )
            return cur.fetchall()
