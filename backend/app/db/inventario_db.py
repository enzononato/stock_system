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
from app.services.inventario import (
    Encontrado,
    ORDEM_CATEGORIAS,
    Esperado,
    chave_da_leitura,
    classificar,
    montar_detalhe,
    resolver_codigo,
)

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

    # ── Leituras ────────────────────────────────────────────────────────────

    _SELECT_LEITURA = (
        "SELECT l.*, i.tipo AS item_tipo, i.brand AS item_brand, i.model AS item_model, "
        "i.identificador AS item_identificador, i.status AS item_status, l.revenda_item AS item_revenda, "
        "le.nome AS local_encontrado_nome, lx.nome AS local_esperado_nome, "
        "c.nome AS colaborador_encontrado_nome "
        "FROM inventario_leituras l "
        "LEFT JOIN items i ON i.id = l.item_id "
        "LEFT JOIN locais le ON le.id = l.local_id_encontrado "
        "LEFT JOIN locais lx ON lx.id = l.local_id_esperado "
        "LEFT JOIN colaboradores c ON c.id = l.colaborador_id_encontrado"
    )

    def _leitura(self, cur, leitura_id: int) -> Optional[dict]:
        cur.execute(self._SELECT_LEITURA + " WHERE l.id = %s", (leitura_id,))
        return cur.fetchone()

    def obter_leitura(self, leitura_id: int) -> Optional[dict]:
        with get_cursor(commit=False) as cur:
            return self._leitura(cur, leitura_id)

    @staticmethod
    def _exigir_aberto(inventario: dict, mensagem: str) -> None:
        if inventario["status"] != "aberto":
            raise InventarioFechado(mensagem)

    @staticmethod
    def _validar_local(cur, inventario: dict, local_id: Optional[int]) -> Optional[str]:
        """Nome do local; precisa pertencer à unidade do inventário."""
        if local_id is None:
            return None
        cur.execute(
            "SELECT nome FROM locais WHERE id = %s AND unidade_id = %s",
            (local_id, inventario["unidade_id"]),
        )
        local = cur.fetchone()
        if not local:
            raise ValueError("O local informado não pertence à unidade do inventário.")
        return local["nome"]

    @staticmethod
    def _validar_colaborador(cur, colaborador_id: Optional[int]) -> Optional[str]:
        """Nome do colaborador; precisa existir."""
        if colaborador_id is None:
            return None
        cur.execute("SELECT nome FROM colaboradores WHERE id = %s", (colaborador_id,))
        colaborador = cur.fetchone()
        if not colaborador:
            raise ValueError("Colaborador não encontrado.")
        return colaborador["nome"]

    @staticmethod
    def _nome_do_local(cur, local_id: Optional[int]) -> Optional[str]:
        if local_id is None:
            return None
        cur.execute("SELECT nome FROM locais WHERE id = %s", (local_id,))
        linha = cur.fetchone()
        return linha["nome"] if linha else None

    def registrar_leitura(
        self,
        inventario_id: int,
        *,
        codigo: str,
        local_id_encontrado: Optional[int],
        colaborador_id_encontrado: Optional[int],
        usuario_conferido: bool,
        observacao: Optional[str],
        por: str,
    ) -> tuple[dict, bool]:
        """Registra uma leitura e devolve `(leitura, duplicada)`. A mesma leitura
        (mesma chave) repetida não grava nada e volta com `duplicada=True`."""
        with get_cursor() as cur:
            inventario = self._travar(cur, inventario_id)
            self._exigir_aberto(inventario, "Este inventário não está aberto: não aceita leituras.")
            local_encontrado_nome = self._validar_local(cur, inventario, local_id_encontrado)
            colaborador_nome = self._validar_colaborador(cur, colaborador_id_encontrado)
            if colaborador_id_encontrado is not None:
                usuario_conferido = True
            alvo = resolver_codigo(codigo)

            if alvo.tipo == "id":
                cur.execute(
                    "SELECT i.*, c.nome AS colaborador_nome FROM items i "
                    "LEFT JOIN colaboradores c ON c.id = i.colaborador_id WHERE i.id = %s",
                    (alvo.valor,),
                )
            else:
                cur.execute(
                    "SELECT i.*, c.nome AS colaborador_nome FROM items i "
                    "LEFT JOIN colaboradores c ON c.id = i.colaborador_id "
                    "WHERE UPPER(TRIM(i.identificador)) = %s",
                    (alvo.valor,),
                )
            itens = cur.fetchall()
            if len(itens) > 1:
                raise ValueError(
                    "Há mais de um equipamento com esse número de série. "
                    "Leia o QR code ou digite o número do patrimônio."
                )
            item = itens[0] if itens else None

            chave = chave_da_leitura(alvo, item["id"] if item else None)
            cur.execute(
                "SELECT id FROM inventario_leituras WHERE inventario_id = %s AND chave = %s",
                (inventario_id, chave),
            )
            existente = cur.fetchone()
            if existente:
                return self._leitura(cur, existente["id"]), True

            esperado = None
            if item:
                esperado = Esperado(
                    item_id=item["id"],
                    ativo=bool(item["is_active"]),
                    revenda=item["revenda"],
                    local_id=item["local_id"],
                    colaborador_id=item["colaborador_id"],
                    usuario_nome=item["colaborador_nome"] or item["assigned_to"],
                )
            resultado, divergencias = classificar(
                esperado,
                inventario["unidade_nome"],
                Encontrado(
                    local_id=local_id_encontrado,
                    usuario_conferido=usuario_conferido,
                    colaborador_id=colaborador_id_encontrado,
                    colaborador_nome=colaborador_nome,
                ),
            )
            detalhe = montar_detalhe(
                divergencias,
                revenda=item["revenda"] if item else None,
                local_esperado_nome=self._nome_do_local(cur, item["local_id"]) if item else None,
                local_encontrado_nome=local_encontrado_nome,
                usuario_esperado=esperado.usuario_nome if esperado else None,
                usuario_encontrado=colaborador_nome,
            )
            try:
                cur.execute(
                    "INSERT INTO inventario_leituras (inventario_id, item_id, chave, codigo_lido, "
                    "lido_por, lido_em, local_id_encontrado, usuario_conferido, "
                    "colaborador_id_encontrado, revenda_item, local_id_esperado, "
                    "colaborador_id_esperado, usuario_esperado, resultado, detalhe, observacao) "
                    "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                    (
                        inventario_id, item["id"] if item else None, chave, codigo.strip()[:255],
                        por, datetime.now(), local_id_encontrado, int(usuario_conferido),
                        colaborador_id_encontrado,
                        item["revenda"] if item else None,
                        item["local_id"] if item else None,
                        item["colaborador_id"] if item else None,
                        (esperado.usuario_nome or None) if esperado else None,
                        resultado, (detalhe or "")[:500] or None, observacao,
                    ),
                )
            except pymysql.MySQLError as e:
                if not _erro_de_duplicado(e):
                    raise
                # Defesa: a trava do inventário já serializa, mas a chave única é a garantia final.
                cur.execute(
                    "SELECT id FROM inventario_leituras WHERE inventario_id = %s "
                    "AND (chave = %s OR item_id <=> %s)",
                    (inventario_id, chave, item["id"] if item else None),
                )
                return self._leitura(cur, cur.fetchone()["id"]), True
            return self._leitura(cur, cur.lastrowid), False

    def atualizar_leitura(
        self, inventario_id: int, leitura_id: int, mudancas: dict, por: str
    ) -> tuple[dict, dict]:
        """Corrige local/usuário/observação de uma leitura e reclassifica com a
        fotografia do esperado. Devolve `(antes, depois)`."""
        with get_cursor() as cur:
            inventario = self._travar(cur, inventario_id)
            self._exigir_aberto(inventario, "Este inventário não está aberto: não aceita correções.")
            antes = self._leitura(cur, leitura_id)
            if not antes or antes["inventario_id"] != inventario_id:
                raise LeituraNaoEncontrada()

            local_id = mudancas.get("local_id_encontrado", antes["local_id_encontrado"])
            colaborador_id = mudancas.get(
                "colaborador_id_encontrado", antes["colaborador_id_encontrado"]
            )
            conferido = bool(mudancas.get("usuario_conferido", antes["usuario_conferido"]))
            if "colaborador_id_encontrado" in mudancas and "usuario_conferido" not in mudancas:
                conferido = True
            if "usuario_conferido" in mudancas and not mudancas["usuario_conferido"]:
                colaborador_id = None
            observacao = mudancas.get("observacao", antes["observacao"])

            local_encontrado_nome = self._validar_local(cur, inventario, local_id)
            colaborador_nome = self._validar_colaborador(cur, colaborador_id)

            resultado, detalhe = antes["resultado"], antes["detalhe"]
            if antes["resultado"] != "nao_cadastrado":
                resultado, divergencias = classificar(
                    Esperado(
                        item_id=antes["item_id"],
                        ativo=True,
                        revenda=antes["revenda_item"],
                        local_id=antes["local_id_esperado"],
                        colaborador_id=antes["colaborador_id_esperado"],
                        usuario_nome=antes["usuario_esperado"],
                    ),
                    inventario["unidade_nome"],
                    Encontrado(
                        local_id=local_id,
                        usuario_conferido=conferido,
                        colaborador_id=colaborador_id,
                        colaborador_nome=colaborador_nome,
                    ),
                )
                detalhe = montar_detalhe(
                    divergencias,
                    revenda=antes["revenda_item"],
                    local_esperado_nome=antes["local_esperado_nome"],
                    local_encontrado_nome=local_encontrado_nome,
                    usuario_esperado=antes["usuario_esperado"],
                    usuario_encontrado=colaborador_nome,
                )
            cur.execute(
                "UPDATE inventario_leituras SET local_id_encontrado = %s, usuario_conferido = %s, "
                "colaborador_id_encontrado = %s, observacao = %s, resultado = %s, detalhe = %s, "
                "atualizado_em = %s WHERE id = %s",
                (
                    local_id, int(conferido), colaborador_id, observacao, resultado,
                    (detalhe or "")[:500] or None, datetime.now(), leitura_id,
                ),
            )
            return antes, self._leitura(cur, leitura_id)

    def excluir_leitura(self, inventario_id: int, leitura_id: int) -> dict:
        with get_cursor() as cur:
            inventario = self._travar(cur, inventario_id)
            self._exigir_aberto(inventario, "Este inventário não está aberto: não aceita correções.")
            leitura = self._leitura(cur, leitura_id)
            if not leitura or leitura["inventario_id"] != inventario_id:
                raise LeituraNaoEncontrada()
            cur.execute("DELETE FROM inventario_leituras WHERE id = %s", (leitura_id,))
            return leitura

    def listar_leituras(
        self, inventario_id: int, *, resultado: Optional[str], limit: int, offset: int
    ) -> tuple[list[dict], int]:
        where, params = ["l.inventario_id = %s"], [inventario_id]
        if resultado:
            where.append("l.resultado = %s")
            params.append(resultado)
        where_sql = " AND ".join(where)
        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT COUNT(*) AS total FROM inventario_leituras l WHERE {where_sql}", params)
            total = cur.fetchone()["total"]
            cur.execute(
                f"{self._SELECT_LEITURA} WHERE {where_sql} "
                "ORDER BY l.lido_em DESC, l.id DESC LIMIT %s OFFSET %s",
                [*params, limit, offset],
            )
            return cur.fetchall(), total

    # ── Encerramento, relatório e histórico do item ─────────────────────────

    def encerrar(self, inventario_id: int, por: str) -> None:
        """Congela os não encontrados (itens ativos da unidade sem leitura) e fecha o
        inventário, tudo na mesma transação e com o inventário travado."""
        with get_cursor() as cur:
            inventario = self._travar(cur, inventario_id)
            if inventario["status"] != "aberto":
                raise InventarioFechado("Só inventários abertos podem ser encerrados.")
            cur.execute(
                "INSERT INTO inventario_nao_encontrados "
                "(inventario_id, item_id, status_item, local_id, colaborador_id, usuario) "
                "SELECT %s, i.id, i.status, i.local_id, i.colaborador_id, COALESCE(c.nome, i.assigned_to) "
                "FROM items i LEFT JOIN colaboradores c ON c.id = i.colaborador_id "
                "WHERE i.is_active = 1 AND i.revenda = %s "
                "AND NOT EXISTS (SELECT 1 FROM inventario_leituras l "
                "WHERE l.inventario_id = %s AND l.item_id = i.id)",
                (inventario_id, inventario["unidade_nome"], inventario_id),
            )
            cur.execute(
                "UPDATE inventarios SET status = 'encerrado', encerrado_por = %s, encerrado_em = %s "
                "WHERE id = %s",
                (por, datetime.now(), inventario_id),
            )

    @staticmethod
    def _linha_vazia(categoria: str) -> dict:
        return {
            "categoria": categoria, "leitura_id": None, "item_id": None, "tipo": None,
            "brand": None, "model": None, "identificador": None, "status_item": None,
            "unidade_item": None, "local_esperado": None, "local_encontrado": None,
            "usuario_esperado": None, "usuario_encontrado": None, "codigo_lido": None,
            "lido_por": None, "lido_em": None, "detalhe": None, "observacao": None,
        }

    def linhas_relatorio(self, inventario: dict) -> list[dict]:
        """Uma linha por leitura e uma por não encontrado, na ordem das categorias."""
        inventario_id = inventario["id"]
        status = inventario["status"]
        linhas: list[dict] = []
        with get_cursor(commit=False) as cur:
            cur.execute(
                self._SELECT_LEITURA + " WHERE l.inventario_id = %s", (inventario_id,)
            )
            for leitura in cur.fetchall():
                linha = self._linha_vazia(leitura["resultado"])
                linha.update(
                    leitura_id=leitura["id"], item_id=leitura["item_id"],
                    tipo=leitura["item_tipo"], brand=leitura["item_brand"],
                    model=leitura["item_model"], identificador=leitura["item_identificador"],
                    status_item=leitura["item_status"], unidade_item=leitura["revenda_item"],
                    local_esperado=leitura["local_esperado_nome"],
                    local_encontrado=leitura["local_encontrado_nome"],
                    usuario_esperado=leitura["usuario_esperado"],
                    usuario_encontrado=leitura["colaborador_encontrado_nome"],
                    codigo_lido=leitura["codigo_lido"], lido_por=leitura["lido_por"],
                    lido_em=leitura["lido_em"], detalhe=leitura["detalhe"],
                    observacao=leitura["observacao"],
                )
                linhas.append(linha)

            nao_encontrados: list[dict] = []
            if status == "aberto":
                cur.execute(
                    "SELECT i.id, i.tipo, i.brand, i.model, i.identificador, i.status, "
                    "lo.nome AS local_nome, COALESCE(c.nome, i.assigned_to) AS usuario "
                    "FROM items i LEFT JOIN locais lo ON lo.id = i.local_id "
                    "LEFT JOIN colaboradores c ON c.id = i.colaborador_id "
                    "WHERE i.is_active = 1 AND i.revenda = %s "
                    "AND NOT EXISTS (SELECT 1 FROM inventario_leituras l "
                    "WHERE l.inventario_id = %s AND l.item_id = i.id)",
                    (inventario["unidade_nome"], inventario_id),
                )
                nao_encontrados = cur.fetchall()
            elif status == "encerrado":
                cur.execute(
                    "SELECT i.id, i.tipo, i.brand, i.model, i.identificador, n.status_item AS status, "
                    "lo.nome AS local_nome, n.usuario AS usuario "
                    "FROM inventario_nao_encontrados n JOIN items i ON i.id = n.item_id "
                    "LEFT JOIN locais lo ON lo.id = n.local_id WHERE n.inventario_id = %s",
                    (inventario_id,),
                )
                nao_encontrados = cur.fetchall()
            for item in nao_encontrados:
                linha = self._linha_vazia("nao_encontrado")
                linha.update(
                    item_id=item["id"], tipo=item["tipo"], brand=item["brand"],
                    model=item["model"], identificador=item["identificador"],
                    status_item=item["status"], unidade_item=inventario["unidade_nome"],
                    local_esperado=item["local_nome"], usuario_esperado=item["usuario"],
                )
                linhas.append(linha)

        posicao = {c: i for i, c in enumerate(ORDEM_CATEGORIAS)}
        linhas.sort(
            key=lambda linha: (
                posicao[linha["categoria"]],
                linha["item_id"] is None,
                linha["item_id"] or 0,
                linha["leitura_id"] or 0,
            )
        )
        return linhas

    def historico_do_item(self, item_id: int) -> list[dict]:
        """Participações do item em inventários não cancelados (leitura ou não
        encontrado), da mais antiga para a mais recente. Alimenta a linha do tempo (F7)."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT l.inventario_id, u.nome AS unidade_nome, i.status AS status_inventario, "
                "'leitura' AS evento, l.resultado AS resultado, l.lido_em AS data, "
                "l.lido_por AS por, lo.nome AS local_encontrado_nome "
                "FROM inventario_leituras l JOIN inventarios i ON i.id = l.inventario_id "
                "JOIN unidades u ON u.id = i.unidade_id "
                "LEFT JOIN locais lo ON lo.id = l.local_id_encontrado "
                "WHERE l.item_id = %s AND i.status <> 'cancelado' "
                "UNION ALL "
                "SELECT n.inventario_id, u.nome, i.status, 'nao_encontrado', NULL, "
                "i.encerrado_em, i.encerrado_por, NULL "
                "FROM inventario_nao_encontrados n JOIN inventarios i ON i.id = n.inventario_id "
                "JOIN unidades u ON u.id = i.unidade_id "
                "WHERE n.item_id = %s AND i.status <> 'cancelado' "
                "ORDER BY data, inventario_id",
                (item_id, item_id),
            )
            return cur.fetchall()

    # ── Etiquetas ───────────────────────────────────────────────────────────

    def dados_etiquetas(self, item_ids: list[int]) -> list[dict]:
        """Só o que a etiqueta imprime, dos itens ativos, em ordem crescente de id."""
        if not item_ids:
            return []
        marcadores = ", ".join(["%s"] * len(item_ids))
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT id, tipo, brand, model, identificador, revenda FROM items "
                f"WHERE is_active = 1 AND id IN ({marcadores}) ORDER BY id",
                list(item_ids),
            )
            return cur.fetchall()
