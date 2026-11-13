"""Consultas do painel (F7): indicadores e dados crus dos alertas. Só leitura, SQL cru."""
from datetime import date, datetime

from app.core.config import ITEM_STATUSES
from app.db.acesso_db import STATUS_COM_COLABORADOR
from app.db.database_mysql import get_cursor
from app.db.linha_db import LinhaDBManager
from app.db.manutencao_db import STATUS_EM_ANDAMENTO
from app.db.termo_db import PENDENTES


def _escapar_like(texto: str) -> str:
    return texto.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def _lista_sql(valores) -> str:
    return ", ".join(f"'{v}'" for v in valores)


_PENDENTES_SQL = _lista_sql(PENDENTES)
_EM_ANDAMENTO_SQL = _lista_sql(STATUS_EM_ANDAMENTO)

_COLUNAS_ITEM_TERMO = (
    "i.id AS item_id, i.tipo AS item_tipo, i.brand AS item_brand, i.model AS item_model, "
    "i.identificador AS item_identificador, i.revenda AS item_revenda, "
    "i.colaborador_id AS colaborador_id, COALESCE(c.nome, i.assigned_to) AS colaborador_nome"
)


class PainelDBManager:
    # ── Indicadores ─────────────────────────────────────────────────────────

    def contagem_itens(self) -> list[dict]:
        """Itens ativos por unidade (`None` = sem unidade) e status."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT NULLIF(TRIM(revenda), '') AS unidade, status, COUNT(*) AS n "
                "FROM items WHERE is_active=1 GROUP BY 1, 2"
            )
            return list(cur.fetchall())

    def ordens_em_andamento(self) -> int:
        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT COUNT(*) AS n FROM manutencoes WHERE status IN ({_EM_ANDAMENTO_SQL})")
            return int(cur.fetchone()["n"])

    def termos_por_status(self) -> dict[str, int]:
        """Solicitações pendentes de itens ativos; as 4 chaves de `PENDENTES` sempre presentes."""
        contagem = {s: 0 for s in PENDENTES}
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT t.status, COUNT(*) AS n FROM termo_solicitacoes t "
                "JOIN items i ON i.id = t.item_id "
                f"WHERE i.is_active=1 AND t.status IN ({_PENDENTES_SQL}) GROUP BY t.status"
            )
            for linha in cur.fetchall():
                contagem[linha["status"]] = int(linha["n"])
        return contagem

    def legados_sem_solicitacao(self) -> int:
        """Itens ativos `Pendente` sem nenhuma solicitação pendente (empréstimos anteriores à F3)."""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT COUNT(*) AS n FROM items i WHERE i.is_active=1 AND i.status='Pendente' "
                "AND NOT EXISTS (SELECT 1 FROM termo_solicitacoes t WHERE t.item_id = i.id "
                f"AND t.status IN ({_PENDENTES_SQL}))"
            )
            return int(cur.fetchone()["n"])

    def resumo(self) -> dict:
        por_status = {s: 0 for s in ITEM_STATUSES}
        unidades: dict[str, dict] = {}
        total = 0
        for linha in self.contagem_itens():
            n = int(linha["n"])
            total += n
            if linha["status"] in por_status:
                por_status[linha["status"]] += n
            nome = linha["unidade"]
            chave = nome.casefold() if nome else ""
            u = unidades.setdefault(
                chave, {"unidade": nome, "total": 0, "por_status": {s: 0 for s in ITEM_STATUSES}}
            )
            u["total"] += n
            if linha["status"] in u["por_status"]:
                u["por_status"][linha["status"]] += n
        por_unidade = sorted(
            unidades.values(), key=lambda u: (u["unidade"] is None, (u["unidade"] or "").casefold())
        )
        termos = self.termos_por_status()
        custo = LinhaDBManager().relatorio_custo()
        return {
            "gerado_em": datetime.now(),
            "equipamentos": {"total": total, "por_status": por_status, "por_unidade": por_unidade},
            "manutencao": {
                "equipamentos_em_manutencao": por_status["Em Manutenção"],
                "ordens_em_andamento": self.ordens_em_andamento(),
            },
            "termos": {
                "pendentes": sum(termos.values()),
                "por_status": termos,
                "sem_solicitacao": self.legados_sem_solicitacao(),
            },
            "linhas": {
                "ativas": custo["total_linhas"],
                "custo_mensal": custo["custo_total"],
                "por_unidade": [
                    {
                        "unidade_nome": u["unidade_nome"],
                        "quantidade": u["quantidade"],
                        "custo_total": u["custo_total"],
                    }
                    for u in custo["unidades"]
                ],
            },
        }

    # ── Dados crus dos alertas (o serviço aplica os limiares) ───────────────

    @staticmethod
    def _filtro_busca(busca: str | None) -> tuple[str, list]:
        texto = (busca or "").strip()
        if not texto:
            return "", []
        padrao = f"%{_escapar_like(texto)}%"
        sql = (
            " AND (i.brand LIKE %s OR i.model LIKE %s OR i.identificador LIKE %s "
            "OR c.nome LIKE %s OR i.assigned_to LIKE %s)"
        )
        return sql, [padrao] * 5

    def termos_pendentes(
        self, *, busca: str | None = None, limit: int | None = None
    ) -> tuple[list[dict], int]:
        """Solicitações pendentes de itens ativos mais os itens `Pendente` legados (sem solicitação
        pendente, `legado=True`). Ordem: `referencia` crescente (a mais antiga primeiro)."""
        filtro, params_busca = self._filtro_busca(busca)
        solicitacoes = (
            f"SELECT t.id AS termo_id, {_COLUNAS_ITEM_TERMO}, t.tipo AS tipo, t.status AS status, "
            "t.solicitado_por AS solicitado_por, t.solicitado_em AS solicitado_em, "
            "t.enviado_em AS enviado_em, t.agendado_para AS agendado_para, t.motivo AS motivo, "
            "CASE t.status WHEN 'agendado' THEN COALESCE(t.agendado_para, t.solicitado_em) "
            "WHEN 'enviado' THEN COALESCE(t.enviado_em, t.solicitado_em) "
            "ELSE t.solicitado_em END AS referencia, 0 AS legado "
            "FROM termo_solicitacoes t JOIN items i ON i.id = t.item_id "
            "LEFT JOIN colaboradores c ON c.id = t.colaborador_id "
            f"WHERE i.is_active = 1 AND t.status IN ({_PENDENTES_SQL})" + filtro
        )
        legados = (
            f"SELECT NULL, {_COLUNAS_ITEM_TERMO}, 'entrega', 'sem_solicitacao', NULL, NULL, NULL, "
            "NULL, NULL, COALESCE(h.data, i.date_issued, i.date_registered), 1 "
            "FROM items i LEFT JOIN colaboradores c ON c.id = i.colaborador_id "
            "LEFT JOIN (SELECT item_id, MAX(data_operacao) AS data FROM history "
            "WHERE operation = 'Empréstimo' AND is_reversed = 0 GROUP BY item_id) h "
            "ON h.item_id = i.id "
            "WHERE i.is_active = 1 AND i.status = 'Pendente' AND NOT EXISTS "
            "(SELECT 1 FROM termo_solicitacoes t2 WHERE t2.item_id = i.id "
            f"AND t2.status IN ({_PENDENTES_SQL}))" + filtro
        )
        uniao = f"({solicitacoes}) UNION ALL ({legados})"
        params = params_busca + params_busca
        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT COUNT(*) AS n FROM ({uniao}) u", params)
            total = int(cur.fetchone()["n"])
            sql = f"SELECT * FROM ({uniao}) u ORDER BY referencia ASC, item_id ASC, termo_id ASC"
            params_lista = list(params)
            if limit is not None:
                sql += " LIMIT %s"
                params_lista.append(limit)
            cur.execute(sql, params_lista)
            linhas = list(cur.fetchall())
        for linha in linhas:
            linha["legado"] = bool(linha["legado"])
        return linhas, total

    def garantias_entre(self, inicio: date, fim: date) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT id, tipo, brand, model, identificador, revenda, garantia_ate FROM items "
                "WHERE is_active = 1 AND garantia_ate BETWEEN %s AND %s ORDER BY garantia_ate, id",
                (inicio, fim),
            )
            return list(cur.fetchall())

    def devolucoes_pendentes(self) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT i.id, i.tipo, i.brand, i.model, i.identificador, i.revenda, i.assigned_to, "
                "i.colaborador_id, c.nome AS colaborador_nome, h.desde AS desde "
                "FROM items i LEFT JOIN colaboradores c ON c.id = i.colaborador_id "
                "LEFT JOIN (SELECT item_id, MAX(data_operacao) AS desde FROM history "
                "WHERE operation = 'Devolução' AND is_reversed = 0 GROUP BY item_id) h "
                "ON h.item_id = i.id "
                "WHERE i.is_active = 1 AND i.status = 'Pendente Devolução' ORDER BY i.id"
            )
            return list(cur.fetchall())

    def manutencoes_em_andamento(self) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT m.id, m.item_id, m.status, m.assistencia, "
                "GREATEST(m.aberta_em, COALESCE(m.enviada_em, m.aberta_em), "
                "COALESCE(m.em_reparo_em, m.aberta_em)) AS ultimo_movimento, "
                "i.tipo, i.brand, i.model, i.identificador, i.revenda "
                "FROM manutencoes m JOIN items i ON i.id = m.item_id "
                f"WHERE m.status IN ({_EM_ANDAMENTO_SQL}) ORDER BY m.id"
            )
            return list(cur.fetchall())

    def desligados_com_equipamento(self) -> list[dict]:
        """Itens ativos com colaborador desligado: pelo vínculo (`colaborador_id`) ou, nos
        registros sem vínculo, pelo CPF (só dígitos). Mesma regra de
        `AcessoDBManager.equipamentos_com_colaborador`, numa consulta só."""
        colunas = (
            "i.id AS item_id, i.tipo, i.brand, i.model, i.identificador, i.status, i.revenda, "
            "c.id AS colaborador_id, c.nome AS colaborador_nome"
        )
        base = (
            f"SELECT {colunas} FROM items i JOIN colaboradores c ON {{juncao}} "
            "WHERE c.situacao = 'desligado' AND i.is_active = 1 "
            f"AND i.status IN ({_lista_sql(STATUS_COM_COLABORADOR)})"
        )
        por_vinculo = base.format(juncao="c.id = i.colaborador_id")
        por_cpf = base.format(
            juncao=(
                "i.colaborador_id IS NULL AND i.cpf IS NOT NULL AND i.cpf <> '' AND c.cpf = "
                "REPLACE(REPLACE(REPLACE(REPLACE(i.cpf, '.', ''), '-', ''), ' ', ''), '/', '')"
            )
        )
        with get_cursor(commit=False) as cur:
            cur.execute(f"({por_vinculo}) UNION ({por_cpf}) ORDER BY item_id")
            return list(cur.fetchall())

    def linhas_em_aparelho_baixado(self) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT l.id AS linha_id, l.numero, u.nome AS unidade_nome, i.id AS item_id, "
                "i.tipo, i.brand, i.model, i.identificador "
                "FROM linhas l JOIN items i ON i.id = l.item_id JOIN unidades u ON u.id = l.unidade_id "
                "WHERE l.status = 'ativa' AND i.is_active = 0 ORDER BY l.id"
            )
            return list(cur.fetchall())
