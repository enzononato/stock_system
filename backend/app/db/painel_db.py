"""Consultas do painel (F7): indicadores e dados crus dos alertas. Só leitura, SQL cru."""
from datetime import datetime

from app.core.config import ITEM_STATUSES
from app.db.database_mysql import get_cursor
from app.db.linha_db import LinhaDBManager
from app.db.manutencao_db import STATUS_EM_ANDAMENTO
from app.db.termo_db import PENDENTES

_PENDENTES_SQL = ", ".join(f"'{s}'" for s in PENDENTES)
_EM_ANDAMENTO_SQL = ", ".join(f"'{s}'" for s in STATUS_EM_ANDAMENTO)


class PainelDBManager:
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
            cur.execute(
                f"SELECT COUNT(*) AS n FROM manutencoes WHERE status IN ({_EM_ANDAMENTO_SQL})"
            )
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
