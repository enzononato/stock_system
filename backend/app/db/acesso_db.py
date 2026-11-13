"""
Camada de acesso do módulo de Acessos (F8): catálogo de sistemas, pacotes por cargo e
ficha de acessos do colaborador.

As tabelas vêm da migração 0801 (este módulo não cria schema) e NÃO guardam senha.
Toda escrita grava `audit.record(..., cur=cur)` na mesma transação da alteração.

Comparações sem caixa/acento (nome de sistema, login, cargo) não dependem da collation
do banco: nome e login comparam com `LOWER(...)` na consulta; cargo compara em Python
com `chave_cargo` (as tabelas de pacotes são curtas).
"""
from contextlib import contextmanager
from datetime import datetime
from typing import Optional

import pymysql

from app.db.database_mysql import get_cursor
from app.schemas.acessos import chave_cargo, normalizar_cargo
from app.services import audit
from app.services.audit import AuditContext

STATUS_COM_COLABORADOR = ("Pendente", "Indisponível", "Pendente Devolução")


class AcessoErro(Exception):
    """Erro de domínio do módulo; `status_code` é o HTTP correspondente."""

    status_code = 400


class NaoEncontrado(AcessoErro):
    status_code = 404


class Conflito(AcessoErro):
    status_code = 409


class Invalido(AcessoErro):
    status_code = 422


@contextmanager
def cursor_ou_novo(cur=None, commit: bool = True):
    """Devolve `cur` quando recebido (a transação é de quem chamou); senão abre um novo."""
    if cur is not None:
        yield cur
        return
    with get_cursor(commit=commit) as novo:
        yield novo


def agora() -> datetime:
    return datetime.now()


def _sistema(linha: dict) -> dict:
    return {
        "id": linha["id"],
        "nome": linha["nome"],
        "descricao": linha["descricao"],
        "ativo": bool(linha["ativo"]),
    }


def _nomes_ordenados(sistemas: list[dict]) -> list[str]:
    return sorted((s["nome"] for s in sistemas), key=str.casefold)


class AcessoDBManager:
    """Catálogo, pacotes e ficha. Não cria tabelas no `__init__` (a migração cria)."""

    # ── Sistemas ────────────────────────────────────────────────────────────

    def listar_sistemas(self, *, include_inactive: bool = False) -> list[dict]:
        sql = "SELECT id, nome, descricao, ativo FROM sistemas"
        if not include_inactive:
            sql += " WHERE ativo = 1"
        with get_cursor(commit=False) as cur:
            cur.execute(sql + " ORDER BY nome")
            return [_sistema(l) for l in cur.fetchall()]

    def obter_sistema(self, sistema_id: int, cur=None, *, travar: bool = False) -> Optional[dict]:
        sql = "SELECT id, nome, descricao, ativo FROM sistemas WHERE id = %s"
        if travar:
            sql += " FOR UPDATE"
        with cursor_ou_novo(cur, commit=False) as c:
            c.execute(sql, (sistema_id,))
            linha = c.fetchone()
        return _sistema(linha) if linha else None

    def _nome_em_uso(self, cur, nome: str, ignorar_id: Optional[int] = None) -> bool:
        cur.execute(
            "SELECT id FROM sistemas WHERE LOWER(nome) = LOWER(%s) AND id <> %s LIMIT 1",
            (nome, ignorar_id or 0),
        )
        return cur.fetchone() is not None

    def criar_sistema(self, ctx: AuditContext, *, nome: str, descricao: Optional[str]) -> dict:
        with get_cursor() as cur:
            if self._nome_em_uso(cur, nome):
                raise Conflito("Já existe um sistema com esse nome.")
            try:
                cur.execute(
                    "INSERT INTO sistemas (nome, descricao, ativo) VALUES (%s, %s, 1)",
                    (nome, descricao),
                )
            except pymysql.err.IntegrityError as e:
                if e.args and e.args[0] == 1062:
                    raise Conflito("Já existe um sistema com esse nome.") from e
                raise
            sistema = self.obter_sistema(cur.lastrowid, cur)
            audit.record(ctx, "acesso_sistema", sistema["id"], "criar", depois=sistema, cur=cur)
            return sistema

    def atualizar_sistema(self, ctx: AuditContext, sistema_id: int, dados: dict) -> dict:
        with get_cursor() as cur:
            antes = self.obter_sistema(sistema_id, cur, travar=True)
            if antes is None:
                raise NaoEncontrado("Sistema não encontrado.")
            dados = {k: v for k, v in dados.items() if k in ("nome", "descricao", "ativo")}
            if "nome" in dados and self._nome_em_uso(cur, dados["nome"], sistema_id):
                raise Conflito("Já existe um sistema com esse nome.")
            if dados:
                colunas = ", ".join(f"{k} = %s" for k in dados)
                valores = [int(v) if k == "ativo" else v for k, v in dados.items()]
                try:
                    cur.execute(
                        f"UPDATE sistemas SET {colunas} WHERE id = %s", (*valores, sistema_id)
                    )
                except pymysql.err.IntegrityError as e:
                    if e.args and e.args[0] == 1062:
                        raise Conflito("Já existe um sistema com esse nome.") from e
                    raise
            depois = self.obter_sistema(sistema_id, cur)
            if depois != antes:
                if antes["ativo"] != depois["ativo"]:
                    acao = "reativar" if depois["ativo"] else "desativar"
                else:
                    acao = "editar"
                a, d = audit.diff(antes, depois)
                audit.record(ctx, "acesso_sistema", sistema_id, acao, antes=a, depois=d, cur=cur)
            return depois

    # ── Pacotes por cargo ───────────────────────────────────────────────────

    def _linhas_de_pacotes(self, cur) -> list[dict]:
        cur.execute(
            "SELECT p.id AS pacote_id, p.cargo, s.id, s.nome, s.descricao, s.ativo "
            "FROM pacotes_funcao p JOIN sistemas s ON s.id = p.sistema_id "
            "ORDER BY p.cargo, s.nome"
        )
        return cur.fetchall()

    def listar_pacotes(self) -> list[dict]:
        with get_cursor(commit=False) as cur:
            linhas = self._linhas_de_pacotes(cur)
        grupos: dict[str, dict] = {}
        for linha in linhas:
            grupo = grupos.setdefault(
                chave_cargo(linha["cargo"]), {"cargo": normalizar_cargo(linha["cargo"]), "sistemas": []}
            )
            grupo["sistemas"].append(_sistema(linha))
        for grupo in grupos.values():
            grupo["sistemas"].sort(key=lambda s: s["nome"].casefold())
        return sorted(grupos.values(), key=lambda g: chave_cargo(g["cargo"]))

    def sistemas_do_pacote(
        self, cargo: str, *, somente_ativos: bool = True, cur=None
    ) -> list[dict]:
        chave = chave_cargo(cargo)
        if not chave:
            return []
        with cursor_ou_novo(cur, commit=False) as c:
            linhas = self._linhas_de_pacotes(c)
        sistemas = [
            _sistema(l) for l in linhas
            if chave_cargo(l["cargo"]) == chave and (l["ativo"] or not somente_ativos)
        ]
        return sorted(sistemas, key=lambda s: s["nome"].casefold())

    def definir_pacote(
        self, ctx: AuditContext, *, cargo: str, sistema_ids: list[int]
    ) -> dict:
        cargo = normalizar_cargo(cargo)
        ids = list(dict.fromkeys(sistema_ids))
        with get_cursor() as cur:
            novos: list[dict] = []
            for sistema_id in ids:
                sistema = self.obter_sistema(sistema_id, cur, travar=True)
                if sistema is None or not sistema["ativo"]:
                    raise Invalido("Sistema inválido ou desativado.")
                novos.append(sistema)

            atuais = [
                l for l in self._linhas_de_pacotes_travadas(cur)
                if chave_cargo(l["cargo"]) == chave_cargo(cargo)
            ]
            antes = {"cargo": cargo, "sistemas": _nomes_ordenados([_sistema(l) for l in atuais])}
            if atuais:
                cur.execute(
                    "DELETE FROM pacotes_funcao WHERE id IN %s", (tuple(l["pacote_id"] for l in atuais),)
                )
            for sistema in novos:
                cur.execute(
                    "INSERT INTO pacotes_funcao (cargo, sistema_id) VALUES (%s, %s)",
                    (cargo, sistema["id"]),
                )
            depois = {"cargo": cargo, "sistemas": _nomes_ordenados(novos)}
            audit.record(ctx, "acesso_pacote", None, "editar", antes=antes, depois=depois, cur=cur)
            return {
                "cargo": cargo,
                "sistemas": sorted(novos, key=lambda s: s["nome"].casefold()),
            }

    def _linhas_de_pacotes_travadas(self, cur) -> list[dict]:
        cur.execute(
            "SELECT p.id AS pacote_id, p.cargo, s.id, s.nome, s.descricao, s.ativo "
            "FROM pacotes_funcao p JOIN sistemas s ON s.id = p.sistema_id "
            "ORDER BY p.cargo, s.nome FOR UPDATE"
        )
        return cur.fetchall()
