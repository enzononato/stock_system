"""
Camada de acesso do módulo de Acessos (F8): catálogo de sistemas, pacotes por cargo e
ficha de acessos do colaborador.

As tabelas vêm da migração 0801 (este módulo não cria schema) e NÃO guardam senha.
Toda escrita grava `audit.record(..., cur=cur)` na mesma transação da alteração.

Comparações sem caixa/acento (nome de sistema, login, cargo) não dependem da collation
do banco: nome e login comparam com `LOWER(...)` na consulta; cargo compara em Python
com `chave_cargo` (as tabelas de pacotes são curtas).
"""
import functools
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


MSG_DEADLOCK = "Outra alteração estava em andamento. Tente novamente."


def mapear_deadlock(funcao):
    """Deadlock do MySQL (1213) vira 409 com mensagem em português, em vez de 500.
    O rollback já aconteceu quando a exceção sai do `get_cursor` da própria função.

    Ordem GLOBAL das travas neste módulo (todo caminho de escrita a segue, para não haver
    deadlock): tarefa -> evento -> colaborador -> acessos (`colaborador_acessos`)."""

    @functools.wraps(funcao)
    def interna(*args, **kwargs):
        try:
            return funcao(*args, **kwargs)
        except pymysql.err.OperationalError as e:
            if e.args and e.args[0] == 1213:
                raise Conflito(MSG_DEADLOCK) from e
            raise

    return interna


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

    @mapear_deadlock
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

    @mapear_deadlock
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

    @mapear_deadlock
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

    # ── Ficha de acessos do colaborador ─────────────────────────────────────

    _SELECT_ACESSO = (
        "SELECT a.id, a.colaborador_id, c.nome AS colaborador_nome, a.sistema_id, "
        "s.nome AS sistema_nome, a.login, a.status, a.item_id, a.criado_por, a.criado_em, "
        "a.desativado_por, a.desativado_em, "
        "i.tipo AS item_tipo, i.brand AS item_brand, i.model AS item_model, "
        "i.identificador AS item_identificador "
        "FROM colaborador_acessos a "
        "JOIN sistemas s ON s.id = a.sistema_id "
        "JOIN colaboradores c ON c.id = a.colaborador_id "
        "LEFT JOIN items i ON i.id = a.item_id"
    )

    @staticmethod
    def _acesso(linha: dict) -> dict:
        saida = {k: v for k, v in linha.items() if not k.startswith("item_") or k == "item_id"}
        descricao = None
        if linha["item_id"] is not None:
            descricao = descrever_item(
                {
                    "tipo": linha["item_tipo"],
                    "brand": linha["item_brand"],
                    "model": linha["item_model"],
                    "identificador": linha["item_identificador"],
                }
            )
        saida["item_descricao"] = descricao
        return saida

    def listar_acessos(
        self,
        *,
        colaborador_id: Optional[int] = None,
        item_id: Optional[int] = None,
        status: Optional[str] = None,
        cur=None,
    ) -> list[dict]:
        condicoes, params = [], []
        if colaborador_id is not None:
            condicoes.append("a.colaborador_id = %s")
            params.append(colaborador_id)
        if item_id is not None:
            condicoes.append("a.item_id = %s")
            params.append(item_id)
        if status is not None:
            condicoes.append("a.status = %s")
            params.append(status)
        sql = self._SELECT_ACESSO
        if condicoes:
            sql += " WHERE " + " AND ".join(condicoes)
        sql += " ORDER BY a.status = 'ativo' DESC, s.nome, a.login"
        with cursor_ou_novo(cur, commit=False) as c:
            c.execute(sql, tuple(params))
            return [self._acesso(l) for l in c.fetchall()]

    def obter_acesso(self, acesso_id: int, cur=None) -> Optional[dict]:
        with cursor_ou_novo(cur, commit=False) as c:
            c.execute(self._SELECT_ACESSO + " WHERE a.id = %s", (acesso_id,))
            linha = c.fetchone()
        return self._acesso(linha) if linha else None

    @staticmethod
    def _foto(acesso: dict) -> dict:
        """Estado do acesso para a auditoria (nunca há senha: ela não existe no módulo)."""
        return {
            "colaborador_id": acesso["colaborador_id"],
            "sistema": acesso["sistema_nome"],
            "login": acesso["login"],
            "status": acesso["status"],
            "item_id": acesso["item_id"],
        }

    @staticmethod
    def _exigir_item(cur, item_id: int) -> None:
        cur.execute("SELECT id FROM items WHERE id = %s AND is_active = 1", (item_id,))
        if cur.fetchone() is None:
            raise NaoEncontrado("Equipamento não encontrado.")

    @staticmethod
    def _login_em_uso(cur, colaborador_id: int, sistema_id: int, login: str,
                      ignorar_id: Optional[int] = None) -> Optional[dict]:
        cur.execute(
            "SELECT id, status FROM colaborador_acessos "
            "WHERE colaborador_id = %s AND sistema_id = %s AND LOWER(login) = LOWER(%s) "
            "AND id <> %s ORDER BY status = 'ativo' DESC LIMIT 1 FOR UPDATE",
            (colaborador_id, sistema_id, login, ignorar_id or 0),
        )
        return cur.fetchone()

    @mapear_deadlock
    def criar_acesso(
        self,
        ctx: AuditContext,
        *,
        colaborador_id: int,
        sistema_id: int,
        login: str,
        item_id: Optional[int] = None,
        cur=None,
    ) -> dict:
        login = " ".join(login.split())
        with cursor_ou_novo(cur) as c:
            # Travar o colaborador serializa os registros de acesso dele (sem deadlock de gap lock).
            if obter_colaborador(c, colaborador_id, travar=True) is None:
                raise NaoEncontrado("Colaborador não encontrado.")
            sistema = self.obter_sistema(sistema_id, c)
            if sistema is None:
                raise NaoEncontrado("Sistema não encontrado.")
            if not sistema["ativo"]:
                raise Invalido("Sistema desativado: não é possível registrar acesso nele.")
            if item_id is not None:
                self._exigir_item(c, item_id)

            existente = self._login_em_uso(c, colaborador_id, sistema_id, login)
            if existente is not None and existente["status"] == "ativo":
                raise Conflito("Este login já está ativo para o colaborador neste sistema.")
            if existente is not None:
                antes = self.obter_acesso(existente["id"], c)
                c.execute(
                    "UPDATE colaborador_acessos SET login = %s, status = 'ativo', item_id = %s, "
                    "criado_por = %s, criado_em = %s, desativado_por = NULL, desativado_em = NULL "
                    "WHERE id = %s",
                    (login, item_id, ctx.username, agora(), existente["id"]),
                )
                acesso = self.obter_acesso(existente["id"], c)
                a, d = audit.diff(self._foto(antes), self._foto(acesso))
                audit.record(ctx, "acesso", acesso["id"], "reativar", antes=a, depois=d, cur=c)
                return acesso

            try:
                c.execute(
                    "INSERT INTO colaborador_acessos (colaborador_id, sistema_id, login, status, "
                    "item_id, criado_por, criado_em) VALUES (%s, %s, %s, 'ativo', %s, %s, %s)",
                    (colaborador_id, sistema_id, login, item_id, ctx.username, agora()),
                )
            except pymysql.err.IntegrityError as e:
                if e.args and e.args[0] == 1062:
                    raise Conflito(
                        "Este login já está ativo para o colaborador neste sistema."
                    ) from e
                raise
            acesso = self.obter_acesso(c.lastrowid, c)
            audit.record(ctx, "acesso", acesso["id"], "criar", depois=self._foto(acesso), cur=c)
            return acesso

    @mapear_deadlock
    def atualizar_acesso(self, ctx: AuditContext, acesso_id: int, dados: dict) -> dict:
        with get_cursor() as cur:
            cur.execute(
                "SELECT id, colaborador_id, sistema_id FROM colaborador_acessos "
                "WHERE id = %s FOR UPDATE",
                (acesso_id,),
            )
            base = cur.fetchone()
            if base is None:
                raise NaoEncontrado("Acesso não encontrado.")
            antes = self.obter_acesso(acesso_id, cur)
            dados = {k: v for k, v in dados.items() if k in ("login", "item_id")}
            if "login" in dados and dados["login"] is None:
                del dados["login"]  # login não é anulável: null equivale a campo ausente
            if "login" in dados and self._login_em_uso(
                cur, base["colaborador_id"], base["sistema_id"], dados["login"], acesso_id
            ):
                raise Conflito("Já existe um acesso com esse login para o colaborador neste sistema.")
            if dados.get("item_id") is not None:
                self._exigir_item(cur, dados["item_id"])
            if dados:
                colunas = ", ".join(f"{k} = %s" for k in dados)
                try:
                    cur.execute(
                        f"UPDATE colaborador_acessos SET {colunas} WHERE id = %s",
                        (*dados.values(), acesso_id),
                    )
                except pymysql.err.IntegrityError as e:
                    if e.args and e.args[0] == 1062:
                        raise Conflito(
                            "Já existe um acesso com esse login para o colaborador neste sistema."
                        ) from e
                    raise
            depois = self.obter_acesso(acesso_id, cur)
            a, d = audit.diff(self._foto(antes), self._foto(depois))
            if a or d:
                audit.record(ctx, "acesso", acesso_id, "editar", antes=a, depois=d, cur=cur)
            return depois

    @mapear_deadlock
    def desativar_acesso(self, ctx: AuditContext, acesso_id: int, cur=None) -> dict:
        with cursor_ou_novo(cur) as c:
            c.execute(
                "SELECT status FROM colaborador_acessos WHERE id = %s FOR UPDATE", (acesso_id,)
            )
            linha = c.fetchone()
            if linha is None:
                raise NaoEncontrado("Acesso não encontrado.")
            if linha["status"] != "ativo":
                raise Conflito("O acesso já está desativado.")
            antes = self.obter_acesso(acesso_id, c)
            c.execute(
                "UPDATE colaborador_acessos SET status = 'desativado', desativado_por = %s, "
                "desativado_em = %s WHERE id = %s",
                (ctx.username, agora(), acesso_id),
            )
            depois = self.obter_acesso(acesso_id, c)
            a, d = audit.diff(self._foto(antes), self._foto(depois))
            audit.record(ctx, "acesso", acesso_id, "desativar", antes=a, depois=d, cur=c)
            return depois

    # ── Equipamentos com o colaborador ──────────────────────────────────────

    def equipamentos_com_colaborador(self, colaborador_id: int, cur=None) -> list[dict]:
        """Itens ativos que estão com o colaborador: pelo vínculo (`colaborador_id`) ou, nos
        registros anteriores à F0 (sem vínculo), pelo CPF -- com ou sem máscara. O filtro de
        status protege contra `colaborador_id` esquecido num item já devolvido."""
        base = (
            "SELECT i.id, i.tipo, i.brand, i.model, i.identificador, i.status, i.revenda, "
            "i.cpf, i.colaborador_id FROM items i "
            "WHERE i.is_active = 1 AND i.status IN %s AND "
        )
        with cursor_ou_novo(cur, commit=False) as c:
            colaborador = obter_colaborador(c, colaborador_id)
            if colaborador is None:
                return []
            digitos = "".join(ch for ch in (colaborador["cpf"] or "") if ch.isdigit())
            if len(digitos) == 11:
                c.execute(
                    base + "(i.colaborador_id = %s OR (i.colaborador_id IS NULL "
                    "AND i.cpf IS NOT NULL AND i.cpf <> '')) ORDER BY i.id",
                    (STATUS_COM_COLABORADOR, colaborador_id),
                )
            else:
                c.execute(
                    base + "i.colaborador_id = %s ORDER BY i.id",
                    (STATUS_COM_COLABORADOR, colaborador_id),
                )
            linhas = c.fetchall()
        saida = []
        for linha in linhas:
            if linha["colaborador_id"] is None:
                cpf_item = "".join(ch for ch in (linha["cpf"] or "") if ch.isdigit())
                if cpf_item != digitos:
                    continue
            saida.append({k: linha[k] for k in
                          ("id", "tipo", "brand", "model", "identificador", "status", "revenda")})
        return saida


def descrever_item(linha: dict) -> Optional[str]:
    """`"Notebook Dell Latitude 5420 · SN-AB12"`; tudo vazio vira `None`."""
    partes = [str(linha[k]).strip() for k in ("tipo", "brand", "model") if linha.get(k)]
    texto = " ".join(p for p in partes if p)
    identificador = (linha.get("identificador") or "").strip()
    if identificador:
        texto = f"{texto} · {identificador}" if texto else identificador
    return texto or None


def obter_colaborador(cur, colaborador_id: int, *, travar: bool = False) -> Optional[dict]:
    sql = "SELECT id, nome, cpf, cargo, situacao FROM colaboradores WHERE id = %s"
    if travar:
        sql += " FOR UPDATE"
    cur.execute(sql, (colaborador_id,))
    return cur.fetchone()
