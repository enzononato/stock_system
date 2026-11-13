"""
Eventos de acesso (F8): admissão, desligamento e mudança de função registrados à mão,
cada um com uma lista de tarefas (criar/desativar acesso em cada sistema).

Reaproveita `AcessoDBManager` (catálogo, pacotes, ficha) no mesmo cursor/transação. Toda
escrita grava `audit.record(..., cur=cur)`. Nenhuma coluna guarda senha.
"""
import re
from typing import Optional

from app.db.acesso_db import (
    AcessoDBManager,
    Conflito,
    Invalido,
    NaoEncontrado,
    agora,
    cursor_ou_novo,
    obter_colaborador,
)
from app.db.database_mysql import get_cursor
from app.schemas.acessos import chave_cargo, normalizar_cargo
from app.services import audit
from app.services.audit import AuditContext


def diferenca_pacotes(anterior: list[dict], novo: list[dict]) -> tuple[list[dict], list[dict]]:
    """`(entram, saem)`: sistemas só do pacote novo e só do anterior (por `id`), por nome."""
    ids_anterior = {s["id"] for s in anterior}
    ids_novo = {s["id"] for s in novo}
    entram = [s for s in novo if s["id"] not in ids_anterior]
    saem = [s for s in anterior if s["id"] not in ids_novo]
    ordem = lambda s: s["nome"].casefold()  # noqa: E731
    return sorted(entram, key=ordem), sorted(saem, key=ordem)


def pendencias_do_evento(tipo: str, tarefas_pendentes: int, equipamentos: int) -> list[str]:
    """Motivos pelos quais o evento ainda não pode ser concluído. Equipamentos só contam no
    desligamento."""
    pendencias = []
    if tarefas_pendentes:
        pendencias.append(
            "1 tarefa pendente" if tarefas_pendentes == 1 else f"{tarefas_pendentes} tarefas pendentes"
        )
    if tipo == "desligamento" and equipamentos:
        pendencias.append(
            "1 equipamento ainda com o colaborador"
            if equipamentos == 1
            else f"{equipamentos} equipamentos ainda com o colaborador"
        )
    return pendencias


def _escapar_like(texto: str) -> str:
    return texto.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


_SELECT_EVENTO = (
    "SELECT e.id, e.colaborador_id, c.nome AS colaborador_nome, c.situacao AS colaborador_situacao, "
    "e.tipo, e.cargo_anterior, e.cargo_novo, e.status, e.observacao, e.aberto_por, e.aberto_em, "
    "e.concluido_por, e.concluido_em, e.cancelado_por, e.cancelado_em, e.motivo_cancelamento, "
    "(SELECT COUNT(*) FROM acesso_tarefas t WHERE t.evento_id = e.id) AS total_tarefas, "
    "(SELECT COUNT(*) FROM acesso_tarefas t WHERE t.evento_id = e.id AND t.status = 'pendente') "
    "AS tarefas_pendentes "
    "FROM acesso_eventos e JOIN colaboradores c ON c.id = e.colaborador_id"
)


class AcessoEventoDBManager:
    def __init__(self, acessos: AcessoDBManager):
        self.acessos = acessos

    # ── Consulta ────────────────────────────────────────────────────────────

    def listar(
        self,
        *,
        status: Optional[str] = None,
        tipo: Optional[str] = None,
        colaborador_id: Optional[int] = None,
        busca: Optional[str] = None,
        limit: int = 25,
        offset: int = 0,
    ) -> tuple[list[dict], int]:
        condicoes, params = [], []
        if status:
            condicoes.append("e.status = %s")
            params.append(status)
        if tipo:
            condicoes.append("e.tipo = %s")
            params.append(tipo)
        if colaborador_id is not None:
            condicoes.append("e.colaborador_id = %s")
            params.append(colaborador_id)
        busca = (busca or "").strip()
        if busca:
            parte = "c.nome LIKE %s"
            params.append(f"%{_escapar_like(busca)}%")
            digitos = "".join(ch for ch in busca if ch.isdigit())
            # CPF só entra quando o termo é número/pontuação (com ≥ 3 dígitos): "silva 7" ou "7"
            # não devem listar todo CPF que contém esse dígito.
            if len(digitos) >= 3 and re.fullmatch(r"[\d.\-\s/]+", busca):
                parte += " OR c.cpf LIKE %s"
                params.append(f"%{digitos}%")
            condicoes.append(f"({parte})")
        where = (" WHERE " + " AND ".join(condicoes)) if condicoes else ""
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT COUNT(*) AS n FROM acesso_eventos e "
                "JOIN colaboradores c ON c.id = e.colaborador_id" + where,
                tuple(params),
            )
            total = cur.fetchone()["n"]
            cur.execute(
                _SELECT_EVENTO + where + " ORDER BY e.aberto_em DESC, e.id DESC LIMIT %s OFFSET %s",
                (*params, limit, offset),
            )
            linhas = cur.fetchall()
        for linha in linhas:
            linha.pop("colaborador_situacao", None)
        return linhas, total

    def obter(self, evento_id: int, cur=None) -> Optional[dict]:
        with cursor_ou_novo(cur, commit=False) as c:
            c.execute(_SELECT_EVENTO + " WHERE e.id = %s", (evento_id,))
            evento = c.fetchone()
            if evento is None:
                return None
            c.execute(
                "SELECT t.id, t.evento_id, t.sistema_id, s.nome AS sistema_nome, t.acao, t.status, "
                "t.login, t.acesso_id, t.feito_por, t.feito_em, t.observacao "
                "FROM acesso_tarefas t JOIN sistemas s ON s.id = t.sistema_id "
                "WHERE t.evento_id = %s ORDER BY t.acao, s.nome, t.id",
                (evento_id,),
            )
            evento["tarefas"] = c.fetchall()
            evento["equipamentos"] = (
                self.acessos.equipamentos_com_colaborador(evento["colaborador_id"], cur=c)
                if evento["tipo"] == "desligamento"
                else []
            )
        aberto = evento["status"] == "aberto"
        evento["pendencias"] = (
            pendencias_do_evento(
                evento["tipo"], evento["tarefas_pendentes"], len(evento["equipamentos"])
            )
            if aberto
            else []
        )
        evento["pode_concluir"] = aberto and not evento["pendencias"]
        return evento

    # ── Abertura ────────────────────────────────────────────────────────────

    @staticmethod
    def _inserir_tarefa(
        cur,
        evento_id: int,
        sistema_id: int,
        acao: str,
        *,
        acesso_id: Optional[int] = None,
        login: Optional[str] = None,
        resolvida_por: Optional[str] = None,
        observacao: Optional[str] = None,
    ) -> None:
        resolvida = resolvida_por is not None
        cur.execute(
            "INSERT INTO acesso_tarefas (evento_id, sistema_id, acesso_id, acao, status, login, "
            "feito_por, feito_em, observacao) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)",
            (
                evento_id, sistema_id, acesso_id, acao,
                "nao_se_aplica" if resolvida else "pendente",
                login, resolvida_por, agora() if resolvida else None, observacao,
            ),
        )

    def _cargos_do_evento(
        self, tipo: str, colaborador: dict, cargo_anterior: Optional[str], cargo_novo: Optional[str]
    ) -> tuple[Optional[str], Optional[str]]:
        cargo_atual = normalizar_cargo(colaborador["cargo"])
        if tipo == "admissao":
            cargo_novo = normalizar_cargo(cargo_novo) or cargo_atual
            if not cargo_novo:
                raise Invalido("Informe o cargo: o colaborador não tem cargo cadastrado.")
            return None, cargo_novo
        if tipo == "mudanca_funcao":
            cargo_anterior = normalizar_cargo(cargo_anterior)
            if not cargo_anterior:
                raise Invalido("Informe o cargo anterior.")
            cargo_novo = normalizar_cargo(cargo_novo) or cargo_atual
            if not cargo_novo:
                raise Invalido("Informe o cargo novo.")
            if chave_cargo(cargo_anterior) == chave_cargo(cargo_novo):
                raise Invalido("O cargo novo é igual ao anterior.")
            return cargo_anterior, cargo_novo
        return cargo_atual, None  # desligamento

    def abrir(
        self,
        ctx: AuditContext,
        *,
        colaborador_id: int,
        tipo: str,
        cargo_anterior: Optional[str],
        cargo_novo: Optional[str],
        observacao: Optional[str],
    ) -> dict:
        with get_cursor() as cur:
            # "Um evento aberto por colaborador" depende da ORDEM destas instruções: a trava do
            # colaborador vem primeiro (leitura com trava não fixa o snapshot), então só depois
            # da espera a consulta abaixo enxerga o evento que outra requisição acabou de abrir.
            # Não troque a ordem nem coloque uma leitura comum antes da trava.
            colaborador = obter_colaborador(cur, colaborador_id, travar=True)
            if colaborador is None:
                raise NaoEncontrado("Colaborador não encontrado.")
            cur.execute(
                "SELECT id FROM acesso_eventos WHERE colaborador_id = %s AND status = 'aberto' "
                "LIMIT 1",
                (colaborador_id,),
            )
            if cur.fetchone() is not None:
                raise Conflito("Já existe um evento de acesso aberto para este colaborador.")

            cargo_anterior, cargo_novo = self._cargos_do_evento(
                tipo, colaborador, cargo_anterior, cargo_novo
            )
            pacote_novo = pacote_anterior = None
            if tipo == "admissao":
                pacote_novo = self.acessos.sistemas_do_pacote(cargo_novo, cur=cur)
                if not pacote_novo:
                    raise Invalido(f'O cargo "{cargo_novo}" não tem pacote de acessos definido.')
            elif tipo == "mudanca_funcao":
                pacote_anterior = self.acessos.sistemas_do_pacote(cargo_anterior, cur=cur)
                pacote_novo = self.acessos.sistemas_do_pacote(cargo_novo, cur=cur)

            cur.execute(
                "INSERT INTO acesso_eventos (colaborador_id, tipo, cargo_anterior, cargo_novo, "
                "status, observacao, aberto_por, aberto_em) VALUES (%s, %s, %s, %s, 'aberto', %s, %s, %s)",
                (colaborador_id, tipo, cargo_anterior, cargo_novo, observacao, ctx.username, agora()),
            )
            evento_id = cur.lastrowid

            ativos = self.acessos.listar_acessos(
                colaborador_id=colaborador_id, status="ativo", cur=cur
            )
            ativo_por_sistema: dict[int, dict] = {}
            for acesso in ativos:
                ativo_por_sistema.setdefault(acesso["sistema_id"], acesso)

            quantidade = 0
            if tipo == "desligamento":
                for acesso in ativos:
                    self._inserir_tarefa(
                        cur, evento_id, acesso["sistema_id"], "desativar",
                        acesso_id=acesso["id"], login=acesso["login"],
                    )
                    quantidade += 1
            else:
                if tipo == "admissao":
                    criar, desativar = pacote_novo, []
                else:
                    criar, desativar = diferenca_pacotes(pacote_anterior, pacote_novo)
                for sistema in criar:
                    existente = ativo_por_sistema.get(sistema["id"])
                    if existente is not None:
                        self._inserir_tarefa(
                            cur, evento_id, sistema["id"], "criar",
                            acesso_id=existente["id"], login=existente["login"],
                            resolvida_por=ctx.username,
                            observacao=f"Acesso já ativo: {existente['login']}",
                        )
                    else:
                        self._inserir_tarefa(cur, evento_id, sistema["id"], "criar")
                    quantidade += 1
                for sistema in desativar:
                    for acesso in ativos:
                        if acesso["sistema_id"] == sistema["id"]:
                            self._inserir_tarefa(
                                cur, evento_id, sistema["id"], "desativar",
                                acesso_id=acesso["id"], login=acesso["login"],
                            )
                            quantidade += 1

            audit.record(
                ctx, "acesso_evento", evento_id, "criar",
                depois={
                    "tipo": tipo,
                    "colaborador_id": colaborador_id,
                    "cargo_anterior": cargo_anterior,
                    "cargo_novo": cargo_novo,
                    "tarefas": quantidade,
                },
                cur=cur,
            )
            self._tentar_concluir(ctx, evento_id, cur)
            return self.obter(evento_id, cur)

    # ── Conclusão ───────────────────────────────────────────────────────────

    def _tentar_concluir(self, ctx: AuditContext, evento_id: int, cur) -> bool:
        """Conclui o evento aberto quando não resta tarefa pendente nem (no desligamento)
        equipamento com o colaborador."""
        evento = self.obter(evento_id, cur)
        if evento is None or evento["status"] != "aberto" or evento["pendencias"]:
            return False
        cur.execute(
            "UPDATE acesso_eventos SET status = 'concluido', concluido_por = %s, concluido_em = %s "
            "WHERE id = %s AND status = 'aberto'",
            (ctx.username, agora(), evento_id),
        )
        audit.record(
            ctx, "acesso_evento", evento_id, "concluir",
            antes={"status": "aberto"}, depois={"status": "concluido"}, cur=cur,
        )
        return True

    # ── Resolução de tarefas, conclusão e cancelamento ──────────────────────

    @staticmethod
    def _travar_evento(cur, evento_id: int) -> dict:
        cur.execute(
            "SELECT id, colaborador_id, tipo, status FROM acesso_eventos WHERE id = %s FOR UPDATE",
            (evento_id,),
        )
        evento = cur.fetchone()
        if evento is None:
            raise NaoEncontrado("Evento não encontrado.")
        return evento

    def resolver_tarefa(
        self,
        ctx: AuditContext,
        tarefa_id: int,
        *,
        status: str,
        login: Optional[str],
        item_id: Optional[int],
        observacao: Optional[str],
    ) -> dict:
        with get_cursor() as cur:
            # Concorrência: a PRIMEIRA instrução da transação precisa ser uma leitura com
            # trava. Uma leitura comum fixaria o snapshot do REPEATABLE READ antes de esperar
            # a trava do evento, e as leituras seguintes (acesso, pendências, conclusão)
            # enxergariam o estado de antes do outro técnico. Com a trava primeiro, o snapshot
            # só nasce depois que o outro commit já está visível; quem escreve nas tarefas de
            # um evento sempre passa pela trava do evento, então esse estado não muda mais
            # durante a transação. (Não trave as OUTRAS tarefas do evento: quem espera a trava
            # do evento já segura a da própria tarefa, e travá-las geraria deadlock.)
            cur.execute(
                "SELECT id, evento_id FROM acesso_tarefas WHERE id = %s FOR UPDATE", (tarefa_id,)
            )
            achada = cur.fetchone()
            if achada is None:
                raise NaoEncontrado("Tarefa não encontrada.")
            evento = self._travar_evento(cur, achada["evento_id"])
            # Relê a tarefa com trava: leitura com trava devolve o último estado confirmado.
            cur.execute(
                "SELECT id, sistema_id, acesso_id, acao, status, login "
                "FROM acesso_tarefas WHERE id = %s FOR UPDATE",
                (tarefa_id,),
            )
            tarefa = cur.fetchone()
            if tarefa is None:
                raise NaoEncontrado("Tarefa não encontrada.")
            sistema = self.acessos.obter_sistema(tarefa["sistema_id"], cur)
            tarefa["sistema_nome"] = sistema["nome"] if sistema else ""
            if evento["status"] != "aberto":
                raise Conflito("O evento não está aberto.")
            if tarefa["status"] != "pendente":
                raise Conflito("Esta tarefa já foi resolvida.")

            acesso_id = tarefa["acesso_id"]
            login_final = tarefa["login"]
            nota = observacao
            if status == "manter":
                if tarefa["acao"] != "desativar" or evento["tipo"] != "mudanca_funcao":
                    raise Invalido(
                        "Manter só vale para tarefas de desativar em mudança de função."
                    )
                if not observacao:
                    raise Invalido("Informe o motivo para manter o acesso.")
                novo_status = "nao_se_aplica"
                nota = f"Acesso mantido: {observacao}"
            elif status == "nao_se_aplica":
                if not observacao:
                    raise Invalido("Informe a observação para marcar como não se aplica.")
                novo_status = "nao_se_aplica"
            else:
                novo_status = "feito"
                if tarefa["acao"] == "criar":
                    if not login:
                        raise Invalido("Informe o login criado no sistema.")
                    existente = self._acesso_ativo_com_login(
                        cur, evento["colaborador_id"], tarefa["sistema_id"], login
                    )
                    if existente is not None:
                        acesso_id, login_final = existente["id"], existente["login"]
                    else:
                        acesso = self.acessos.criar_acesso(
                            ctx,
                            colaborador_id=evento["colaborador_id"],
                            sistema_id=tarefa["sistema_id"],
                            login=login,
                            item_id=item_id,
                            cur=cur,
                        )
                        acesso_id, login_final = acesso["id"], acesso["login"]
                elif acesso_id is not None:
                    # Leitura com trava: outro técnico pode ter desativado o acesso agora mesmo.
                    cur.execute(
                        "SELECT status FROM colaborador_acessos WHERE id = %s FOR UPDATE",
                        (acesso_id,),
                    )
                    atual = cur.fetchone()
                    if atual is not None and atual["status"] == "ativo":
                        self.acessos.desativar_acesso(ctx, acesso_id, cur=cur)

            cur.execute(
                "UPDATE acesso_tarefas SET status = %s, login = %s, acesso_id = %s, "
                "feito_por = %s, feito_em = %s, observacao = COALESCE(%s, observacao) WHERE id = %s",
                (novo_status, login_final, acesso_id, ctx.username, agora(), nota, tarefa_id),
            )
            audit.record(
                ctx, "acesso_evento", evento["id"], "resolver_tarefa",
                depois={
                    "tarefa_id": tarefa_id,
                    "sistema": tarefa["sistema_nome"],
                    "acao": tarefa["acao"],
                    "status": novo_status,
                    "login": login_final,
                    "observacao": nota,
                },
                cur=cur,
            )
            self._tentar_concluir(ctx, evento["id"], cur)
            return self.obter(evento["id"], cur)

    @staticmethod
    def _acesso_ativo_com_login(cur, colaborador_id: int, sistema_id: int, login: str):
        cur.execute(
            "SELECT id, login FROM colaborador_acessos WHERE colaborador_id = %s AND "
            "sistema_id = %s AND status = 'ativo' AND LOWER(login) = LOWER(%s) LIMIT 1 FOR UPDATE",
            (colaborador_id, sistema_id, login),
        )
        return cur.fetchone()

    def concluir(self, ctx: AuditContext, evento_id: int) -> dict:
        with get_cursor() as cur:
            evento = self._travar_evento(cur, evento_id)
            if evento["status"] != "aberto":
                raise Conflito("O evento não está aberto.")
            pendencias = self.obter(evento_id, cur)["pendencias"]
            if pendencias:
                raise Conflito("Não é possível concluir: " + "; ".join(pendencias) + ".")
            self._tentar_concluir(ctx, evento_id, cur)
            return self.obter(evento_id, cur)

    def cancelar(self, ctx: AuditContext, evento_id: int, *, motivo: str) -> dict:
        with get_cursor() as cur:
            evento = self._travar_evento(cur, evento_id)
            if evento["status"] != "aberto":
                raise Conflito("O evento não está aberto.")
            cur.execute(
                "UPDATE acesso_eventos SET status = 'cancelado', cancelado_por = %s, "
                "cancelado_em = %s, motivo_cancelamento = %s WHERE id = %s",
                (ctx.username, agora(), motivo, evento_id),
            )
            audit.record(
                ctx, "acesso_evento", evento_id, "cancelar",
                depois={"motivo_cancelamento": motivo},
                cur=cur,
            )
            return self.obter(evento_id, cur)
