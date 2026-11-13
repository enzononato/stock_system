"""
Camada de acesso ao cadastro de colaboradores.

As tabelas `colaboradores` e `colaboradores_sync` são criadas pela migração
0001 (e `unidades.alias_rh` pela 0002); este módulo não cria schema.

O cadastro é um espelho da base de RH (somente leitura, ver
`app.services.colaboradores_source`) com complemento manual:

- `sincronizar` grava os campos que vêm do RH; e-mail e telefone só são
  sobrescritos enquanto ninguém os tenha editado (`email_manual`/`telefone_manual`);
- colaboradores `origem='manual'` não vêm do RH e podem ser editados por inteiro;
- datas são gravadas em UTC e devolvidas com fuso (`+00:00`).
"""
import logging
import re
import unicodedata
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from typing import Optional

import pymysql

from app.db.database_mysql import get_cursor
from app.schemas.validators import normalizar_telefone, validar_email
from app.services import audit
from app.services.audit import AuditContext
from app.services.colaboradores_source import (
    ColaboradorRH,
    ColaboradoresSourceError,
    cpf_valido,
    ler_todos,
)

logger = logging.getLogger(__name__)

MSG_CAMPO_DO_RH = "Campo vem da base de RH e não pode ser editado aqui."
MSG_SUPERIOR_DE_SI = "O colaborador não pode ser superior de si mesmo."
MSG_CPF_DUPLICADO = "Já existe um colaborador com este CPF."

# Campos que, em colaborador vindo do RH, só a sincronização altera.
_CAMPOS_DO_RH = ("nome", "cargo", "setor", "unidade_id", "situacao")
_CAMPOS_DATA = ("sincronizado_em", "criado_em", "atualizado_em")
_LIMITES = {
    "nome": 200, "matricula": 50, "cargo": 150, "setor": 150,
    "unidade_rh": 150, "gestor_nome_rh": 200, "email": 200,
}

_SELECT_COM_NOMES = (
    "SELECT c.*, u.nome AS unidade_nome, s.nome AS superior_nome "
    "FROM colaboradores c "
    "LEFT JOIN unidades u ON u.id = c.unidade_id "
    "LEFT JOIN colaboradores s ON s.id = c.superior_id"
)


@dataclass
class ResumoSync:
    """Resultado de uma sincronização. `total_base` é o número de linhas lidas da
    base de RH; `ignorados` (CPF inválido) e `duplicados` (CPF repetido, vale o
    primeiro por ordem de nome) são subconjuntos dele que não entram no cadastro."""

    total_base: int
    criados: int
    atualizados: int
    desligados: int
    ignorados: int = 0
    duplicados: int = 0


def _agora() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def _chave(texto: Optional[str]) -> str:
    """Comparação sem acento, caixa nem espaços repetidos."""
    sem_acento = "".join(
        c for c in unicodedata.normalize("NFD", texto or "") if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", sem_acento).strip().casefold()


def _cortar(valor: Optional[str], campo: str) -> Optional[str]:
    return None if valor is None else valor[: _LIMITES[campo]]


def _com_fuso(linha: Optional[dict]) -> Optional[dict]:
    if linha is None:
        return None
    for campo in _CAMPOS_DATA:
        valor = linha.get(campo)
        if isinstance(valor, datetime) and valor.tzinfo is None:
            linha[campo] = valor.replace(tzinfo=timezone.utc)
    return linha


def _escapar_like(texto: str) -> str:
    return texto.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def _email_do_rh(valor: Optional[str]) -> Optional[str]:
    try:
        return validar_email(valor)
    except ValueError:
        return None


def _telefone_do_rh(valor: Optional[str]) -> Optional[str]:
    try:
        return normalizar_telefone(valor)
    except ValueError:
        return None


class _ResolvedorUnidades:
    """Liga o nome de unidade do RH a `unidades.id`: primeiro por `alias_rh` igual
    (sem acento/caixa), na falta por `unidades.nome` contendo o nome do RH (só se
    houver exatamente uma)."""

    def __init__(self, unidades: list[dict]):
        self._unidades = sorted(unidades, key=lambda u: u["id"])

    def resolver(self, unidade_rh: Optional[str]) -> Optional[int]:
        alvo = _chave(unidade_rh)
        if not alvo:
            return None
        for u in self._unidades:
            if _chave(u.get("alias_rh")) == alvo:
                return u["id"]
        contem = [u["id"] for u in self._unidades if alvo in _chave(u["nome"])]
        return contem[0] if len(contem) == 1 else None


def _depurar(registros: list[ColaboradorRH]) -> tuple[list[ColaboradorRH], int, int]:
    """Descarta CPF inválido e CPF repetido (vale o primeiro por ordem de nome).
    Devolve (válidos, ignorados, duplicados)."""
    ordenados = sorted(registros, key=lambda r: (_chave(r.nome), r.cpf))
    vistos: set[str] = set()
    validos: list[ColaboradorRH] = []
    ignorados = duplicados = 0
    for r in ordenados:
        if not cpf_valido(r.cpf):
            ignorados += 1
        elif r.cpf in vistos:
            duplicados += 1
        else:
            vistos.add(r.cpf)
            validos.append(r)
    return validos, ignorados, duplicados


class ColaboradorDBManager:
    # ── Sincronização ────────────────────────────────────────────────────────

    def sincronizar(self, registros: list[ColaboradorRH]) -> ResumoSync:
        """Upsert por CPF dos registros da base de RH, tudo numa transação.

        Colaboradores `origem='rh'` ausentes da base viram `desligado`. Uma
        colaborador criado manualmente cujo CPF aparece na base passa a `origem='rh'`.
        """
        validos, ignorados, duplicados = _depurar(registros)
        agora = _agora()
        criados = atualizados = desligados = 0

        with get_cursor() as cur:
            cur.execute("SELECT id, nome, alias_rh FROM unidades")
            unidades = _ResolvedorUnidades(cur.fetchall())
            cur.execute("SELECT * FROM colaboradores")
            existentes = {linha["cpf"]: linha for linha in cur.fetchall()}

            for r in validos:
                novo = {
                    "nome": _cortar(r.nome, "nome"),
                    "matricula": _cortar(r.matricula, "matricula"),
                    "cargo": _cortar(r.cargo, "cargo"),
                    "setor": _cortar(r.setor, "setor"),
                    "unidade_rh": _cortar(r.unidade, "unidade_rh"),
                    "unidade_id": unidades.resolver(r.unidade),
                    "situacao": r.situacao,
                    "gestor_nome_rh": _cortar(r.gestor_nome, "gestor_nome_rh"),
                    "origem": "rh",
                }
                email = _cortar(_email_do_rh(r.email), "email")
                telefone = _telefone_do_rh(r.telefone)
                atual = existentes.get(r.cpf)

                if atual is None:
                    novo.update(email=email, telefone=telefone)
                    colunas = ["cpf", *novo, "sincronizado_em", "criado_em", "atualizado_em"]
                    cur.execute(
                        f"INSERT INTO colaboradores ({', '.join(colunas)}) "
                        f"VALUES ({', '.join(['%s'] * len(colunas))})",
                        [r.cpf, *novo.values(), agora, agora, agora],
                    )
                    criados += 1
                    continue

                if not atual["email_manual"]:
                    novo["email"] = email
                if not atual["telefone_manual"]:
                    novo["telefone"] = telefone
                mudou = {k: v for k, v in novo.items() if atual[k] != v}
                if mudou:
                    sets = ", ".join(f"{k}=%s" for k in mudou)
                    cur.execute(
                        f"UPDATE colaboradores SET {sets}, sincronizado_em=%s, atualizado_em=%s WHERE id=%s",
                        [*mudou.values(), agora, agora, atual["id"]],
                    )
                    atualizados += 1
                else:
                    cur.execute(
                        "UPDATE colaboradores SET sincronizado_em=%s WHERE id=%s", (agora, atual["id"])
                    )

            cpfs_na_base = {r.cpf for r in validos}
            ausentes = [
                linha["id"]
                for cpf, linha in existentes.items()
                if linha["origem"] == "rh" and cpf not in cpfs_na_base and linha["situacao"] != "desligado"
            ]
            if ausentes:
                cur.execute(
                    "UPDATE colaboradores SET situacao='desligado', atualizado_em=%s "
                    f"WHERE id IN ({', '.join(['%s'] * len(ausentes))})",
                    [agora, *ausentes],
                )
                desligados = len(ausentes)

            self._preencher_superiores(cur, agora)

        return ResumoSync(
            total_base=len(registros),
            criados=criados,
            atualizados=atualizados,
            desligados=desligados,
            ignorados=ignorados,
            duplicados=duplicados,
        )

    @staticmethod
    def _preencher_superiores(cur, agora: datetime) -> None:
        """Segunda passada: `superior_id` vazio + `gestor_nome_rh` igual ao nome de
        exatamente um colaborador -> preenche."""
        cur.execute("SELECT id, nome, superior_id, gestor_nome_rh FROM colaboradores")
        todos = cur.fetchall()
        por_nome: dict[str, list[int]] = {}
        for linha in todos:
            por_nome.setdefault(_chave(linha["nome"]), []).append(linha["id"])
        for linha in todos:
            if linha["superior_id"] is not None or not linha["gestor_nome_rh"]:
                continue
            candidatos = por_nome.get(_chave(linha["gestor_nome_rh"]), [])
            if len(candidatos) == 1 and candidatos[0] != linha["id"]:
                cur.execute(
                    "UPDATE colaboradores SET superior_id=%s, atualizado_em=%s WHERE id=%s",
                    (candidatos[0], agora, linha["id"]),
                )

    def registrar_sync_inicio(self) -> int:
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "INSERT INTO colaboradores_sync (iniciado_em, status) VALUES (%s, 'em_andamento')",
                (_agora(),),
            )
            return cur.lastrowid

    def registrar_sync_fim(self, sync_id: int, resumo: Optional[ResumoSync], erro: Optional[str]) -> None:
        if resumo is None:
            resumo = ResumoSync(0, 0, 0, 0)
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "UPDATE colaboradores_sync SET concluido_em=%s, status=%s, total_base=%s, criados=%s, "
                "atualizados=%s, desligados=%s, ignorados=%s, duplicados=%s, erro=%s WHERE id=%s",
                (
                    _agora(),
                    "erro" if erro else "ok",
                    resumo.total_base, resumo.criados, resumo.atualizados, resumo.desligados,
                    resumo.ignorados, resumo.duplicados,
                    erro[:500] if erro else None,
                    sync_id,
                ),
            )

    def ultima_sync(self) -> Optional[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT * FROM colaboradores_sync ORDER BY id DESC LIMIT 1")
            linha = cur.fetchone()
        if linha is None:
            return None
        for campo in ("iniciado_em", "concluido_em"):
            if linha.get(campo) is not None:
                linha[campo] = linha[campo].replace(tzinfo=timezone.utc)
        return linha

    # ── Leitura ──────────────────────────────────────────────────────────────

    def listar(
        self,
        *,
        busca: Optional[str] = None,
        situacao: Optional[str] = None,
        unidade_id: Optional[int] = None,
        limit: int = 25,
        offset: int = 0,
    ) -> tuple[list[dict], int]:
        """Colaboradores por ordem de nome. `busca` procura no nome, na matrícula e,
        se for só dígitos/pontuação de CPF, no CPF."""
        condicoes: list[str] = []
        params: list = []
        termo = (busca or "").strip()
        if termo:
            padrao = f"%{_escapar_like(termo)}%"
            alternativas = ["c.nome LIKE %s", "c.matricula LIKE %s"]
            params += [padrao, padrao]
            digitos = re.sub(r"\D", "", termo)
            if digitos and re.fullmatch(r"[\d.\-/\s]+", termo):
                alternativas.append("c.cpf LIKE %s")
                params.append(f"%{digitos}%")
            condicoes.append(f"({' OR '.join(alternativas)})")
        if situacao:
            condicoes.append("c.situacao = %s")
            params.append(situacao)
        if unidade_id is not None:
            condicoes.append("c.unidade_id = %s")
            params.append(unidade_id)
        where = f"WHERE {' AND '.join(condicoes)}" if condicoes else ""

        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT COUNT(*) AS total FROM colaboradores c {where}", params)
            total = cur.fetchone()["total"]
            cur.execute(
                f"{_SELECT_COM_NOMES} {where} ORDER BY c.nome, c.id LIMIT %s OFFSET %s",
                [*params, limit, offset],
            )
            linhas = [_com_fuso(linha) for linha in cur.fetchall()]
        return linhas, total

    def obter(self, id: int) -> Optional[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(f"{_SELECT_COM_NOMES} WHERE c.id = %s", (id,))
            return _com_fuso(cur.fetchone())

    def obter_por_cpf(self, cpf: str) -> Optional[dict]:
        digitos = re.sub(r"\D", "", cpf or "")
        with get_cursor(commit=False) as cur:
            cur.execute(f"{_SELECT_COM_NOMES} WHERE c.cpf = %s", (digitos,))
            return _com_fuso(cur.fetchone())

    def cargos(self) -> list[str]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT DISTINCT cargo FROM colaboradores "
                "WHERE cargo IS NOT NULL AND cargo <> '' ORDER BY cargo"
            )
            return [linha["cargo"] for linha in cur.fetchall()]

    def resolver_unidade_id(self, unidade_rh: Optional[str]) -> Optional[int]:
        """`unidades.id` para um nome de unidade da base de RH (ver `_ResolvedorUnidades`)."""
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT id, nome, alias_rh FROM unidades")
            return _ResolvedorUnidades(cur.fetchall()).resolver(unidade_rh)

    # ── Escrita ──────────────────────────────────────────────────────────────

    @staticmethod
    def _exigir_existe(cur, tabela: str, id_: Optional[int], mensagem: str) -> None:
        if id_ is None:
            return
        cur.execute(f"SELECT 1 FROM {tabela} WHERE id=%s", (id_,))
        if cur.fetchone() is None:
            raise ValueError(mensagem)

    def criar_manual(self, dados: dict) -> int:
        """Cria um colaborador `origem='manual'`. CPF repetido -> ValueError."""
        agora = _agora()
        campos = {
            "cpf": dados["cpf"],
            "nome": dados["nome"],
            "cargo": dados.get("cargo"),
            "setor": dados.get("setor"),
            "unidade_id": dados.get("unidade_id"),
            "situacao": dados.get("situacao") or "ativo",
            "email": dados.get("email"),
            "telefone": dados.get("telefone"),
            # Preenchido à mão: uma futura sincronização (se o CPF entrar na base) não apaga.
            "email_manual": 1 if dados.get("email") else 0,
            "telefone_manual": 1 if dados.get("telefone") else 0,
            "superior_id": dados.get("superior_id"),
            "origem": "manual",
            "criado_em": agora,
            "atualizado_em": agora,
        }
        try:
            with get_cursor() as cur:
                self._exigir_existe(cur, "unidades", campos["unidade_id"], "Unidade não encontrada.")
                self._exigir_existe(cur, "colaboradores", campos["superior_id"], "Superior não encontrado.")
                cur.execute(
                    f"INSERT INTO colaboradores ({', '.join(campos)}) "
                    f"VALUES ({', '.join(['%s'] * len(campos))})",
                    list(campos.values()),
                )
                return cur.lastrowid
        except pymysql.IntegrityError as e:
            if e.args and e.args[0] == 1062:
                raise ValueError(MSG_CPF_DUPLICADO) from None
            raise

    def atualizar(self, id: int, dados: dict) -> None:
        """Aplica só as chaves de `dados` (valor `None` limpa o campo).

        E-mail/telefone diferentes do gravado passam a `*_manual = 1`. `nome`, `cargo`,
        `setor`, `unidade_id` e `situacao` só valem para `origem='manual'`. Levanta
        `LookupError` se o colaborador não existe e `ValueError` para pedido inválido.
        """
        with get_cursor() as cur:
            cur.execute("SELECT * FROM colaboradores WHERE id=%s FOR UPDATE", (id,))
            atual = cur.fetchone()
            if atual is None:
                raise LookupError("Colaborador não encontrado.")

            if atual["origem"] != "manual" and any(c in dados for c in _CAMPOS_DO_RH):
                raise ValueError(MSG_CAMPO_DO_RH)
            if dados.get("superior_id") is not None and dados["superior_id"] == id:
                raise ValueError(MSG_SUPERIOR_DE_SI)
            if "unidade_id" in dados:
                self._exigir_existe(cur, "unidades", dados["unidade_id"], "Unidade não encontrada.")
            if "superior_id" in dados:
                self._exigir_existe(cur, "colaboradores", dados["superior_id"], "Superior não encontrado.")

            novos = dict(dados)
            for campo in ("email", "telefone"):
                if campo in dados and dados[campo] != atual[campo]:
                    novos[f"{campo}_manual"] = 1
            if not novos:
                return
            sets = ", ".join(f"{k}=%s" for k in novos)
            cur.execute(
                f"UPDATE colaboradores SET {sets}, atualizado_em=%s WHERE id=%s",
                [*novos.values(), _agora(), id],
            )


def executar_sincronizacao(db: ColaboradorDBManager, ctx: AuditContext) -> ResumoSync:
    """Lê a base de RH e sincroniza o cadastro, registrando início, fim e auditoria.

    Falha da base (indisponível ou mal configurada) é gravada no log de sincronização
    e relançada. Uma base que devolve zero linhas é tratada como falha: aplicar isso
    marcaria todos os colaboradores do RH como desligados.
    """
    sync_id = db.registrar_sync_inicio()
    try:
        registros = ler_todos()
        if not registros:
            raise ColaboradoresSourceError(
                "A base de colaboradores não retornou nenhum registro; sincronização cancelada."
            )
        resumo = db.sincronizar(registros)
    except ColaboradoresSourceError as e:
        _registrar_falha(db, sync_id, str(e))
        raise
    except Exception as e:
        # Nunca o texto da exceção do driver: pode trazer host/usuário de conexão.
        _registrar_falha(db, sync_id, f"Erro inesperado ({type(e).__name__}).")
        raise
    db.registrar_sync_fim(sync_id, resumo, None)
    audit.record(ctx, "colaborador", None, "sincronizar", depois=asdict(resumo))
    return resumo


def _registrar_falha(db: ColaboradorDBManager, sync_id: int, erro: str) -> None:
    try:
        db.registrar_sync_fim(sync_id, None, erro)
    except Exception:
        logger.exception("Falha ao registrar o fim da sincronização de colaboradores.")
