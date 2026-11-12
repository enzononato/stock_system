"""
Fonte de colaboradores: leitura SOMENTE LEITURA da base de RH (Postgres).

Este é o único módulo do sistema que fala com o Postgres de RH (tabela
`"colaboradores-revalle"` do database `agente_revalle`). A base é de outro sistema:

- a sessão é aberta com `default_transaction_read_only=on`, então qualquer gravação é
  recusada pelo servidor mesmo que o usuário da conexão tenha permissão de escrita;
- só são lidas as colunas mapeadas em `CAMPOS` (a data de nascimento nunca é selecionada
  nem exposta);
- os nomes reais das colunas não são conhecidos de antemão: são descobertos por sinônimos
  (`SINONIMOS`), com `COLABORADORES_COLUMNS` ("campo=coluna,campo=coluna") como ajuste manual;
- a string de conexão (com senha) nunca vai para log nem para mensagem de erro.

As configurações são lidas a cada chamada (não no import), para que testes e
reconfigurações valham sem reiniciar o processo.
"""
import logging
import re
import unicodedata
from dataclasses import dataclass
from typing import Literal

import psycopg
from psycopg import sql

from app.core import config as _config

logger = logging.getLogger(__name__)

MSG_INDISPONIVEL = "Base de colaboradores indisponível no momento."


class ColaboradoresSourceError(Exception):
    """Configuração/estrutura da base de RH inválida (tabela ou coluna não encontrada)."""


class ColaboradoresIndisponivel(ColaboradoresSourceError):
    """Base de RH sem URL configurada ou inacessível."""

    def __init__(self, mensagem: str = MSG_INDISPONIVEL):
        super().__init__(mensagem)


CAMPOS = (
    "cpf", "nome", "cargo", "setor", "unidade", "matricula",
    "email", "telefone", "situacao", "gestor", "demissao",
)

SINONIMOS: dict[str, tuple[str, ...]] = {
    "cpf": ("cpf", "cpf_completo", "nr_cpf", "num_cpf", "documento"),
    "nome": ("nome", "nome_completo", "colaborador", "nome_colaborador", "funcionario"),
    "cargo": ("cargo", "cargo_nome", "funcao", "nome_cargo"),
    "setor": ("setor", "setor_nome", "departamento", "area"),
    "unidade": ("unidade", "unidade_nome", "filial", "empresa", "revenda", "loja"),
    "matricula": ("matricula", "matricula_rh", "registro", "codigo", "id_colaborador"),
    "email": ("email", "e_mail", "email_corp", "email_corporativo", "email_pessoal"),
    "telefone": ("telefone", "celular", "fone", "whatsapp", "telefone_celular"),
    "situacao": ("situacao", "status", "status_rh", "situacao_rh", "ativo"),
    "gestor": ("gestor", "superior", "gestor_imediato", "lider", "nome_gestor", "superior_imediato"),
    "demissao": ("demissao", "data_demissao", "dt_demissao", "desligamento", "data_desligamento"),
}

_OBRIGATORIOS = ("cpf", "nome")


@dataclass
class ColaboradorRH:
    cpf: str
    nome: str
    cargo: str | None
    setor: str | None
    unidade: str | None
    matricula: str | None
    email: str | None
    telefone: str | None
    situacao: str
    gestor_nome: str | None


# ────────────────────────────────────────────────────────────────────────────
# Normalização
# ────────────────────────────────────────────────────────────────────────────

def _sem_acento(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


def normalizar_nome_coluna(nome: str) -> str:
    """Minúsculas, sem acento, e todo trecho não alfanumérico vira `_` ("E-mail" -> "e_mail")."""
    return re.sub(r"[^a-z0-9]+", "_", _sem_acento(nome).lower()).strip("_")


def _texto_ou_none(valor: object) -> str | None:
    if valor is None:
        return None
    texto = str(valor).strip()
    return texto or None


def _so_digitos(valor: object) -> str:
    return re.sub(r"\D", "", "" if valor is None else str(valor))


def _parse_override(override: str) -> dict[str, str]:
    resultado: dict[str, str] = {}
    for parte in (override or "").split(","):
        if "=" not in parte:
            continue
        campo, coluna = parte.split("=", 1)
        campo, coluna = campo.strip().lower(), coluna.strip()
        if campo in CAMPOS and coluna:
            resultado[campo] = coluna
    return resultado


def descobrir_colunas(colunas_reais: list[str], override: str = "") -> dict[str, str]:
    """Mapeia cada campo de `CAMPOS` para o nome REAL da coluna na tabela.

    Para cada campo vale a primeira coluna real cujo nome normalizado está entre os sinônimos
    (na ordem dos sinônimos). `override` ("campo=coluna,campo=coluna") tem prioridade.
    Campos sem coluna ficam de fora; só `cpf` e `nome` são obrigatórios.
    """
    por_normalizado: dict[str, str] = {}
    for real in colunas_reais:
        por_normalizado.setdefault(normalizar_nome_coluna(real), real)

    mapa: dict[str, str] = {}
    for campo, coluna in _parse_override(override).items():
        if coluna in colunas_reais:
            mapa[campo] = coluna
        elif normalizar_nome_coluna(coluna) in por_normalizado:
            mapa[campo] = por_normalizado[normalizar_nome_coluna(coluna)]
        else:
            raise ColaboradoresSourceError(
                f"Base de colaboradores: a coluna '{coluna}' indicada para {campo} não existe na tabela"
            )

    for campo in CAMPOS:
        if campo in mapa:
            continue
        for sinonimo in SINONIMOS[campo]:
            if sinonimo in por_normalizado:
                mapa[campo] = por_normalizado[sinonimo]
                break

    for campo in _OBRIGATORIOS:
        if campo not in mapa:
            raise ColaboradoresSourceError(f"Base de colaboradores: coluna não encontrada para {campo}")
    return mapa


def normalizar_situacao(valor: object, demissao: object = None) -> Literal["ativo", "afastado", "desligado"]:
    """Situação de RH -> ativo | afastado | desligado. Data de demissão preenchida = desligado."""
    if _texto_ou_none(demissao) is not None:
        return "desligado"
    texto = _sem_acento(str(valor)).strip().lower() if valor is not None else ""
    if texto in ("false", "0", "n", "nao"):
        return "desligado"
    if any(chave in texto for chave in ("deslig", "demit", "inativ", "rescis")):
        return "desligado"
    if any(chave in texto for chave in ("afast", "licen")):
        return "afastado"
    return "ativo"


# ────────────────────────────────────────────────────────────────────────────
# Acesso à base (somente leitura)
# ────────────────────────────────────────────────────────────────────────────

def conectar() -> psycopg.Connection:
    """Conexão com a base de RH, em sessão somente leitura. Quem chama fecha (`with`)."""
    url = (_config.settings.COLABORADORES_DATABASE_URL or "").strip()
    if not url:
        raise ColaboradoresIndisponivel()
    try:
        return psycopg.connect(
            url,
            connect_timeout=10,
            options="-c default_transaction_read_only=on",
        )
    except psycopg.Error as e:
        # Só o tipo do erro: a mensagem do driver pode trazer host/usuário da conexão.
        logger.error("Falha ao conectar na base de colaboradores (%s).", type(e).__name__)
        raise ColaboradoresIndisponivel() from None


def _descobrir_na_base(conn: psycopg.Connection) -> dict[str, str]:
    schema = _config.settings.COLABORADORES_SCHEMA or "public"
    tabela = _config.settings.COLABORADORES_TABLE
    linhas = conn.execute(
        "SELECT column_name FROM information_schema.columns "
        "WHERE table_schema = %s AND table_name = %s ORDER BY ordinal_position",
        (schema, tabela),
    ).fetchall()
    if not linhas:
        raise ColaboradoresSourceError(
            f"Base de colaboradores: tabela {schema}.{tabela} não encontrada ou sem permissão"
        )
    return descobrir_colunas([str(linha[0]) for linha in linhas], _config.settings.COLABORADORES_COLUMNS)


def _consultar(conn: psycopg.Connection, mapa: dict[str, str], cpf_digitos: str | None) -> list[ColaboradorRH]:
    campos = [c for c in CAMPOS if c in mapa]
    tabela = sql.Identifier(_config.settings.COLABORADORES_SCHEMA or "public", _config.settings.COLABORADORES_TABLE)
    consulta = sql.SQL("SELECT {cols} FROM {tabela}").format(
        cols=sql.SQL(", ").join(sql.Identifier(mapa[c]) for c in campos),
        tabela=tabela,
    )
    params: tuple = ()
    if cpf_digitos is not None:
        digitos_coluna = sql.SQL("regexp_replace(CAST({col} AS text), '\\D', '', 'g')").format(
            col=sql.Identifier(mapa["cpf"])
        )
        consulta += sql.SQL(
            " WHERE length({d}) BETWEEN 1 AND 11 AND lpad({d}, 11, '0') = %s"
        ).format(d=digitos_coluna)
        params = (cpf_digitos,)

    resultado: list[ColaboradorRH] = []
    for linha in conn.execute(consulta, params).fetchall():
        bruto = dict(zip(campos, linha))
        digitos = _so_digitos(bruto.get("cpf"))
        nome = _texto_ou_none(bruto.get("nome"))
        if not digitos or len(digitos) > 11 or not nome:
            continue
        resultado.append(
            ColaboradorRH(
                cpf=digitos.zfill(11),
                nome=nome,
                cargo=_texto_ou_none(bruto.get("cargo")),
                setor=_texto_ou_none(bruto.get("setor")),
                unidade=_texto_ou_none(bruto.get("unidade")),
                matricula=_texto_ou_none(bruto.get("matricula")),
                email=_texto_ou_none(bruto.get("email")),
                telefone=_texto_ou_none(bruto.get("telefone")),
                situacao=normalizar_situacao(bruto.get("situacao"), bruto.get("demissao")),
                gestor_nome=_texto_ou_none(bruto.get("gestor")),
            )
        )
    return resultado


def _executar(cpf_digitos: str | None) -> list[ColaboradorRH]:
    try:
        with conectar() as conn:
            return _consultar(conn, _descobrir_na_base(conn), cpf_digitos)
    except ColaboradoresSourceError:
        raise
    except psycopg.Error as e:
        logger.error("Falha ao consultar a base de colaboradores (%s).", type(e).__name__)
        raise ColaboradoresIndisponivel() from None


def ler_todos() -> list[ColaboradorRH]:
    """Todos os colaboradores da base de RH (CPF só dígitos com 11 posições, nome não vazio)."""
    return _executar(None)


def buscar_por_cpf(cpf: str) -> ColaboradorRH | None:
    """Um colaborador pelo CPF (com ou sem máscara), ou None se não existir."""
    digitos = _so_digitos(cpf)
    if not digitos or len(digitos) > 11:
        # Ainda assim exige a base configurada, para não mascarar indisponibilidade com "não achou".
        if not (_config.settings.COLABORADORES_DATABASE_URL or "").strip():
            raise ColaboradoresIndisponivel()
        return None
    achados = _executar(digitos.zfill(11))
    return achados[0] if achados else None
