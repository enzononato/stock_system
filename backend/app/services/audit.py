"""Trilha de auditoria: registro das alterações feitas no sistema.

Toda escrita nova chama `record(...)`. Com `cur`, o registro entra na MESMA
transação da alteração (e um erro aqui desfaz a alteração, de propósito: não
existe mudança sem rastro). Sem `cur`, o registro é "melhor esforço": abre o
próprio cursor e, se falhar, só loga -- nunca derruba a operação principal.

CPF nunca é gravado por inteiro: a chave `cpf` é mascarada como `***.456.789-**`.
"""
import json
import logging
import re
from dataclasses import dataclass
from datetime import date, datetime, timezone
from decimal import Decimal
from typing import Any, Optional

from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)


@dataclass
class AuditContext:
    usuario_id: Optional[int]
    username: Optional[str]
    papel: Optional[str]
    ip: Optional[str]

    @classmethod
    def sistema(cls) -> "AuditContext":
        """Contexto para ações automáticas (jobs do agendador, sincronizações)."""
        return cls(usuario_id=None, username="sistema", papel=None, ip=None)


def _mascarar_cpf(valor: Any) -> Any:
    if valor is None:
        return None
    digitos = re.sub(r"\D", "", str(valor))
    if len(digitos) != 11:
        return "***"  # formato inesperado: melhor não vazar nada
    return f"***.{digitos[3:6]}.{digitos[6:9]}-**"


def _mascarar_folhas_cpf(valor: Any) -> Any:
    """Sob uma chave `cpf`: mascara toda folha escalar, mesmo dentro de
    dict/lista/tupla (`{"cpfs": ["12345678909"]}`)."""
    if isinstance(valor, dict):
        return {chave: _mascarar_folhas_cpf(v) for chave, v in valor.items()}
    if isinstance(valor, (list, tuple)):
        return [_mascarar_folhas_cpf(v) for v in valor]
    return _mascarar_cpf(valor)


def _mascarar_valor(valor: Any) -> Any:
    if isinstance(valor, dict):
        return {
            chave: (
                _mascarar_folhas_cpf(v)
                if "cpf" in str(chave).lower()
                else _mascarar_valor(v)
            )
            for chave, v in valor.items()
        }
    if isinstance(valor, (list, tuple)):
        return [_mascarar_valor(v) for v in valor]
    return valor


def mascarar(valor: Optional[dict]) -> Optional[dict]:
    """Cópia de `valor` com mascarado todo valor cuja chave contém `cpf`
    (`cpf`, `cpf_colaborador`...), recursivo em dicts/listas."""
    if valor is None:
        return None
    return _mascarar_valor(valor)


def _serializar(valor: Any) -> Any:
    """Datas em ISO e decimais como texto; demais valores ficam como estão."""
    if isinstance(valor, (datetime, date)):
        return valor.isoformat()
    if isinstance(valor, Decimal):
        return str(valor)
    return valor


def _default_json(valor: Any) -> Any:
    serializado = _serializar(valor)
    return serializado if serializado is not valor else str(valor)


_AUSENTE = object()


def diff(
    antes: Optional[dict], depois: Optional[dict]
) -> tuple[Optional[dict], Optional[dict]]:
    """Reduz `antes`/`depois` às chaves cujo valor mudou. Se um dos lados for
    None, devolve os lados como estão (criação/exclusão). Datas e decimais saem
    como texto."""
    if antes is None or depois is None:
        return (
            None if antes is None else {k: _serializar(v) for k, v in antes.items()},
            None if depois is None else {k: _serializar(v) for k, v in depois.items()},
        )
    novo_antes: dict = {}
    novo_depois: dict = {}
    for chave in sorted(antes.keys() | depois.keys(), key=str):
        a = _serializar(antes.get(chave, _AUSENTE))
        d = _serializar(depois.get(chave, _AUSENTE))
        if a != d:
            # Chave ausente de um lado aparece como None nesse lado.
            novo_antes[chave] = None if a is _AUSENTE else a
            novo_depois[chave] = None if d is _AUSENTE else d
    return novo_antes, novo_depois


def _json(valor: Optional[dict]) -> Optional[str]:
    if valor is None:
        return None
    return json.dumps(mascarar(valor), ensure_ascii=False, default=_default_json)


_INSERT = (
    "INSERT INTO audit_log (ocorrido_em, usuario_id, username, papel, ip, "
    "entidade, entidade_id, acao, antes, depois) "
    "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
)


def record(
    ctx: AuditContext,
    entidade: str,
    entidade_id: "int | str | None",
    acao: str,
    antes: Optional[dict] = None,
    depois: Optional[dict] = None,
    cur=None,
) -> None:
    """Grava uma linha em `audit_log`.

    Com `cur`: usa o cursor recebido (mesma transação) e propaga qualquer erro.
    Sem `cur`: abre `get_cursor()` próprio; em erro faz `logger.exception` e
    NÃO propaga.
    """
    def _params() -> tuple:
        return (
            datetime.now(timezone.utc).replace(tzinfo=None),  # UTC, sem fuso (DATETIME)
            ctx.usuario_id,
            ctx.username,
            ctx.papel,
            ctx.ip,
            entidade,
            None if entidade_id is None else str(entidade_id),
            acao,
            _json(antes),
            _json(depois),
        )

    if cur is not None:
        cur.execute(_INSERT, _params())
        return
    try:
        with get_cursor(dict_cursor=False) as proprio:
            proprio.execute(_INSERT, _params())
    except Exception:
        logger.exception(
            "Falha ao registrar auditoria (entidade=%s, id=%s, acao=%s).",
            entidade, entidade_id, acao,
        )
