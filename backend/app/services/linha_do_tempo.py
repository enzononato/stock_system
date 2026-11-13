"""Linha do tempo do equipamento: um registro de fontes.

Cada fonte é uma função `(item_id) -> list[Evento]` que lê UMA tabela e devolve os
eventos dela. `montar_linha_do_tempo` junta todas e ordena. A F4 implementa as
fontes de histórico, manutenções, transferências e baixa
(`app.services.linha_do_tempo_fontes`); termos, linha telefônica e inventário
entram depois, acrescentando o módulo delas em `MODULOS_DE_FONTES` (a F7 faz isso).
"""
import importlib
import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, Optional

from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)


@dataclass
class Evento:
    ocorrido_em: datetime
    origem: str
    tipo: str
    titulo: str
    descricao: Optional[str] = None
    usuario: Optional[str] = None
    referencia_id: Optional[int] = None


FonteLinhaDoTempo = Callable[[int], list[Evento]]

# Ordem de registro preservada; a chave é o nome da função.
_FONTES: dict[str, FonteLinhaDoTempo] = {}

# A F7 acrescenta aqui o módulo das fontes de termos, linha telefônica e inventário.
MODULOS_DE_FONTES: tuple[str, ...] = ("app.services.linha_do_tempo_fontes",)


def registrar_fonte(func: FonteLinhaDoTempo) -> FonteLinhaDoTempo:
    """Registra `func` como fonte (usável como decorador). Registrar de novo o mesmo
    nome substitui a fonte sem mudar a posição dela."""
    _FONTES[func.__name__] = func
    return func


def fontes_registradas() -> list[str]:
    return list(_FONTES)


def carregar_fontes() -> None:
    """Importa os módulos de fontes (idempotente: o import só roda uma vez)."""
    for modulo in MODULOS_DE_FONTES:
        importlib.import_module(modulo)


@dataclass
class LinhaDoTempo:
    eventos: list[Evento] = field(default_factory=list)
    fontes_com_erro: list[str] = field(default_factory=list)


def montar_linha_do_tempo(item_id: int) -> LinhaDoTempo:
    """Eventos de todas as fontes em ordem cronológica. Uma fonte que levanta exceção
    é registrada no log e listada em `fontes_com_erro`; as outras continuam."""
    carregar_fontes()
    eventos: list[tuple[Evento, int]] = []
    com_erro: list[str] = []
    for posicao, (nome, fonte) in enumerate(list(_FONTES.items())):
        try:
            eventos.extend((e, posicao) for e in fonte(item_id))
        except Exception:
            logger.exception("Fonte %s da linha do tempo falhou (item %s).", nome, item_id)
            com_erro.append(nome)
    eventos.sort(key=lambda par: (par[0].ocorrido_em, par[1], par[0].referencia_id or 0))
    return LinhaDoTempo(eventos=[e for e, _ in eventos], fontes_com_erro=com_erro)


def item_existe(item_id: int) -> bool:
    """Existe o equipamento (ativo ou não): equipamento baixado mantém a linha do tempo."""
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT 1 AS ok FROM items WHERE id=%s", (item_id,))
        return cur.fetchone() is not None
