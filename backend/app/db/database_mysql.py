"""
Camada de conexão com o MySQL.

Antes, cada método dos managers abria sua própria conexão
(`pymysql.connect(...)`) e a fechava ao final da chamada — sem pool, sem
reaproveitamento e, em vários métodos, sem `try/except`/`finally` garantindo
`rollback()`/`close()` em caso de erro.

Este módulo centraliza o acesso ao banco em duas peças:

- Um pool de conexões (`dbutils.pooled_db.PooledDB` sobre `pymysql`), criado de
  forma PREGUIÇOSA — só na primeira vez que alguém pedir uma conexão/cursor,
  nunca na importação do módulo. Isso permite que o processo da aplicação
  suba normalmente mesmo que o MySQL esteja temporariamente indisponível; o
  erro de conexão só ocorre quando a primeira query de fato for executada.
- `get_cursor()`, um context manager que entrega um cursor e garante
  commit (sucesso) / rollback (exceção) / close do cursor / devolução da
  conexão ao pool, sempre — eliminando a necessidade de cada método repetir
  esse boilerplate manualmente.

`get_connection()` é mantida EXATAMENTE com a assinatura e o comportamento
originais (uma conexão nova por chamada, via `pymysql.connect()` direto, sem
pool) para não quebrar nenhum código externo que ainda dependa dela. Todo
código deste módulo (`InventoryDBManager`, `UserDBManager`, `TokenDBManager`)
foi migrado para `get_cursor()` — `get_connection()` fica só como shim de
compatibilidade.
"""
from contextlib import contextmanager
from typing import Iterator, Optional

import pymysql
from pymysql.cursors import Cursor, DictCursor
from dbutils.pooled_db import PooledDB

from app.core.config import settings

# Pool "singleton" do processo. Não criar no import — apenas na primeira
# chamada a `_get_pool()` (ver docstring do módulo).
_pool: Optional[PooledDB] = None


def _get_pool() -> PooledDB:
    """Cria o pool de conexões na primeira chamada (lazy init) e o reutiliza
    nas chamadas seguintes."""
    global _pool
    if _pool is None:
        _pool = PooledDB(
            creator=pymysql,
            mincached=settings.DB_POOL_SIZE,
            maxconnections=settings.DB_POOL_MAX_CONNECTIONS,
            blocking=True,
            ping=1,
            host=settings.DB_HOST,
            port=settings.DB_PORT,
            user=settings.DB_USER,
            password=settings.DB_PASSWORD,
            database=settings.DB_NAME,
            charset=settings.DB_CHARSET,
            autocommit=False,
        )
    return _pool


def get_connection() -> pymysql.connections.Connection:
    """Obtém uma conexão direta ao MySQL — comportamento ORIGINAL preservado
    (uma conexão nova por chamada, sem pool) para compatibilidade com
    qualquer código legado que ainda dependa exatamente desta assinatura e
    deste comportamento.

    Código novo deve preferir `get_cursor()`: usa o pool e já garante
    commit/rollback/close corretamente. Quem usar `get_connection()`
    diretamente continua responsável por fazer commit/close manualmente.
    """
    return pymysql.connect(
        host=settings.DB_HOST,
        port=settings.DB_PORT,
        user=settings.DB_USER,
        password=settings.DB_PASSWORD,
        database=settings.DB_NAME,
        charset=settings.DB_CHARSET,
        cursorclass=pymysql.cursors.DictCursor,
    )


@contextmanager
def get_cursor(dict_cursor: bool = True, commit: bool = True) -> Iterator[Cursor]:
    """Entrega um cursor com commit/rollback/close garantidos.

    - `dict_cursor=True` (padrão): cursor retorna linhas como `dict`
      (`pymysql.cursors.DictCursor`), igual ao usado hoje na maioria das
      consultas. `dict_cursor=False` usa o cursor padrão (tuplas) para
      comandos que não leem resultado (INSERT/UPDATE/DELETE puros).
    - `commit=True` (padrão): ao sair do bloco `with` sem exceção, chama
      `conn.commit()`. Use `commit=False` para operações somente leitura.
    - Em qualquer exceção dentro do bloco `with`: chama `conn.rollback()` e
      relança a exceção original (o chamador decide como tratá-la).
    - Sempre (sucesso ou erro): fecha o cursor e devolve a conexão ao pool.
    """
    conn = _get_pool().connection()
    cursor_cls = DictCursor if dict_cursor else Cursor
    cur = conn.cursor(cursor_cls)
    try:
        yield cur
        if commit:
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()
