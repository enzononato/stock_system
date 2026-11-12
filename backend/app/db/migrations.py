"""
Executor de migrações SQL versionadas.

As tabelas legadas continuam sendo criadas pelos managers (`CREATE TABLE IF NOT
EXISTS` no `__init__`). Toda mudança NOVA de schema entra como um par de arquivos
em `backend/migrations/` (`NNNN_descricao.sql` + `NNNN_descricao.down.sql`) e é
aplicada por este módulo, na ordem do número, uma única vez cada, registrando o
resultado em `schema_migrations`.

Cada versão ausente de `schema_migrations` é aplicada em ordem crescente, mesmo
que uma versão maior já esteja aplicada. O processo todo roda dentro de um
`GET_LOCK`, então vários workers subindo ao mesmo tempo não aplicam a mesma
migração duas vezes.

Atenção: DDL no MySQL faz commit implícito, então uma migração que falha no meio
pode deixar parte das alterações aplicadas. Escreva os `.sql` de forma
idempotente quando possível (veja backend/migrations/README.md).
"""
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)

MIGRATIONS_DIR: Path = Path(__file__).resolve().parent.parent.parent / "migrations"

_LOCK_NAME = "patrimonio_migrations"
_LOCK_TIMEOUT_SEGUNDOS = 30
_PADRAO_ARQUIVO = re.compile(r"^(\d{4})_(.+)\.sql$")


@dataclass
class Migracao:
    versao: int
    nome: str
    up_path: Path
    down_path: Optional[Path]


def listar_migracoes() -> list[Migracao]:
    """Migrações de `MIGRATIONS_DIR`, ordenadas por versão. `NNNN_nome.down.sql`
    é o par de reversão de `NNNN_nome.sql`; arquivos fora do padrão são ignorados."""
    migracoes: list[Migracao] = []
    for arquivo in sorted(Path(MIGRATIONS_DIR).glob("*.sql")):
        if arquivo.name.endswith(".down.sql"):
            continue
        m = _PADRAO_ARQUIVO.match(arquivo.name)
        if not m:
            continue
        down = arquivo.with_name(f"{arquivo.stem}.down.sql")
        migracoes.append(
            Migracao(
                versao=int(m.group(1)),
                nome=m.group(2),
                up_path=arquivo,
                down_path=down if down.exists() else None,
            )
        )
    return sorted(migracoes, key=lambda mig: mig.versao)


def garantir_baseline() -> None:
    """Instancia os managers que criam as tabelas legadas, para que elas existam
    antes de qualquer migração que as referencie."""
    from app.db.inventory_manager_db import InventoryDBManager
    from app.db.token_db import TokenDBManager
    from app.db.unidade_db import UnidadeDBManager

    InventoryDBManager()
    UnidadeDBManager()
    TokenDBManager()


def executar_sql(cur, texto: str) -> None:
    """Executa um script SQL no cursor, comando a comando.

    Os comandos são separados por `;` no fim da linha; linhas iniciadas por `--`
    são comentários e ignoradas.
    """
    comando: list[str] = []
    for linha in texto.splitlines():
        if linha.strip().startswith("--"):
            continue
        comando.append(linha)
        if linha.rstrip().endswith(";"):
            _executar_comando(cur, comando)
            comando = []
    _executar_comando(cur, comando)


def _executar_comando(cur, linhas: list[str]) -> None:
    sql = "\n".join(linhas).strip().rstrip(";").strip()
    if sql:
        cur.execute(sql)


def criar_tabela_de_controle(cur) -> None:
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS schema_migrations (
            versao INT PRIMARY KEY,
            nome VARCHAR(200) NOT NULL,
            aplicada_em DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )


def _obter_lock(cur) -> None:
    cur.execute("SELECT GET_LOCK(%s, %s) AS ok", (_LOCK_NAME, _LOCK_TIMEOUT_SEGUNDOS))
    if cur.fetchone()["ok"] != 1:
        raise RuntimeError(
            "Não foi possível obter o bloqueio de migrações "
            f"({_LOCK_NAME}) em {_LOCK_TIMEOUT_SEGUNDOS}s: outra instância está migrando."
        )


def _liberar_lock(cur) -> None:
    try:
        cur.execute("SELECT RELEASE_LOCK(%s)", (_LOCK_NAME,))
    except Exception:
        logger.exception("Falha ao liberar o bloqueio de migrações.")


def versoes_aplicadas(cur) -> set[int]:
    cur.execute("SELECT versao FROM schema_migrations")
    return {linha["versao"] for linha in cur.fetchall()}


def aplicar_pendentes() -> list[int]:
    """Aplica, em ordem crescente, toda migração ainda ausente de
    `schema_migrations`. Retorna as versões aplicadas nesta chamada."""
    aplicadas_agora: list[int] = []
    with get_cursor() as cur:
        _obter_lock(cur)
        try:
            criar_tabela_de_controle(cur)
            ja_aplicadas = versoes_aplicadas(cur)
            for mig in listar_migracoes():
                if mig.versao in ja_aplicadas:
                    continue
                logger.info("Aplicando migração %04d_%s", mig.versao, mig.nome)
                executar_sql(cur, mig.up_path.read_text(encoding="utf-8"))
                cur.execute(
                    "INSERT INTO schema_migrations (versao, nome, aplicada_em) VALUES (%s, %s, NOW())",
                    (mig.versao, mig.nome),
                )
                cur.execute("COMMIT")
                aplicadas_agora.append(mig.versao)
        finally:
            _liberar_lock(cur)
    return aplicadas_agora


def reverter_ultima() -> Optional[int]:
    """Reverte a migração de maior versão aplicada, usando o `.down.sql`.
    Retorna a versão revertida, ou None se não houver nenhuma aplicada."""
    with get_cursor() as cur:
        _obter_lock(cur)
        try:
            criar_tabela_de_controle(cur)
            cur.execute("SELECT MAX(versao) AS versao FROM schema_migrations")
            versao = cur.fetchone()["versao"]
            if versao is None:
                return None
            mig = next((m for m in listar_migracoes() if m.versao == versao), None)
            if mig is None or mig.down_path is None:
                raise RuntimeError(
                    f"Migração {versao:04d} não tem arquivo .down.sql; não é possível reverter."
                )
            logger.info("Revertendo migração %04d_%s", mig.versao, mig.nome)
            executar_sql(cur, mig.down_path.read_text(encoding="utf-8"))
            cur.execute("DELETE FROM schema_migrations WHERE versao=%s", (versao,))
            cur.execute("COMMIT")
            return versao
        finally:
            _liberar_lock(cur)
