"""Apoio dos testes da Fase 2 (login, 2FA, redefinição de senha, perfil).

Os testes da F2 importam o que usam daqui (`# noqa: F401`); estes fixtures não
entram no `conftest.py`.
"""
import json
from pathlib import Path


def linhas_auditoria(entidade: str, acao: str | None = None) -> list[dict]:
    """Linhas de `audit_log` da entidade (e da ação, se dada), em ordem de id,
    com `antes`/`depois` já decodificados."""
    from app.db.database_mysql import get_cursor

    sql = "SELECT * FROM audit_log WHERE entidade = %s"
    params: list = [entidade]
    if acao is not None:
        sql += " AND acao = %s"
        params.append(acao)
    sql += " ORDER BY id"
    with get_cursor(commit=False) as cur:
        cur.execute(sql, tuple(params))
        linhas = cur.fetchall()
    for linha in linhas:
        for chave in ("antes", "depois"):
            valor = linha.get(chave)
            if isinstance(valor, (str, bytes)):
                linha[chave] = json.loads(valor)
    return linhas


def ida_e_volta_migracao(versao: int) -> None:
    """Roda o `.down.sql` e depois o `.sql` da migração `versao` (o `.sql` roda
    num `finally`, para nunca deixar o banco de teste sem a migração)."""
    from app.db import migrations
    from app.db.database_mysql import get_cursor

    pasta = Path(migrations.MIGRATIONS_DIR)
    up = next(p for p in sorted(pasta.glob(f"{versao:04d}_*.sql")) if not p.name.endswith(".down.sql"))
    down = up.with_name(f"{up.stem}.down.sql")
    try:
        with get_cursor() as cur:
            migrations.executar_sql(cur, down.read_text(encoding="utf-8"))
    finally:
        with get_cursor() as cur:
            migrations.executar_sql(cur, up.read_text(encoding="utf-8"))
