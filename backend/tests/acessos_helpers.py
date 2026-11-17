"""Auxiliares dos testes do módulo de acessos (F8).

Todos importam `app.db.database_mysql` dentro da função (regra do conftest: nada de
importar a camada de dados antes de o ambiente de teste ser configurado).
"""
import json
from datetime import datetime

CPF_MARIA = "11144477735"
CPF_JOAO = "52998224725"


def criar_colaborador(
    nome: str = "Maria Silva",
    cpf: str = CPF_MARIA,
    cargo: "str | None" = "Vendedor",
    situacao: str = "ativo",
) -> int:
    from app.db.database_mysql import get_cursor

    with get_cursor() as cur:
        cur.execute(
            "INSERT INTO colaboradores (cpf, nome, cargo, situacao, origem, criado_em, atualizado_em) "
            "VALUES (%s, %s, %s, %s, 'manual', NOW(), NOW())",
            (cpf, nome, cargo, situacao),
        )
        return cur.lastrowid


def criar_sistema(nome: str, ativo: bool = True) -> int:
    from app.db.database_mysql import get_cursor

    with get_cursor() as cur:
        cur.execute("INSERT INTO sistemas (nome, ativo) VALUES (%s, %s)", (nome, 1 if ativo else 0))
        return cur.lastrowid


def definir_pacote(cargo: str, *sistema_ids: int) -> None:
    from app.db.database_mysql import get_cursor

    with get_cursor() as cur:
        cur.execute("DELETE FROM pacotes_funcao WHERE cargo = %s", (cargo,))
        for sistema_id in sistema_ids:
            cur.execute(
                "INSERT INTO pacotes_funcao (cargo, sistema_id) VALUES (%s, %s)", (cargo, sistema_id)
            )


def criar_acesso(
    colaborador_id: int,
    sistema_id: int,
    login: str,
    status: str = "ativo",
    item_id: "int | None" = None,
) -> int:
    from app.db.database_mysql import get_cursor

    desativado = status == "desativado"
    with get_cursor() as cur:
        cur.execute(
            "INSERT INTO colaborador_acessos (colaborador_id, sistema_id, login, status, item_id, "
            "criado_por, criado_em, desativado_por, desativado_em) "
            "VALUES (%s, %s, %s, %s, %s, 'fixture', NOW(), %s, %s)",
            (
                colaborador_id, sistema_id, login, status, item_id,
                "fixture" if desativado else None,
                datetime.now() if desativado else None,
            ),
        )
        return cur.lastrowid


def vincular_item(
    item_id: int,
    *,
    colaborador_id: "int | None" = None,
    cpf: "str | None" = None,
    status: str = "Indisponível",
) -> None:
    from app.db.database_mysql import get_cursor

    with get_cursor() as cur:
        cur.execute(
            "UPDATE items SET colaborador_id=%s, cpf=%s, status=%s WHERE id=%s",
            (colaborador_id, cpf, status, item_id),
        )


def contar(tabela: str, where: str = "1=1", params: tuple = ()) -> int:
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(f"SELECT COUNT(*) AS n FROM {tabela} WHERE {where}", params)
        return cur.fetchone()["n"]


def auditoria(entidade: str, acao: "str | None" = None) -> list[dict]:
    """Linhas de `audit_log` (ordem de `id`), com `antes`/`depois` já decodificados."""
    from app.db.database_mysql import get_cursor

    sql = "SELECT * FROM audit_log WHERE entidade = %s"
    params: list = [entidade]
    if acao is not None:
        sql += " AND acao = %s"
        params.append(acao)
    with get_cursor(commit=False) as cur:
        cur.execute(sql + " ORDER BY id", tuple(params))
        linhas = cur.fetchall()
    for linha in linhas:
        for campo in ("antes", "depois"):
            if isinstance(linha[campo], str):
                linha[campo] = json.loads(linha[campo])
    return linhas
