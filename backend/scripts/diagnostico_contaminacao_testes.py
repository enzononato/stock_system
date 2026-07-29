"""Diagnóstico SOMENTE-LEITURA de resíduo deixado pela suíte de testes no banco.

Contexto
--------
Durante a rodada de correções de 2026-07-29, a suíte de testes gravou no banco
apontado por `backend/.env` em vez do banco de teste. A causa foi um import de
`app.core.config` no nível de módulo dos testes unitários: como imports de
módulo acontecem na coleta do pytest, o singleton `settings` era construído a
partir do `.env` real antes de as variáveis `TEST_DB_*` serem aplicadas. A
limpeza entre testes usava conexão crua para o banco de teste, então nada era
removido de onde as escritas de fato aconteciam.

O `conftest.py` foi corrigido (hook `pytest_configure` + verificação de que o
`settings` materializado aponta para o banco de teste), mas os registros já
criados continuam onde foram gravados. Este script apenas os LOCALIZA.

Este script NÃO escreve, NÃO apaga e NÃO altera nada. A remoção, se for o caso,
é decisão do responsável pelo sistema.

Uso:
    cd backend && python scripts/diagnostico_contaminacao_testes.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pymysql  # noqa: E402

from app.core.config import settings  # noqa: E402

# Assinaturas dos dados que as fixtures em backend/tests/conftest.py produzem.
USUARIOS_DE_TESTE = ("gestor.teste", "tecnico.teste", "aprendiz.teste")
PREFIXO_ITEM_TESTE = "SN-"
PREFIXO_PERIFERICO_TESTE = "PERIF-"
OPERADORES_DE_TESTE = ("fixture_teste",)


def main() -> int:
    print("=" * 72)
    print("DIAGNÓSTICO DE CONTAMINAÇÃO POR TESTES — SOMENTE LEITURA")
    print("=" * 72)
    print(f"Banco inspecionado: {settings.DB_HOST}:{settings.DB_PORT}/{settings.DB_NAME}")
    print()

    conn = pymysql.connect(
        host=settings.DB_HOST,
        port=settings.DB_PORT,
        user=settings.DB_USER,
        password=settings.DB_PASSWORD,
        database=settings.DB_NAME,
        charset=settings.DB_CHARSET,
        cursorclass=pymysql.cursors.DictCursor,
    )
    try:
        cur = conn.cursor()

        cur.execute("SHOW TABLES")
        tabelas = {list(r.values())[0] for r in cur.fetchall()}
        print("Contagem geral por tabela:")
        for t in ("usuarios", "items", "peripherals", "equipment_peripherals", "history", "revoked_tokens"):
            if t in tabelas:
                cur.execute(f"SELECT COUNT(*) AS n FROM `{t}`")
                print(f"  {t:<24} {cur.fetchone()['n']:>6}")
            else:
                print(f"  {t:<24} (tabela não existe)")
        print()

        marcadores = ", ".join(["%s"] * len(USUARIOS_DE_TESTE))
        cur.execute(f"SELECT id, username, role FROM usuarios WHERE username IN ({marcadores})", USUARIOS_DE_TESTE)
        usuarios = cur.fetchall()
        print(f"Usuários criados pelas fixtures: {len(usuarios)}")
        for u in usuarios:
            print(f"  id={u['id']:<5} {u['username']:<20} {u['role']}")
        print()

        cur.execute(
            "SELECT id, tipo, brand, model, identificador, status, is_active, date_registered "
            "FROM items WHERE identificador LIKE %s ORDER BY id",
            (PREFIXO_ITEM_TESTE + "%",),
        )
        itens = cur.fetchall()
        print(f"Itens com identificador de fixture ('{PREFIXO_ITEM_TESTE}...'): {len(itens)}")
        for i in itens[:20]:
            print(f"  id={i['id']:<5} {i['identificador']:<16} {i['tipo']}/{i['brand']} "
                  f"status={i['status']} ativo={i['is_active']} em {i['date_registered']}")
        if len(itens) > 20:
            print(f"  ... e mais {len(itens) - 20}")
        print()

        cur.execute(
            "SELECT id, tipo, brand, model, identificador, status, is_active "
            "FROM peripherals WHERE identificador LIKE %s ORDER BY id",
            (PREFIXO_PERIFERICO_TESTE + "%",),
        )
        perifericos = cur.fetchall()
        print(f"Periféricos com identificador de fixture ('{PREFIXO_PERIFERICO_TESTE}...'): {len(perifericos)}")
        for p in perifericos[:20]:
            print(f"  id={p['id']:<5} {p['identificador']:<18} {p['tipo']}/{p['brand']} status={p['status']}")
        if len(perifericos) > 20:
            print(f"  ... e mais {len(perifericos) - 20}")
        print()

        ids_itens = [i["id"] for i in itens]
        ids_perifericos = [p["id"] for p in perifericos]
        condicoes, params = [], []
        if ids_itens:
            condicoes.append(f"item_id IN ({', '.join(['%s'] * len(ids_itens))})")
            params.extend(ids_itens)
        if ids_perifericos:
            condicoes.append(f"peripheral_id IN ({', '.join(['%s'] * len(ids_perifericos))})")
            params.extend(ids_perifericos)
        marcadores_op = ", ".join(["%s"] * len(OPERADORES_DE_TESTE))
        condicoes.append(f"operador IN ({marcadores_op})")
        params.extend(OPERADORES_DE_TESTE)

        cur.execute(f"SELECT COUNT(*) AS n FROM history WHERE {' OR '.join(condicoes)}", params)
        print(f"Entradas de histórico ligadas a esses registros: {cur.fetchone()['n']}")
        print()

        cur.execute(
            "SELECT MIN(date_registered) AS primeiro, MAX(date_registered) AS ultimo "
            "FROM items WHERE identificador LIKE %s",
            (PREFIXO_ITEM_TESTE + "%",),
        )
        janela = cur.fetchone()
        if janela and janela["primeiro"]:
            print(f"Janela de criação dos itens de fixture: {janela['primeiro']} → {janela['ultimo']}")
        print()

        total = len(usuarios) + len(itens) + len(perifericos)
        if total == 0:
            print("RESULTADO: nenhum resíduo de teste encontrado neste banco.")
        else:
            print(f"RESULTADO: {total} registro(s) de teste encontrado(s). Nada foi alterado")
            print("por este script — a remoção é decisão do responsável pelo sistema.")

        cur.close()
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
