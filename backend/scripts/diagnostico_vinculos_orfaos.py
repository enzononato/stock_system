"""
Script de DIAGNÓSTICO — SOMENTE LEITURA. Não escreve nada no banco.

Contexto: até a correção do bug em `InventoryDBManager.confirm_return`, a
devolução de um equipamento marcava os periféricos vinculados como
'Disponível' mas nunca apagava a linha correspondente em
`equipment_peripherals`. Resultado: o periférico ficava "Disponível" e, ao
mesmo tempo, continuava aparecendo vinculado ao equipamento antigo — podendo
inclusive ser vinculado a um segundo equipamento sem nunca ter sido
desvinculado do primeiro.

A correção impede que o problema aconteça daqui pra frente, mas não limpa
vínculos órfãos que já existem na base por causa do bug. Este script lista
esses registros para que alguém avalie manualmente a limpeza dos dados
históricos (o script em si não decide nem executa nada).

Uso:
    python backend/scripts/diagnostico_vinculos_orfaos.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from app.db.database_mysql import get_cursor  # noqa: E402


def listar_vinculos_orfaos() -> list[dict]:
    """Retorna os vínculos em equipment_peripherals cujo periférico já está
    marcado como 'Disponível' (ou seja: deveria ter sido desvinculado)."""
    with get_cursor(commit=False) as cur:
        cur.execute("""
            SELECT
                ep.id AS link_id,
                ep.equipment_id,
                i.identificador AS equipamento_identificador,
                i.status AS equipamento_status,
                p.id AS peripheral_id,
                p.tipo AS peripheral_tipo,
                p.brand AS peripheral_brand,
                p.model AS peripheral_model,
                p.identificador AS peripheral_identificador,
                p.status AS peripheral_status
            FROM equipment_peripherals ep
            JOIN peripherals p ON p.id = ep.peripheral_id
            LEFT JOIN items i ON i.id = ep.equipment_id
            WHERE p.status = 'Disponível'
            ORDER BY ep.equipment_id, p.id
        """)
        return cur.fetchall()


def main():
    rows = listar_vinculos_orfaos()

    if not rows:
        print("Nenhum vínculo órfão encontrado. Nada a limpar.")
        return

    print(
        f"Encontrado(s) {len(rows)} vínculo(s) órfão(s): periférico com "
        f"status='Disponível' que ainda aparece vinculado a um equipamento "
        f"em equipment_peripherals.\n"
    )
    for r in rows:
        equipamento_desc = r.get("equipamento_identificador") or "sem identificador"
        print(
            f"  link_id={r['link_id']:<6} "
            f"equipamento_id={r['equipment_id']} ({equipamento_desc}, "
            f"status={r.get('equipamento_status') or 'equipamento não encontrado'}) "
            f"<-> periférico_id={r['peripheral_id']} "
            f"{r['peripheral_tipo']} {r.get('peripheral_brand') or ''} {r.get('peripheral_model') or ''} "
            f"(S/N: {r.get('peripheral_identificador') or 'N/A'})"
        )

    print(
        "\nEste script NÃO altera o banco. Para limpar um vínculo confirmado "
        "como órfão, remova manualmente a linha correspondente, por exemplo:\n"
        "  DELETE FROM equipment_peripherals WHERE id IN (<link_id_1>, <link_id_2>, ...);\n"
        "Avalie caso a caso antes de apagar — confirme que o periférico realmente "
        "não está mais em uso pelo equipamento listado."
    )


if __name__ == "__main__":
    main()
