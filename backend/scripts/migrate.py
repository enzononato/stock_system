"""
Executor de migrações do banco, via linha de comando.

Uso (a partir de backend/):
    python -m scripts.migrate            # aplica as pendentes (mesmo que "aplicar")
    python -m scripts.migrate aplicar    # aplica as pendentes
    python -m scripts.migrate pendentes  # lista as pendentes, sem aplicar
    python -m scripts.migrate reverter   # reverte a última migração aplicada
"""
import argparse
import logging
import sys

from app.db import migrations
from app.db.database_mysql import get_cursor


def _pendentes() -> list[migrations.Migracao]:
    with get_cursor() as cur:
        migrations.criar_tabela_de_controle(cur)
        aplicadas = migrations.versoes_aplicadas(cur)
    return [m for m in migrations.listar_migracoes() if m.versao not in aplicadas]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Migrações do banco do Patrimônio.")
    parser.add_argument(
        "comando", nargs="?", default="aplicar", choices=["aplicar", "pendentes", "reverter"]
    )
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    migrations.garantir_baseline()
    if args.comando == "aplicar":
        aplicadas = migrations.aplicar_pendentes()
        print("Migrações aplicadas: " + (", ".join(f"{v:04d}" for v in aplicadas) or "nenhuma"))
    elif args.comando == "pendentes":
        pendentes = _pendentes()
        print(
            "Migrações pendentes: "
            + (", ".join(f"{m.versao:04d}_{m.nome}" for m in pendentes) or "nenhuma")
        )
    else:
        versao = migrations.reverter_ultima()
        print(
            f"Migração revertida: {versao:04d}"
            if versao is not None
            else "Nenhuma migração aplicada para reverter."
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
