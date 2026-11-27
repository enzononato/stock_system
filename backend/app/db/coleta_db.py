"""Coletas de dados da máquina (`coletas_maquina`).

Só o SHA-256 do código é guardado. Datas em UTC, sempre passadas pelo Python
(ou por `UTC_TIMESTAMP()` quando a comparação é feita no próprio SQL).
"""
import json
from datetime import datetime, timedelta, timezone

from app.db.database_mysql import get_cursor


def _agora_utc() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def _desserializar(linha: dict | None) -> dict | None:
    if linha is not None and isinstance(linha.get("dados"), (str, bytes, bytearray)):
        linha["dados"] = json.loads(linha["dados"])
    return linha


def _normalizar_mac(mac: str | None) -> str:
    return (mac or "").replace(":", "").replace("-", "").strip().upper()


def _sql_mac(coluna: str) -> str:
    """O mesmo que `_normalizar_mac`, feito no banco."""
    return f"UPPER(REPLACE(REPLACE(TRIM({coluna}), ':', ''), '-', ''))"


class ColetaDBManager:
    def criar(self, codigo_hash: str, criado_por_id: int, expira_em: datetime) -> int:
        with get_cursor() as cur:
            cur.execute(
                "INSERT INTO coletas_maquina (codigo_hash, criado_por_id, criado_em, expira_em) "
                "VALUES (%s, %s, %s, %s)",
                (codigo_hash, criado_por_id, _agora_utc(), expira_em),
            )
            return cur.lastrowid

    def por_hash(self, codigo_hash: str) -> dict | None:
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT * FROM coletas_maquina WHERE codigo_hash = %s", (codigo_hash,))
            return _desserializar(cur.fetchone())

    def por_id(self, coleta_id: int) -> dict | None:
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT * FROM coletas_maquina WHERE id = %s", (coleta_id,))
            return _desserializar(cur.fetchone())

    def registrar_dados(self, codigo_hash: str, dados: dict, ip: str) -> bool:
        """Grava os dados uma única vez, e só antes de vencer. Atômico: devolve False
        se o código não existe, já foi usado ou venceu (nada é sobrescrito)."""
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "UPDATE coletas_maquina SET dados = %s, recebido_em = UTC_TIMESTAMP(), recebido_ip = %s "
                "WHERE codigo_hash = %s AND recebido_em IS NULL AND expira_em > UTC_TIMESTAMP()",
                (json.dumps(dados, ensure_ascii=False), ip, codigo_hash),
            )
            return cur.rowcount == 1

    def apagar_antigas(self, dias: int = 7) -> int:
        """Remove as coletas criadas há mais de `dias` dias. Devolve quantas."""
        limite = _agora_utc() - timedelta(days=dias)
        with get_cursor(dict_cursor=False) as cur:
            cur.execute("DELETE FROM coletas_maquina WHERE criado_em < %s", (limite,))
            return cur.rowcount

    def fornecedor_por_modelo(self, marca: str, modelo: str) -> str | None:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT DISTINCT TRIM(fornecedor) AS fornecedor FROM items "
                "WHERE UPPER(TRIM(brand)) = UPPER(%s) AND UPPER(TRIM(model)) = UPPER(%s) "
                "AND fornecedor IS NOT NULL AND TRIM(fornecedor) <> '' "
                "AND UPPER(TRIM(fornecedor)) NOT IN ('NAO INFORMADO', 'NÃO INFORMADO') LIMIT 2",
                (marca, modelo),
            )
            resultados = cur.fetchall()
        return resultados[0]["fornecedor"] if len(resultados) == 1 else None

    def correspondencias(self, serial: str | None, macs: list[str], anydesk: str | None) -> list[dict]:
        """Itens que já têm o mesmo serial, algum dos MACs ou o mesmo AnyDesk.

        Itens baixados (`is_active = 0`) também voltam, com `ativo=False`. Entradas
        vazias nunca casam com campos vazios."""
        serial_n = (serial or "").strip().upper()
        macs_n = sorted({m for m in (_normalizar_mac(m) for m in (macs or [])) if m})
        anydesk_n = "".join(c for c in (anydesk or "") if c.isdigit())

        # (motivo, expressão SQL, parâmetros): cada critério vira uma coluna 0/1 no
        # SELECT e uma condição do WHERE, para saber por que cada item casou.
        criterios: list[tuple[str, str, list]] = []
        if serial_n:
            criterios.append(("serial", "UPPER(TRIM(identificador)) = %s", [serial_n]))
        if macs_n:
            marcadores = ", ".join(["%s"] * len(macs_n))
            criterios.append((
                "mac",
                f"({_sql_mac('endereco_fisico')} IN ({marcadores}) "
                f"OR {_sql_mac('mac')} IN ({marcadores}))",
                macs_n + macs_n,
            ))
        if anydesk_n:
            criterios.append(("anydesk", "REGEXP_REPLACE(anydesk, '[^0-9]', '') = %s", [anydesk_n]))
        if not criterios:
            return []

        colunas = ", ".join(f"COALESCE({expr}, 0) AS m_{motivo}" for motivo, expr, _ in criterios)
        where = " OR ".join(expr for _, expr, _ in criterios)
        params = [p for _, _, ps in criterios for p in ps] * 2

        with get_cursor(commit=False) as cur:
            cur.execute(
                f"SELECT id, tipo, brand, model, status, revenda, is_active, {colunas} "
                f"FROM items WHERE {where} ORDER BY is_active DESC, id DESC",
                tuple(params),
            )
            linhas = cur.fetchall()

        return [
            {
                "id": linha["id"],
                "tipo": linha["tipo"],
                "brand": linha["brand"],
                "model": linha["model"],
                "status": linha["status"],
                "revenda": linha["revenda"],
                "ativo": bool(linha["is_active"]),
                "motivos": [m for m, _, _ in criterios if linha[f"m_{m}"]],
            }
            for linha in linhas
        ]
