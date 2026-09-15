import csv
import io
from datetime import datetime
from fastapi import APIRouter, HTTPException, Depends
from fastapi.responses import StreamingResponse

from app.schemas.reports import ReportRow, ChartData
from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import get_current_user, gestor_or_tecnico, gestor_only, get_inventory_db, CurrentUser

router = APIRouter(prefix="/api/reports", tags=["reports"])


@router.get("/monthly", response_model=list[ReportRow])
def monthly_report(
    year: int = None,
    month: int = None,
    _: CurrentUser = Depends(gestor_or_tecnico),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    now = datetime.now()
    year = year or now.year
    month = month or now.month
    if not (1 <= month <= 12):
        raise HTTPException(status_code=400, detail="Mês inválido.")
    rows = inv.generate_monthly_report(year, month)
    return rows


@router.get("/monthly/export")
def export_monthly_report(
    year: int = None,
    month: int = None,
    _: CurrentUser = Depends(gestor_only),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    now = datetime.now()
    year = year or now.year
    month = month or now.month
    rows = inv.generate_monthly_report(year, month)

    # Mapeamento de chaves internas -> cabeçalhos em português
    HEADERS_PT = {
        "history_id": "ID Histórico",
        "item_id": "ID Item",
        "peripheral_id": "ID Periférico",
        "operador": "Operador",
        "usuario": "Usuário",
        "cpf": "CPF",
        "cargo": "Cargo",
        "center_cost": "Centro de Custo",
        "setor": "Setor",
        "fornecedor": "Fornecedor",
        "revenda": "Revenda",
        "details": "Detalhes",
        "data_emprestimo": "Data Empréstimo",
        "data_confirmacao": "Data Confirmação",
        "data_devolucao": "Data Devolução",
        "operation_type": "Operação",
        "tipo": "Tipo",
        "brand": "Marca",
        "model": "Modelo",
        "identificador": "Identificador",
        "nota_fiscal": "Nota Fiscal",
    }

    output = io.StringIO()
    # BOM UTF-8: garante que o Excel abre sem problema de encoding
    output.write('\ufeff')

    if rows:
        raw_keys = list(rows[0].keys())
        pt_headers = [HEADERS_PT.get(k, k) for k in raw_keys]
        # Ponto-e-vírgula como delimitador (padrão Excel Brasil/Portugal)
        writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_ALL)
        writer.writerow(pt_headers)
        for row in rows:
            writer.writerow([str(v) if v is not None else "" for v in row.values()])

    output.seek(0)
    filename = f"relatorio_{year}_{month:02d}.csv"
    return StreamingResponse(
        iter([output.getvalue()]),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/charts/loans", response_model=ChartData)
def chart_loans(
    year: int = None,
    month: int = None,
    _: CurrentUser = Depends(get_current_user),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    now = datetime.now()
    year = year or now.year
    month = month or now.month
    days, issues, returns = inv.get_issue_return_counts(year, month)
    return ChartData(days=days, values=issues, values2=returns)


@router.get("/charts/registrations", response_model=ChartData)
def chart_registrations(
    year: int = None,
    month: int = None,
    _: CurrentUser = Depends(get_current_user),
    inv: InventoryDBManager = Depends(get_inventory_db),
):
    now = datetime.now()
    year = year or now.year
    month = month or now.month
    days, registrations = inv.get_registration_counts(year, month)
    return ChartData(days=days, values=registrations)
