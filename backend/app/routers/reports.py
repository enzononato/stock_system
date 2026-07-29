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

    output = io.StringIO()
    if rows:
        writer = csv.DictWriter(output, fieldnames=rows[0].keys())
        writer.writeheader()
        for row in rows:
            # Converte datetime para string para serialização CSV
            writer.writerow({k: str(v) if v is not None else "" for k, v in row.items()})

    output.seek(0)
    filename = f"relatorio_{year}_{month:02d}.csv"
    return StreamingResponse(
        iter([output.getvalue()]),
        media_type="text/csv",
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
