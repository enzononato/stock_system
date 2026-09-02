# Work Plan 004 — Frontend Adjustments from Reference Images

## Date

2026-09-02

## Objective

Implement all 7 visual and functional corrections defined by the reference images in `IMAGENS PARA O ANTIGRAVITY`.

## Scope

FRONTEND ONLY.

## Status

Completed.

## Checklist

- [x] **Item 1 (Image 1)**: Implement tree-style expandable/collapsible accordion navigation for groups (`Inventário`, `Operação`, `Gestão`) in Sidebar.
- [x] **Item 2 (Image 2)**: Relocate sidebar collapse/toggle trigger to the top header as a hamburger button (`Menu` icon) and remove bottom "Recolher" button.
- [x] **Item 3 (Image 3)**: Add Revenda / Unidade filter to `ChartsPage.tsx` (Indicadores) alongside Year and Month.
- [x] **Item 4 (Image 4)**: Add filters (Tipo, Status/Disponibilidade, Busca) to Peripherals table in `PeripheralsPage.tsx`.
- [x] **Item 5 (Image 5)**: Fix Revenda dropdown options in `ReportPage.tsx` so all available revendas/unidades are populated.
- [x] **Item 6 (Image 6)**: Adjust proportions and layout of the two tables in `ReturnPage.tsx` with balanced columns, search, and `clientPageSize={7}`.
- [x] **Item 7 (Image 7)**: Limit `HistoryPage.tsx` pagination to 7 items per page (`PAGE_SIZE = 7`).
- [x] Validate all changes with `npm run build`.
- [x] Update `estado.md`.
- [x] Create `notas/auditoria/auditoria-004.md`.


## Constraints

- FRONTEND ONLY: strictly zero modifications to backend or database.
- Preserve all existing routes, roles, and business features.
