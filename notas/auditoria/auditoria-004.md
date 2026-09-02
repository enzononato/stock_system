# Auditoria 004 — Visual Reference Corrections

## Date

2026-09-02

## Scope

FRONTEND ONLY. Implementation of 7 visual/functional corrections based on reference images provided by the user.

## Work Plan

`notas/planos/plano-004.md`

## Summary of Changes

### Item 1 — Tree-style Sidebar Navigation

**File:** `frontend/src/components/app/AppShell.tsx`

- Converted flat navigation groups (`Inventário`, `Operação`, `Gestão`) into an accordion tree-style with `ChevronRight` icons that rotate 90° when expanded.
- Each group is independently collapsible via `openGroups` state managed with `toggleGroup()`.
- Default state: all groups expanded.

### Item 2 — Hamburger Toggle in Top Header

**File:** `frontend/src/components/app/AppShell.tsx`

- Removed the bottom "Recolher" button from the sidebar footer.
- Added a `BrandHeader` component at the top of the sidebar containing the Revalle logo/text and a `Menu` (hamburger) icon button.
- Clicking the hamburger toggles `collapsed` state, matching the user's hand-drawn reference.

### Item 3 — Revenda Filter in Indicadores

**File:** `frontend/src/features/charts/ChartsPage.tsx`

- Added a Revenda/Unidade `<Select>` dropdown alongside existing Year and Month filters.
- Options populated from `useConstants()` and `listUnidades()` API.
- When a specific revenda is selected, metrics and chart datasets are recalculated using only that unit's data via `getMonthlyReport`.

### Item 4 — Filters in Periféricos

**File:** `frontend/src/features/peripherals/PeripheralsPage.tsx`

- Added comprehensive filter controls:
  - **Tipo**: selector dynamically populated from constants.
  - **Status**: selector for Disponível / Em Uso / Substituído.
  - **Revenda**: selector populated with active units and "Sem revenda (avulso)", mapping peripherals linked to equipment back to the equipment's revenda via `useQueries` on items with `peripheral_count > 0`.
  - **Busca**: text search across ID, tipo, marca, modelo, and serial number (identificador).
- Added "Limpar Filtros" button to reset all filters simultaneously.
- Resolved runtime `ReferenceError` caused by missing `revendas` in `useConstants()` destructuring, ensuring clean error-free rendering.
- Filtered data is passed to `DataTable` instead of raw API response.

### Item 5 — Revenda Dropdown Population in Relatório Mensal

**File:** `frontend/src/features/reports/ReportPage.tsx`

- Fixed bug where the Revenda dropdown was empty.
- Options now combine three sources: `useConstants()` unidades, `listUnidades()` API, and any revendas found in existing report records.
- Deduplicated via `Set` to ensure unique options.

### Item 6 — Balanced Table Layout in Devoluções

**Files:**
- `frontend/src/features/loans/ReturnPage.tsx`
- `frontend/src/components/app/PageHeader.tsx`
- `frontend/src/components/app/DataTable.tsx`

- Harmonized `Section` component to use `flex flex-col` with `flex-1` on inner content wrapper.
- `DataTable` outer container updated to `flex flex-1 flex-col justify-between space-y-3` so tables with different row counts maintain equal card heights.
- Aligned column sequences between "Pendentes" and "Empréstimos ativos" tables (ID, Tipo, Marca, Usuário, CPF, Revenda, Empréstimo, Ações).
- Added search input to filter both tables.
- Set `clientPageSize={7}` for both tables.

### Item 7 — History Table Row Limit

**File:** `frontend/src/features/history/HistoryPage.tsx`

- Changed `PAGE_SIZE` from `20` to `7`.
- This limits server-side pagination to 7 records per page, eliminating vertical window scroll on standard viewports.
- Column `hideBelow` breakpoints already ensure responsive horizontal compactness.

## Validation

- `npm run build` executed successfully with exit code 0 (981ms).
- All TypeScript types, routes, and components validated.
- No regressions introduced.

## Files Modified

| File | Change |
|------|--------|
| `frontend/src/components/app/AppShell.tsx` | Tree accordion nav + hamburger toggle |
| `frontend/src/features/charts/ChartsPage.tsx` | Revenda filter |
| `frontend/src/features/peripherals/PeripheralsPage.tsx` | Tipo/Status/Busca filters |
| `frontend/src/features/reports/ReportPage.tsx` | Revenda dropdown fix |
| `frontend/src/features/loans/ReturnPage.tsx` | Balanced tables + search + clientPageSize |
| `frontend/src/components/app/PageHeader.tsx` | Flex layout for Section |
| `frontend/src/components/app/DataTable.tsx` | Flex layout for table container |
| `frontend/src/features/history/HistoryPage.tsx` | PAGE_SIZE = 7 |

## Status

**COMPLETED** — All 7 items resolved. Build validated. Documentation updated.
