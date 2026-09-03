# Work Plan 005 — Full Execution of Prompt Master (Frontend Redesign)

## Date

2026-09-02

## Objective

Execute all pending requirements from `notas/PROMPTS USADOS PARA STOCK SYSTEM/PROMPT_MASTER_ANTIGRAVITY.md`:
1. Reorganize sidebar navigation hierarchy so "Vincular Periféricos" is a submenu child of "Periféricos" under "Operação".
2. Implement pop-up/modal editing directly on `StockPage` (`EditItemModal.tsx`) with form validation and working "Confirmar" button.
3. Standardize pagination across all application tables to 7 rows per page (`PAGE_SIZE = 7` or `clientPageSize={7}`).
4. Validate frontend build with `npm run build`.
5. Produce Matriz de Funcionalidades (Section 4) and Relatório Final Obrigatório (Section 40).

## Scope

FRONTEND ONLY: strictly zero modifications to backend code or database configuration.

## Status

Completed.

## Checklist

- [x] **Item 1**: Nest `/link` ("Vincular periféricos") as a child of `/peripherals` in `frontend/src/components/app/nav.ts`.
- [x] **Item 2**: Create `frontend/src/components/app/EditItemModal.tsx` with base fields, type-specific fields, validation, and working Confirmar button.
- [x] **Item 3**: Integrate `EditItemModal` into `frontend/src/features/stock/StockPage.tsx`.
- [x] **Item 4**: Standardize tables to 7 rows per page:
  - [x] `StockPage.tsx` (`PAGE_SIZE = 7`)
  - [x] `PeripheralsPage.tsx` (`clientPageSize={7}`)
  - [x] `ReportPage.tsx` (`clientPageSize={7}`)
  - [x] `TermsPage.tsx` (`clientPageSize={7}` on both pending and active tables)
  - [x] `UsersPage.tsx` (`clientPageSize={7}`)
  - [x] `UnidadesPage.tsx` (`clientPageSize={7}`)
- [x] **Item 5**: Run `npm run build` in `frontend` (code 0, 0 errors).
- [x] **Item 6**: Update `estado.md`.
- [x] **Item 7**: Create `notas/auditoria/auditoria-005.md` with complete Feature Matrix and Final Audit Report.
