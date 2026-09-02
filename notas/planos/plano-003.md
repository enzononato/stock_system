# Work Plan 003 — Frontend Only Enhancements

## Date

2026-09-02

## Objective

Refine frontend loan and stock document workflows, prevent 404 errors on unavailable items without active loans in TermsPage/ReturnPage, integrate direct term access into ItemDetailsModal, and maintain complete frontend documentation continuity.

## Scope

FRONTEND ONLY.

## Status

Completed.

## Checklist

- [x] Refine active loans filtering in `TermsPage.tsx` (require `assigned_to` for active loans)
- [x] Refine returnable items filtering in `ReturnPage.tsx` (require `assigned_to` for returnable loans)
- [x] Add direct term access ("Ver Termo Assinado" / "Baixar Termo") in `ItemDetailsModal.tsx`
- [x] Run `npm run build` in `frontend` to validate types and assets
- [x] Update `estado.md`
- [x] Update `memoria.md` if permanent conventions established
- [x] Mark checklist in `plano-003.md` as completed
- [x] Create `notas/auditoria/auditoria-003.md`


## Constraints

- FRONTEND ONLY: strictly zero modifications to backend, database, or API server.
- Preserve existing application behavior and conventions.
