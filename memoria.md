# Project Memory

## Project Overview

Revalle IT stock and inventory management web application ("Controle de Estoque Revalle"), designed to manage enterprise hardware, peripherals, and related IT workflows.

## Architecture

Client-server architecture consisting of:
- A REST API backend built with FastAPI (Python) located in `backend/app`.
- A Single Page Application (SPA) frontend built with React, TypeScript, and Vite located in `frontend/`.
- Containerization setup via Docker Compose.

## Technology Stack

- **Backend:** Python, FastAPI, Uvicorn
- **Frontend:** React, TypeScript, Vite, TailwindCSS / Shadcn/ui
- **Database:** MySQL
- **DevOps / Environment:** Docker, Docker Compose

## Important Modules

- `backend/app/`: Core REST API application endpoints, business logic, and database access.
- `backend/modelos/`: Word document (.docx) templates for generated legal/assignment terms.
- `frontend/src/features/stock/`: Inventory and stock management interface.
- `frontend/src/features/offboarding/`: Employee offboarding / equipment return workflow.
- `frontend/src/api/`: API integration clients (e.g., peripherals, stock, authentication).

## Important Rules

- Existing MySQL database schema must be reused without arbitrary schema modifications.
- Required backend environment variables: `DB_HOST`, `DB_USER`, `DB_PASSWORD`, `DB_NAME`, and `JWT_SECRET`. The backend refuses to start without them.
- Preserve existing application functionality and project conventions.

## Important Decisions

- Replaced legacy desktop application (originally built in Tkinter) with a modern web application stack.
- Retained the existing MySQL database structure to ensure backward compatibility and smooth transition.
- Offboarding workflow operates via client-side state management (localStorage) for the 9-step governance pipeline (Gestor, RH, DP, TI, Patrimônio, Financeiro, Contabilidade) while relying on real backend REST endpoints for asset returns and loan terms.
- **Sidebar navigation** uses a tree-style accordion pattern: groups (Inventário, Operação, Gestão) are collapsible/expandable via `ChevronRight` icons. Sidebar collapse/expand is triggered by a hamburger (`Menu`) icon in the top header area.
- **Default table page size** is 7 rows for most data tables (`clientPageSize={7}` or `PAGE_SIZE = 7`), chosen to avoid vertical window scroll on standard viewports.
- **Dual-table layouts** (e.g. ReturnPage) use `flex flex-col flex-1` on `Section` and `DataTable` containers to ensure balanced card heights regardless of differing row counts.

## Integrations

Not yet confirmed.

## Authentication

- JWT-based authentication using `JWT_SECRET` configured via environment variables.

## Data Layer

- MySQL database, connecting directly to the pre-existing production/legacy schema.

## Known Constraints

- Schema alterations must be avoided to preserve compatibility with existing database records.
- Backend refuses startup if mandatory database credentials and JWT secrets are absent.

## Historical Context

- Project transitioned from a legacy Tkinter desktop application to a modern FastAPI + React web platform.

## Notes

- Backend typically runs via `uvicorn app.main:app --reload`.
- Frontend runs via Vite development server (`npm run dev`).
