# Plano 007 — Redesign Enterprise Global (Prompt Mestre Final)

## 1. Identificação
- **Data:** 2026-09-03
- **Branch:** `redesign-frontend`
- **Escopo:** Execução do Prompt Mestre Final para o Redesign Enterprise completo do frontend, preservando 100% das APIs reais, banco de dados, JWT e rotas.

## 2. Objetivos
1. Transformar a interface em um produto enterprise corporativo (inspiração no fluxo experiencial do Spotify e identidade tecnológica da Solvd).
2. Limpar tokens órfãos e consolidar o Design System em OKLCH com paleta Deep Navy.
3. Incorporar Breadcrumbs dinâmicos, Command Palette (`Ctrl + K`) e indicador de status operacional no Header do AppShell.
4. Migrar badges de status para tokens semânticos (`badge-success`, `badge-warning`, etc.), eliminando cores literais hardcoded.
5. Elevar os `KpiCard`s com superfícies interativas, ícones temáticos e números tabulares na fonte técnica `IBM Plex Mono`.
6. Garantir 7 itens por página nas tabelas com scroll horizontal protegido.
7. Validar build e lint estáticos com zero erros de compilação.

## 3. Fases Executadas
- **Fase 0:** Auditoria completa e plano detalhado aprovado em `implementation_plan.md`.
- **Fase 1:** Limpeza em `tailwind.config.js`, importação oficial de `IBM Plex Mono` no `index.html` e utilitários semânticos em `styles.css`.
- **Fase 2:** Redesenho do `AppShell.tsx` com Header ativo (Breadcrumb + Command Dialog + Status + Perfil) e Sidebar Spotify com indicador ciano ativo.
- **Fase 3:** Manutenção e consolidação do elemento 3D (`logo.glb` Draco self-hosted) e fallback 2D.
- **Fase 4:** Componentes base refinados (`StatusBadge.tsx`, `KpiCard.tsx`, `PageHeader.tsx`, `DataTable.tsx`).
- **Fase 5:** Padronização visual em `StockPage`, `UsersPage`, `HistoryPage`, `PeripheralsPage`, `LinkPeripheralPage`, `TermsPage`, `ReportPage`.
- **Fase 6:** Responsividade e suporte nativo a `prefers-reduced-motion`.
- **Fase 7:** `npm run build` (Exit Code 0), `npm run format` (Exit Code 0) e `npx eslint src` (0 erros).
