# Patrimônio — Ajustes e melhorias — Plano mestre

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Entregar, numa única branch (`feat/patrimonio-melhorias`), todas as melhorias do cartão do Trello que o spec marca como "entra", com testes e sem quebrar o que está em produção.

**Architecture:** Fase 0 cria a base (migrações, configuração, e-mail, auditoria, colaboradores, agendador, locais, exportações, componentes de UI). As fases 1–6 e 8 rodam em paralelo em worktrees próprias a partir da F0; a F7 roda por último. Cada fase tem seu plano detalhado em `docs/superpowers/plans/2026-11-12-patrimonio-fN-*.md`.

**Tech Stack:** FastAPI + PyMySQL/DBUtils (SQL cru) + MySQL 8; Postgres de RH via psycopg 3 (somente leitura); React 18 + TS + Tailwind + TanStack Query + Vitest/MSW.

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md`

## Global Constraints

- Tudo que o spec marca como "Fora" continua fora (MDM, travar equipamento, WhatsApp, canal do TI, script, perfis por unidade, consulta de CPF na Receita, revisão de acessos pelo superior).
- Banco: só mudanças aditivas, sempre por migração `backend/migrations/NNNN_*.sql` + `.down.sql`, na faixa da fase (F0 `0001–0099`, F1 `0100–0199`, F2 `0200–0299`, F3 `0300–0399`, F4 `0400–0499`, F5 `0500–0599`, F6 `0600–0699`, F7 `0700–0799`, F8 `0800–0899`).
- Base de RH somente leitura; sem data de nascimento; CPF mascarado em auditoria/logs; QR sem dado pessoal; nenhuma senha de terceiros no módulo de Acessos.
- Toda escrita nova grava auditoria com `audit.record` (F0).
- Paleta `#0131CD` / `#FFFFFF` / `#B9FC01` pelos tokens da F0; nenhuma cor nova fora dos tokens.
- Listagens novas usam `Pagination` + `usePageSize` (25/50/100), `StatusPill` e `Button loading` da F0.
- Unidade/revenda exibidas com `upper()`; nunca regravadas em maiúsculas.
- Textos de interface, mensagens de erro e commits em português; commits `tipo(escopo): descrição` terminando com `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- Dependências novas além das já previstas (`psycopg[binary]`, `pyotp`, `openpyxl`, `fpdf2` no backend; `qrcode` e `@zxing/browser` no frontend, F6) só com justificativa no relatório da fase.

## Review Focus

- Empréstimo/devolução antigos (sem colaborador do cadastro, sem RevSign) continuam funcionando ponta a ponta — o fluxo manual de termo é o plano B em produção.
- Duas fases mexendo no mesmo arquivo compartilhado (`App.tsx`, `navigation.ts`, `main.py`, `jobs.py`, `dependencies.py`, `.env.example`, `inventory_manager_db.py`) → só acréscimos de linhas; a integração resolve conflitos preservando os dois lados e roda as duas suítes inteiras.
- Banco de produção com dados antigos (itens sem `colaborador_id`, unidades sem `alias_rh`, usuários sem e-mail) → telas e endpoints tratam `NULL` sem erro.
- Uso no celular (390 px) das telas novas de inventário, colaboradores e detalhe do equipamento.
- Usuário sem permissão chamando endpoint novo → 403 com a mensagem padrão (`Acesso restrito ao Gestor.` / `Acesso restrito a Gestor ou Técnico.`).

## Ambiente de execução (comum a todas as fases)

- Worktrees fora do Google Drive: `C:/Users/enzo.jz/dev/patrimonio/<nome>`; integração em `C:/Users/enzo.jz/dev/patrimonio/integracao` (branch `feat/patrimonio-melhorias`).
- Branch de cada fase: `fase/fN-<nome>` criada a partir de `feat/patrimonio-melhorias` depois da F0 integrada.
- Backend: `C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh <worktree> <sufixo> [args do pytest]` — container `stock-system-backend-test:dev` (python 3.12, deps já instaladas), MySQL de teste `stock_sys_mysql_test` (porta 3307, database `stock_sys_test_<sufixo>` exclusivo da worktree), Postgres `dev-postgres` (5433) para a base fictícia de RH (database `colab_test_<sufixo>`, apagado pela fixture). Sufixo = nome da fase (`f0b`, `f1`, `f2`...).
- Frontend: `npm ci` uma vez por worktree; `npx vitest run`, `npx tsc --noEmit`, `npm run build` em `frontend/`.
- Linha de base antes da F0: backend 235 testes passando; frontend 174 testes passando, `tsc` limpo.

## Fases

| Fase | Plano | Depende de | Arquivos compartilhados que toca |
|---|---|---|---|
| F0 Base | `2026-11-12-patrimonio-f0-base.md` | — | todos (é a fundação) |
| F1 Visual e listagens | `2026-11-12-patrimonio-f1-visual.md` | F0 | páginas existentes, `ItemDetailsModal`, `StockPage` |
| F2 Login | `2026-11-12-patrimonio-f2-login.md` | F0 | `routers/auth.py`, `routers/users.py`, `AuthContext`, `LoginPage`, `UsersPage` |
| F3 Termos e RevSign | `2026-11-12-patrimonio-f3-termos.md` | F0 | `routers/loans.py`, `inventory_manager_db.py` (empréstimo/devolução), `LoanPage`, `ReturnPage`, `jobs.py` |
| F4 Ciclo de vida e manutenção | `2026-11-12-patrimonio-f4-ciclo-vida.md` | F0 | `routers/items.py`, `schemas/items.py`, `RegisterItemPage`, `RemovePage`, ENUM de `items.status` |
| F5 Linhas telefônicas | `2026-11-12-patrimonio-f5-linhas.md` | F0 | só arquivos novos + rotas/menu |
| F6 Inventário | `2026-11-12-patrimonio-f6-inventario.md` | F0 | `StockPage` (seleção para etiquetas), rotas/menu |
| F8 Acessos | `2026-11-12-patrimonio-f8-acessos.md` | F0 | só arquivos novos + rotas/menu + aba do colaborador |
| F7 Painel, alertas e exportações | `2026-11-12-patrimonio-f7-painel.md` | F0, F3, F4, F5, F6 integradas | `DashboardPage`, exportações nas telas das outras fases |

## Protocolo de execução

- [ ] **Passo 1 — F0:** executar `2026-11-12-patrimonio-f0-base.md` em duas trilhas paralelas: backend (Tasks 1–8, worktree `f0-back`) e frontend (Tasks 9–13, worktree `f0-front`; a trilha de frontend usa MSW com os contratos do plano, sem depender do backend rodando). Revisão de conformidade ao spec e de qualidade em cada trilha; corrigir; merge das duas trilhas na integração; suítes completas verdes.
- [ ] **Passo 2 — planos das fases 1–6 e 8:** cada agente de fase, na sua worktree, primeiro escreve o plano da fase (formato superpowers:writing-plans, a partir do spec e do código já com a F0) e depois o executa tarefa a tarefa com TDD e um commit por tarefa.
- [ ] **Passo 3 — revisão por fase:** para cada fase concluída, um revisor independente confere o resultado contra o spec (seção da fase + matriz da seção 3) e outro revisa a qualidade do código; achados corrigidos pelo agente da fase.
- [ ] **Passo 4 — integração:** merge das fases na branch de integração, uma por vez (F5, F8, F2, F6, F4, F1, F3 — das que tocam menos arquivos compartilhados para as que tocam mais), rodando backend e frontend completos depois de cada merge.
- [ ] **Passo 5 — F7:** escrever e executar o plano da F7 sobre a integração já com as outras fases; revisão.
- [ ] **Passo 6 — fechamento:** revisão final do branch inteiro, teste de fumaça (backend + frontend locais: login com 2FA, colaboradores, empréstimo com termo pelo cliente falso do RevSign, manutenção, linha, inventário, acessos, painel, exportações), atualizar `README.md` (variáveis novas, deploy na rede do EasyPanel, comandos de migração), remover containers/databases de teste e relatar ao Enzo.
