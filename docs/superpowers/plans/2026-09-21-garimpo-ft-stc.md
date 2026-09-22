# Garimpo da branch FT_STC — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Colher da `origin/FT_STC` tudo que agrega, reaplicando sobre a `main` atual, sem importar os defeitos dela nem perder o trabalho funcional já publicado.

**Architecture:** A `FT_STC` **NÃO será mesclada**. Ela partiu de `6e14eac`, antes do trabalho de paridade funcional, e reescreveu 15 páginas em paralelo. Merge produziria conflito em quase todo arquivo e reintroduziria `limit: 500` em 7 telas. O caminho é garimpo: ler com `git show origin/FT_STC:<caminho>` e reaplicar sobre a versão atual.

**Tech Stack:** React 18, TypeScript, Vite 5, Tailwind 3, react-router-dom 6, TanStack Query/Table, Radix UI, recharts, vitest + Testing Library.

**Spec:** Este documento, derivado de uma análise completa da branch.

## Global Constraints

- **NÃO mesclar, cherry-pickar nem dar checkout da `FT_STC`.** Só leitura via `git show`.
- **Preservar a engenharia da `main`**, que é melhor em quase toda tela: paginação e busca no servidor, debounce, `keepPreviousData`, o componente `DataTable`, o `StatusBadge` com ponto colorido, e `navigation.ts` como fonte única de navegação (consumida por sidebar, paleta Ctrl+K e breadcrumbs).
- **Nunca reintroduzir `limit: 500` nem `FETCH_ALL_LIMIT`.** `MAX_PAGE_SIZE` no backend é 500 — teto, não folga.
- **Corrigir todo defeito da FT_STC ao trazer o código.** A lista está abaixo; nenhum entra como está.
- Sem dependência npm nova. Sem sintaxe Tailwind 4. Sem variante `dark:`.
- **`shadow-xs` NÃO EXISTE no Tailwind 3** — é classe do Tailwind 4, descartada em silêncio. Confirmado varrendo o CSS gerado. Trocar por `shadow-sm` onde aparecer.
- **Nunca usar modificador de opacidade sobre token customizado** (`bg-primary/50`) — não gera CSS aqui; há teste-guarda em `src/lib/tokens.test.ts`. Cores nativas (`bg-black/50`) funcionam e são permitidas.
- Todo texto de interface e todo comentário em português do Brasil.
- **Exceção autorizada pelo dono:** esta rodada PODE alterar `backend/app/routers/reports.py`, e só esse arquivo, para trazer o CSV melhorado. Nenhum outro arquivo do backend.
- **Exceção autorizada pelo dono:** trocar a fonte do sistema de Plus Jakarta Sans para Inter.
- Cada task termina com `npx tsc --noEmit`, `npm run test:run` e `npx vite build` verdes.

**Linha de base:** 159 testes em 17 arquivos, tsc limpo, build limpo.

## Defeitos da FT_STC — nenhum entra como está

| # | Defeito | Onde (na FT_STC) |
|---|---|---|
| D1 | `['Computador', 'Notebook']` — o enum tem "Desktop", nunca "Computador". Esconde Hardware & Sistema de todo desktop | `ItemDetailsModal.tsx:223` |
| D2 | Checkbox não responde ao clique no quadrado quando falta `id` (o `<label htmlFor>` perde o alvo) | `checkbox.tsx:22-36`; atinge `ReturnPage:182,192,200` e `RemovePage:260` |
| D3 | `shadow-xs` não existe no Tailwind 3 | `Stepper.tsx:45`, `HistoryPage.tsx:359`, `TermsPage.tsx:169` |
| D4 | 12 dos 13 ícones não animam: o CSS declara os nomes originais do Lucide, nunca renomeados. E a animação dispara no `:hover` do SVG, não do item de menu | `animated-sidebar-icons.css:28-40` |
| D5 | Tempestade N+1: uma requisição por item, resultado nunca lido | `PeripheralsPage.tsx:109-115` |
| D6 | `cargo` e `center_cost` adicionados ao tipo `Item`, mas `ItemResponse` no backend não tem nenhum dos dois | `api/items.ts:38-39`, consumido em `TermsPage.tsx:268` |
| D7 | 22 declarações sem uso (`tsc --noUnusedLocals`), contra 5 na main | vários |
| D8 | `bg-emerald-500` cru | `Sidebar.tsx:118`, `DashboardPage.tsx:105` |

## O que NÃO trazer

Dashboard, ReportPage, PeripheralsPage, LinkPeripheralPage e TermsPage da FT_STC — cerca de 1.900 linhas onde a `main` já está melhor. O Dashboard dela conta os quatro totais com `.filter().length` sobre 500 itens; o da `main` lê o `total` de consultas por status. Também não trazer `relatorio_linguagens_versoes.md`, que contém erro factual sobre si mesmo.

---

### Task 1: Utilitários e componentes soltos

**Files:**
- Create: `frontend/src/lib/api-error.ts`, `frontend/src/components/ui/StateBlocks.tsx`, `frontend/src/components/ui/Stepper.tsx`, `frontend/src/components/ui/checkbox.tsx`, `frontend/src/components/ui/dialog.tsx`
- Modify: `frontend/src/lib/utils.ts`, `frontend/src/components/ui/toast.tsx`

- [ ] **Step 1:** Trazer `lib/api-error.ts` (`getErrorMessage`). Prioriza o `detail` do FastAPI, trata array de validação do Pydantic, `ERR_NETWORK`, 401/403/404, e nunca vaza stack.
- [ ] **Step 2:** Trazer `exportToCsv` para o fim de `lib/utils.ts` — BOM UTF-8, separador `;`, datas dd/mm/aaaa, decimal com vírgula, escape de aspas.
- [ ] **Step 3:** Trazer `StateBlocks.tsx` (`EmptyState`, `ErrorState` com "Tentar novamente", `LoadingState`), com `role="alert"`, `aria-live`, `aria-busy`.
- [ ] **Step 4:** Trazer `Stepper.tsx`. **Corrigir D3** (`shadow-xs` → `shadow-sm`) e trocar `rounded-[4px]` pelo token `rounded`.
- [ ] **Step 5:** Trazer `checkbox.tsx` **corrigindo D2**: gerar `id` interno com `useId()` quando a prop não vier. Escrever teste que clique no QUADRADO (não no texto) e afirme que `onCheckedChange` foi chamado — tem que falhar com o código original dela.
- [ ] **Step 6:** Trazer `dialog.tsx` (Radix, já é dependência). Trocar `shadow-xl` por `shadow-overlay`. Não remove o `alert-dialog.tsx` existente — são complementares.
- [ ] **Step 7:** Trazer `toast.success()` / `toast.error()` como açúcar sobre o `toast(msg, type)` atual.
- [ ] **Step 8:** Adotar `getErrorMessage` onde hoje há tratamento improvisado (`ConfirmacaoTermo.tsx:34-39` e similares). Listar no relatório cada lugar adotado.
- [ ] **Step 9:** Verificação e commit.

---

### Task 2: Ícones animados

**Files:** Create `frontend/src/components/animated-icons/**`; Modify `frontend/src/main.tsx`, `frontend/src/components/layout/navigation.ts`, `Sidebar.tsx`

- [ ] **Step 1:** Trazer os 13 ícones e o CSS.
- [ ] **Step 2: Corrigir D4.** O CSS declara `.animated-packageicon`, `.animated-layoutdashboardicon` etc. — os nomes originais do Lucide. Os componentes emitem `.animated-estoqueicon`, `.animated-dashboardicon`… Só `termosicon` casa. Alinhar os 13 nomes. Verifique lendo os dois lados, não confie nesta lista.
- [ ] **Step 3:** Ainda em D4: a animação dispara no `:hover` do SVG. Passar o mouse no texto do link não anima. Mudar para `group-hover` a partir do item de menu.
- [ ] **Step 4:** Remover o `import * as React` sem uso dos 13 arquivos (D7).
- [ ] **Step 5:** Plugar em `navigation.ts`, que hoje guarda `icon: LucideIcon` e renderiza `<item.icon size={18} />`. A API dos ícones dela é compatível. **Manter `navigation.ts` como fonte única** — sidebar, paleta e breadcrumbs consomem dali.
- [ ] **Step 6:** Verificação e commit. Confirmar no CSS gerado que as classes de animação aparecem.

---

### Task 3: DataTable com paginação no cliente

**Files:** `frontend/src/components/ui/DataTable.tsx`, `DataTable.test.tsx`

**Contexto:** a `main` nunca tocou esse arquivo; a adição é limpa. Hoje `UnidadesPage` e `UsersPage` renderizam a lista inteira sem rodapé — é a queixa de "página gigante".

- [ ] **Step 1:** Trazer a paginação client-side que age quando a prop `pagination` NÃO é passada. Não alterar o caminho server-side, que várias telas usam.
- [ ] **Step 2:** Padrão de página **10**, não 7 — alinhado com o resto do sistema.
- [ ] **Step 3:** Corrigir o defeito: `clientPage` não é resetado quando `data` muda por fora, só quando o filtro global muda. Dá para ficar numa página vazia.
- [ ] **Step 4:** Testes para o caminho client-side. Confirmar que os 8 testes existentes seguem passando.
- [ ] **Step 5:** Verificação e commit.

---

### Task 4: HistoryPage e UsersPage

**Files:** `frontend/src/pages/HistoryPage.tsx`, `UsersPage.tsx`, `HistoryPage.test.tsx`

**Contexto:** a `main` nunca tocou essas duas. É o maior ganho da branch e não há colisão.

- [ ] **Step 1:** Trazer a `HistoryPage` dela: linha do tempo com nó por evento e símbolo monocromático, card expansível individual e "Expandir todos", painel lateral de 280px com busca e sumário da página, exportação CSV (usando o `exportToCsv` da Task 1), e diálogo de estorno com senha distinguindo 403 de erro genérico. Ela já tem paginação no servidor e debounce de 400ms — **preservar**.
- [ ] **Step 2: Corrigir D3** (`shadow-xs` na linha 359 dela).
- [ ] **Step 3:** Confirmar que `HistoryPage.test.tsx` da `main` segue passando. Se quebrar por estrutura, ajustar a FORMA do matcher, nunca o significado, com comentário em português.
- [ ] **Step 4:** Trazer a `UsersPage` dela — cartão de cadastro separado, troca de senha inline, ícones por papel, `AlertDialog` na remoção. Ela já usa `PageHeader`/`PanelHeader`.
- [ ] **Step 5:** Verificação e commit.

---

### Task 5: Edição inline no ItemDetailsModal

**Files:** `frontend/src/components/equipment/ItemDetailsModal.tsx`, `ItemDetailsModal.test.tsx`

**Contexto:** é a maior funcionalidade nova da branch. O modal dela cresceu de ~355 para 776 linhas: botão "Editar" troca os campos por inputs — marca, modelo, nota fiscal com máscara e validação de 9 dígitos, código patrimonial, fornecedor, e os 21 campos específicos por tipo — salvando via `updateItem`.

- [ ] **Step 1:** Reaplicar a edição inline sobre o modal da `main`, preservando o que a `main` tem.
- [ ] **Step 2: Corrigir D1.** Manter `['Desktop', 'Notebook']`, os valores reais do enum. Na versão dela o bug deixa de ser cosmético: sem isso não dá nem para EDITAR CPU, RAM, storage, host, IP e licença de um desktop.
- [ ] **Step 3:** Trazer os três cartões de contexto do topo (Unidade / Usuário Atual / Identificador-NF), o download do termo assinado para itens indisponíveis e a geração de termo para pendentes.
- [ ] **Step 4:** O teste da `main` espera o título "Especificações de Hardware & Sistema"; a FT_STC renomeou a seção. Decidir qual título fica e alinhar teste e código — sem enfraquecer a asserção.
- [ ] **Step 5:** Verificação e commit.

---

### Task 6: Assistente de 4 passos no empréstimo

**Files:** `frontend/src/pages/LoanPage.tsx`

**Contexto:** o `Stepper` da Task 1 é a barra visual; toda a lógica do assistente mora na página. **Construir sobre a `LoanPage` da `main`**, que já tem o carregamento correto — a dela usa `limit: 500` e filtra no cliente.

- [ ] **Step 1:** Quatro etapas: Equipamento → Colaborador → Condições → Confirmação.
- [ ] **Step 2:** Validação por etapa — CPF no passo 2, revenda e setor no passo 3. Volta livre para etapas já concluídas; avanço só com a etapa válida.
- [ ] **Step 3:** Coluna de contexto fixa de 340px com o resumo do que já foi escolhido.
- [ ] **Step 4:** Preenchimento automático da revenda a partir do item.
- [ ] **Step 5:** **Não encostar no toggle "É pessoa jurídica"** — ele muda o texto legal do documento gerado.
- [ ] **Step 6:** Verificação e commit.

---

### Task 7: Refinamentos de tela

**Files:** `StockPage.tsx`, `ReturnPage.tsx`, `RemovePage.tsx`, `RegisterItemPage.tsx`, `ChartsPage.tsx`, `Sidebar.tsx`

- [ ] **Step 1 — StockPage:** trazer a densidade dela — colunas que somem por breakpoint (`Alocado Para` em `md`, `Unidade` em `lg`, `Cadastro` em `xl`), `X` para limpar a busca, ações no hover da linha. **Manter** o `DataTable`, o `StatusBadge` com ponto colorido, o debounce e o `keepPreviousData` da `main`. Não trazer o `MonoBadge` dela.
- [ ] **Step 2 — ReturnPage:** trazer o checklist de inspeção física (chassi, tela/teclado/fonte, periféricos e cabos) com o checkbox já corrigido na Task 1.
- [ ] **Step 3 — RemovePage:** trazer a ficha completa do ativo, a justificativa obrigatória por select e a confirmação de irreversibilidade — também com o checkbox corrigido. **Este é o portão da única ação irreversível do sistema**; escrever teste que clique no quadrado e confirme que o botão destrava.
- [ ] **Step 4 — RegisterItemPage:** trazer o seccionamento numerado do formulário com subtítulo por seção.
- [ ] **Step 5 — ChartsPage:** adotar `StateBlocks` nos estados de carregando, vazio e erro.
- [ ] **Step 6 — Sidebar:** trazer os grupos colapsáveis persistidos em `localStorage` e o bloco de marca com logo. **Manter `navigation.ts` como fonte única** — a Sidebar dela duplica a lista internamente, o que quebraria a paleta Ctrl+K. Trocar `bg-emerald-500` (D8) pelo token `--status-available`. Preservar a gaveta mobile da `main`.
- [ ] **Step 7:** Verificação e commit.

---

### Task 8: Fonte, CSV do backend e varredura final

**Files:** `frontend/tailwind.config.js`, `frontend/index.html`, `backend/app/routers/reports.py`

- [ ] **Step 1:** Trocar a fonte do sistema de Plus Jakarta Sans para **Inter**, no `tailwind.config.js` e no `index.html`. Autorizado pelo dono. Manter IBM Plex Mono como fonte mono.
- [ ] **Step 2:** Trazer o CSV melhorado do relatório mensal — BOM UTF-8, separador `;`, `QUOTE_ALL` e cabeçalhos em português. **Só `backend/app/routers/reports.py`**, nenhum outro arquivo do backend. Resolve o Excel em pt-BR abrir com acento quebrado e tudo numa coluna.
- [ ] **Step 3:** Varreduras: `limit: 500` e `FETCH_ALL_LIMIT` ausentes do código; `shadow-xs` ausente; opacidade sobre token ausente; `dark:` ausente; `bg-emerald-500` ausente.
- [ ] **Step 4:** `tsc --noUnusedLocals` para pegar o código morto que veio junto (D7) — o `tsconfig` tem a flag desligada, então o `tsc` normal não avisa.
- [ ] **Step 5:** `npx tsc --noEmit`, `npm run test:run`, `npx vite build`, e `docker compose build frontend`.
- [ ] **Step 6:** Confirmar que nada regrediu: paleta Ctrl+K funcionando, gaveta mobile, breadcrumbs, ordenação clicável, busca global do relatório, e `navigation.ts` como fonte única.

## Notas para quem executa

- A `FT_STC` é trabalho paralelo de outro desenvolvedor, com boas ideias e execução desigual. Onde ela for melhor, traga. Onde a `main` for melhor, mantenha a `main`. Onde ela estiver quebrada, corrija — nenhum dos oito defeitos listados entra como está.
- Ela passa nos próprios testes e compila: os defeitos são de runtime silencioso e de colisão, não de compilação. Não confie em "o build passou".
