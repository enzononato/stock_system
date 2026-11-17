# Fase 1 — Visual e listagens — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Aplicar a paleta, o `StatusPill`, a `Pagination` 25/50/100 e o `Button loading` da F0 nas telas que a F1 possui, dar ao Estoque colunas Setor/Cargo, filtros por unidade/tipo/status/colaborador e ordenação por coluna no servidor, mostrar unidade/revenda em CAIXA ALTA e fechar as lacunas de responsividade do layout.

**Architecture:** No backend, só a listagem de itens muda (`list_items` ganha `sort`/`dir` com lista branca e ordem SQL montada por uma função pura) e o CSV do relatório mensal passa a sair com revenda em maiúsculas. No frontend, o Estoque ganha um cliente próprio em `features/estoque/` (não toca em `api/items.ts`, que é da F4); as demais telas da F1 trocam selos/paginação locais pelos componentes da F0 seguindo uma receita visual única (seção "Decisões desta fase"). Uma última guarda estática impede paginação fixa e selos de status locais.

**Tech Stack:** FastAPI + PyMySQL (SQL cru) + MySQL 8; React 18 + TS 5.6 + Tailwind 3.4 + TanStack Query 5 + Vitest/MSW.

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md` — seção 6 "F1 — Visual e listagens", matriz da seção 3 (linhas 2.1, 2.2, 2.3, 2.4, 7.1, 7.3, 7.4, 7.5, 7.6/7.7, 7.8, 7.9, 7.10, 7.11, 8.6) e seção 7 (paleta). Interfaces da F0: `docs/superpowers/plans/2026-11-12-patrimonio-f0-base.md`.

## Global Constraints

- Paleta base `#0131CD` (azul), `#FFFFFF` (branco), `#B9FC01` (verde-limão); nenhuma cor fora dos tokens da F0 (`index.css`/`tailwind.config.js`). Modificador de opacidade em token (`bg-primary/10`) compila para nada — não usar.
- "Painéis grandes com fundo branco (claro) e preto (escuro); caixas menores (indicadores, chips, cabeçalhos de seção) em variações de azul."
- "Verde-limão `#B9FC01` só como preenchimento com texto escuro (`#111`) ou sobre o azul: destaque de item ativo na navegação, indicador 'novo', seleção, chips de destaque."
- Texto azul sempre `text-brand-text` (nunca `text-primary` — há teste-guarda).
- Status com cor própria e contraste AA nos dois temas: Disponível (verde), Em uso (azul), Pendente (âmbar), Pendente Devolução (roxo), Em Manutenção (laranja), Baixado (cinza), Erro/Recusado (vermelho). "Texto sempre em negrito." Sempre via `StatusPill`.
- Paginação: seletor 25/50/100, padrão 25, preferência por tela no `localStorage` (`usePageSize(tela)`); "nenhuma tela com paginação de 7 itens"; campo "ir para a página".
- "Unidade e revenda sempre em CAIXA ALTA na exibição e nas exportações (a gravação não muda, para não disparar a cascata de renomeação de `unidade_db.py`)." O valor enviado em filtros continua o nome original.
- Estoque: "colunas Setor e Cargo (do colaborador vinculado); filtros por unidade, tipo, status e colaborador; ordenação por coluna no servidor (`sort` com lista branca de colunas e `dir`); busca existente mantida."
- Responsividade: "navegação em gaveta abaixo de 768 px, tabelas viram cartões ou rolam horizontalmente sem quebrar a página, formulários em uma coluna; conferido em 390, 768 e 1440 px."
- Telas que só a F1 edita: `StockPage`, `PeripheralsPage`, `LinkPeripheralPage`, `HistoryPage`, `ReportPage`, `ChartsPage`, `components/equipment/ItemDetailsModal.tsx`, layout (`AppLayout`, `Sidebar`, `TopBar`, `CommandPalette`, `PageHeader`) e as correções da revisão da F0 em `Pagination`/`Button`/`StatusPill`. **Não editar** `LoanPage`, `ReturnPage`, `TermsPage`, `RegisterItemPage`, `RemovePage`, `UsersPage`, `LoginPage`, `DashboardPage`, `UnidadesPage`, `api/items.ts`, `components/ui/badge.tsx`, registros de abas.
- Backend: a F1 só mexe no handler `list_items` de `routers/items.py`, no método `list_items` de `inventory_manager_db.py` e no export de `routers/reports.py`.
- Arquivos compartilhados (`App.tsx`, `navigation.ts`, `test/handlers.ts`, `main.py`): só acréscimos — esta fase não precisa de nenhum.
- Migrações da F1: faixa `0100–0199` — nenhuma prevista.
- APIs públicas de `Pagination`, `StatusPill`, `Button`, `statusTone` inalteradas.
- Testes: backend `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh <wt> f1 [args]` (`<wt>` = raiz da worktree da F1); frontend em `<wt>/frontend`: `npx vitest run`, `npx tsc --noEmit`, `npm run build`. MSW com `onUnhandledRequest: 'error'`: todo teste mocka o que dispara. Todo arquivo de teste de tela que use `usePageSize` faz `localStorage.clear()` no `beforeEach`. "Requisição pendurada" = handler MSW que faz `await` numa promessa liberada pelo próprio teste no fim (`let liberar!: () => void; const pendente = new Promise<void>((r) => { liberar = r })`).
- Commits em português `tipo(escopo): descrição`, terminando com linha em branco + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus

- Rótulo em CAIXA ALTA mas filtro com o nome original: escolher "REVALLE JUAZEIRO" no filtro de unidade envia `revenda=Revalle Juazeiro` (Estoque, Relatório, Indicadores) — testes nas Tasks 6 e 9.
- Ordenação com itens sem colaborador, `assigned_to` vazio ou `NULL`: esses itens ficam no fim nas duas direções e a paginação não repete nem pula itens — testes na Task 1.
- `sort`/`dir` fora da lista branca (inclusive `id; DROP TABLE items`): 422 com mensagem em português, nenhum SQL montado com o texto recebido — testes na Task 1.
- Página atual além da última depois que o total encolhe (filtro, baixa feita por outra pessoa): a lista volta para a última página válida e o rodapé nunca mostra "126–120 de 120" — testes nas Tasks 3 e 6.
- Duplo clique/toque em ação assíncrona (salvar no modal, estornar, desvincular, sair): uma única requisição — testes nas Tasks 4, 7, 8 e 10.

## Decisões desta fase (o executor não decide de novo)

1. **Gaveta**: mantém o corte em `lg` (1024 px) que já existe — atende "abaixo de 768 px" e evita a sidebar fixa de 256 px espremendo tabelas em tablet. Lacunas a fechar: semântica de diálogo, botão "Fechar menu", busca acessível no toque, nome de usuário truncado, `min-w-0` no `<main>`.
2. **Tabelas**: rolagem horizontal (`<div className="overflow-x-auto">` + `table` com `min-w-[...]`), não cartões. O Estoque deixa de esconder colunas por breakpoint.
3. **Setor/Cargo do Estoque** = `colaborador_setor`/`colaborador_cargo`; sem colaborador mostra "—" (o `items.setor` é o setor do equipamento e não entra como reserva). Coluna "Colaborador" = `colaborador_nome ?? assigned_to ?? '—'`.
4. **Ordenação**: chaves `id, tipo, equipamento, status, colaborador, setor, cargo, unidade, cadastro`; padrão `sort=id&dir=desc` (= ordem atual). Vazio/`NULL` sempre no fim; desempate `i.id DESC`; `status` em ordem alfabética (`CAST(i.status AS CHAR)`); `colaborador` ordena por `i.assigned_to` (sempre gravado no empréstimo). Clique numa coluna nova → `asc`; na coluna ativa → inverte.
5. **Filtro de colaborador** inclui afastados e desligados (`ColaboradorSelect somenteAtivos={false}`): achar equipamento com desligado é o uso principal.
6. **Status "Em Manutenção"** já entra nas opções do filtro do Estoque (o valor é criado pela F4; a `StockPage` é da F1). Até a F4 integrar, o filtro só volta vazio.
7. **Linha do Estoque**: clique na linha abre a ficha rápida (modal, como hoje); `#id` é `Link` para `/items/:id`; ações visíveis no toque (sem depender de hover).
8. **Cliente do Estoque** em `frontend/src/features/estoque/api.ts` (novo) — `api/items.ts` é da F4.
9. **Botões com `loading`** mantêm o rótulo fixo (o spinner é do `Button`); nada de trocar o texto para "Salvando…".
10. **Receita visual** (vale para todas as tarefas de frontend; todos os pares texto/fundo abaixo são verificados por `lib/palette.test.ts`):

| Elemento | Classes |
|---|---|
| Painel grande (tabela, formulário, ficha, modal, rodapé de modal) | `rounded-md border border-border bg-surface` — nunca `bg-surface-alt` como fundo de painel |
| Caixa pequena / indicador | `rounded-md bg-brand-soft p-4`; rótulo `text-caption text-brand-soft-foreground`; número `text-heading font-bold text-brand-soft-foreground` (nada de `text-muted-foreground` dentro) |
| Chip | `rounded-full bg-brand-soft px-2 py-0.5 text-xs font-semibold text-brand-soft-foreground` |
| Chip de destaque ("Editando") | `rounded-full bg-lime px-2 py-0.5 text-xs font-semibold text-lime-foreground` |
| Cabeçalho de tabela (`thead`) | `bg-brand-soft text-brand-soft-foreground` |
| Cabeçalho de seção (eyebrow do `PageHeader`, título do `PanelHeader`, título de grupo da Sidebar) | `text-brand-text` |
| Item ativo na navegação | `bg-lime text-lime-foreground font-semibold` |
| Título de página | `PageHeader`; página com `space-y-6`; painéis `p-5`; botões de barra `size="sm"` |
| Status | `StatusPill` |
| Ação assíncrona | `<Button loading={mutation.isPending}>`; botões que disparariam a mesma ação ficam `disabled` enquanto pende |
| Unidade/revenda | `upper(v)` no texto; `value` de filtro = nome original |
| Tabela | `<div className="overflow-x-auto"><table className="w-full min-w-[NNNpx] text-left">` |
| Formulário | `grid grid-cols-1 sm:grid-cols-N`; selects de filtro `w-full sm:w-40` |

---

### Task 1: Ordenação no servidor (e filtro por colaborador) em `GET /api/items`

**Files:**
- Modify: `backend/app/db/inventory_manager_db.py` (constante `ORDENACOES_ITEMS`, função `montar_order_by_items`, método `list_items`)
- Modify: `backend/app/routers/items.py` (só o handler `list_items`)
- Create: `backend/tests/test_items_ordenacao.py`

**Interfaces:**
- Consumes (F0 Task 8): `list_items` já com `LEFT JOIN` em `colaboradores`/superior/`locais` produzindo `colaborador_nome`, `colaborador_setor`, `colaborador_cargo`, `superior_nome`, `local_nome`; tabela `colaboradores (cpf, nome, setor, cargo, situacao, origem, criado_em, atualizado_em, ...)`; `items.colaborador_id`.
- Produces:
  - `ORDENACOES_ITEMS: dict[str, tuple[str, ...]]` — `"id": ("i.id",)`, `"tipo": ("i.tipo",)`, `"equipamento": ("i.brand", "i.model")`, `"status": ("CAST(i.status AS CHAR)",)`, `"colaborador": ("i.assigned_to",)`, `"setor": (<mesma expressão SQL que o SELECT da F0 usa para `colaborador_setor`, ex. `c.setor`>,)`, `"cargo": (<idem para `colaborador_cargo`>,)`, `"unidade": ("i.revenda",)`, `"cadastro": ("i.date_registered",)`.
  - `montar_order_by_items(sort: str, direcao: str) -> str` — `ValueError("Ordenação inválida.")` se `sort` não é chave; `ValueError("Direção da ordenação inválida: use asc ou desc.")` se `direcao` não é `"asc"`/`"desc"`. Algoritmo: `D = direcao.upper()`; se `sort` não está em `{"id", "cadastro"}`, primeiro termo `NULLIF(<expr0>, '') IS NULL`; depois `<expr> D` para cada expressão; se `sort != "id"`, termo final `i.id DESC`; tudo unido por `", "`.
  - `InventoryDBManager.list_items(*, tipo=None, status=None, revenda=None, search=None, colaborador_id: int | None = None, sort: str = "id", direcao: str = "desc", limit: int | None = None, offset: int = 0) -> tuple[list[dict], int]` — `ORDER BY` vem de `montar_order_by_items` (nunca do texto do usuário).
  - `GET /api/items?sort=<chave>&dir=asc|desc` (padrões `id`/`desc`), além de `tipo`, `status`, `revenda`, `search`, `colaborador_id`, `limit`, `offset`. Inválido → 422 `{"detail": "Ordenação inválida."}` ou `{"detail": "Direção da ordenação inválida: use asc ou desc."}` (`HTTPException`, não `Literal`, para a mensagem sair em português). Parâmetro Python `direcao: str = Query("desc", alias="dir")`.

- [ ] **Step 1: Conferir a pré-condição da F0**

Run: `grep -n "colaborador_id\|colaborador_setor\|colaborador_cargo" backend/app/routers/items.py backend/app/db/inventory_manager_db.py`
Expected: `list_items` do manager seleciona `colaborador_setor`/`colaborador_cargo`. Se não selecionar, pare e reporte (F0 Task 8 incompleta). Se o filtro `colaborador_id` não existir no router e no manager, ele entra no Step 4 (`colaborador_id: Optional[int] = None` → `i.colaborador_id = %s`).

- [ ] **Step 2: Escrever os testes** em `backend/tests/test_items_ordenacao.py`

```python
import pytest
from app.db.inventory_manager_db import montar_order_by_items

def _ids(client, **params):
    resp = client.get("/api/items", params=params)
    assert resp.status_code == 200, resp.text
    return [i["id"] for i in resp.json()["items"]]

def _vincular(executar_sql, item_id, *, cpf, nome, setor, cargo):
    executar_sql(
        "INSERT INTO colaboradores (cpf, nome, setor, cargo, situacao, origem, criado_em, atualizado_em) "
        "VALUES (%s, %s, %s, %s, 'ativo', 'manual', NOW(), NOW())", (cpf, nome, setor, cargo))
    executar_sql("UPDATE items SET colaborador_id = (SELECT id FROM colaboradores WHERE cpf = %s) WHERE id = %s",
                 (cpf, item_id))

@pytest.mark.unit
@pytest.mark.parametrize("sort,direcao,esperado", [
    ("id", "desc", "i.id DESC"),
    ("id", "asc", "i.id ASC"),
    ("tipo", "asc", "NULLIF(i.tipo, '') IS NULL, i.tipo ASC, i.id DESC"),
    ("equipamento", "desc", "NULLIF(i.brand, '') IS NULL, i.brand DESC, i.model DESC, i.id DESC"),
    ("status", "asc", "NULLIF(CAST(i.status AS CHAR), '') IS NULL, CAST(i.status AS CHAR) ASC, i.id DESC"),
    ("cadastro", "desc", "i.date_registered DESC, i.id DESC"),
])
def test_montar_order_by_items(sort, direcao, esperado):
    assert montar_order_by_items(sort, direcao) == esperado

@pytest.mark.unit
@pytest.mark.parametrize("sort,direcao", [("id; DROP TABLE items", "asc"), ("TIPO", "asc"), ("tipo", "up")])
def test_montar_order_by_items_recusa_fora_da_lista(sort, direcao):
    with pytest.raises(ValueError):
        montar_order_by_items(sort, direcao)

@pytest.mark.integration
class TestOrdenacaoDeItems:
    def test_setor_do_colaborador_com_sem_colaborador_no_fim(self, client_gestor, criar_item, executar_sql):
        ti, com, sem = criar_item(), criar_item(), criar_item()
        _vincular(executar_sql, ti["id"], cpf="52998224725", nome="ANA", setor="TI", cargo="Analista")
        _vincular(executar_sql, com["id"], cpf="11144477735", nome="BRUNO", setor="Comercial", cargo="Vendedor")
        assert _ids(client_gestor, sort="setor", dir="asc") == [com["id"], ti["id"], sem["id"]]
        assert _ids(client_gestor, sort="setor", dir="desc") == [ti["id"], com["id"], sem["id"]]
        assert _ids(client_gestor, sort="cargo", dir="asc") == [ti["id"], com["id"], sem["id"]]

    def test_sort_fora_da_lista_branca_responde_422(self, client_gestor, item_disponivel):
        resp = client_gestor.get("/api/items", params={"sort": "id; DROP TABLE items"})
        assert resp.status_code == 422
        assert resp.json()["detail"] == "Ordenação inválida."
        assert client_gestor.get("/api/items").json()["total"] == 1
```

Demais testes da classe (mesmo estilo, `criar_item(...)` e `executar_sql("UPDATE items SET ... WHERE id=%s", ...)`):
- `test_padrao_continua_id_desc`: 3 itens, sem parâmetros → `_ids == sorted(ids, reverse=True)`.
- `test_tipo_asc_e_desc`: tipos `"Notebook"`, `"Desktop"`, `"Monitor"` → `asc` = `[desktop, monitor, notebook]`; `desc` = inverso.
- `test_status_em_ordem_alfabetica`: status gravados por SQL `'Pendente Devolução'`, `'Disponível'`, `'Pendente'`, `'Indisponível'` → `asc` = `[Disponível, Indisponível, Pendente, Pendente Devolução]`.
- `test_colaborador_usa_assigned_to_e_vazios_no_fim`: `assigned_to` = `"Bruno"`, `"Ana"`, `""` → `asc` = `[ana, bruno, vazio]`, `desc` = `[bruno, ana, vazio]`.
- `test_ordenacao_com_paginacao_nao_repete`: tipos `"C"`, `"A"`, `"B"`; `sort=tipo&dir=asc&limit=2&offset=0` → `[a, b]`; `offset=2` → `[c]`; `total == 3` nas duas.
- `test_filtro_colaborador_combina_com_ordenacao`: dois itens (tipos `"Notebook"`, `"Desktop"`) com o colaborador A e um com B → `colaborador_id=<A>&sort=tipo&dir=asc` → `[desktop_A, notebook_A]`, `total == 2`.
- `test_dir_invalido_responde_422`: `dir=up` → 422, `detail == "Direção da ordenação inválida: use asc ou desc."`.

- [ ] **Step 3: Rodar e ver falhar**

Run: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh <wt> f1 tests/test_items_ordenacao.py -q`
Expected: FAIL (`ImportError: cannot import name 'montar_order_by_items'`).

- [ ] **Step 4: Implementar** `ORDENACOES_ITEMS`, `montar_order_by_items` e os parâmetros novos de `list_items` (manager e router) conforme Interfaces. O `ORDER BY i.id DESC` fixo da query passa a ser `ORDER BY {montar_order_by_items(sort, direcao)}`; a contagem não muda. O router valida `sort`/`dir` antes de chamar o manager e repassa `colaborador_id`.

- [ ] **Step 5: Rodar a suíte inteira do backend**

Run: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh <wt> f1 -q -m "not mudanca_esperada"`
Expected: todos passam (incluindo `tests/test_items.py`, que usa a ordem padrão).

- [ ] **Step 6: Commit**

```bash
git add backend/app/db/inventory_manager_db.py backend/app/routers/items.py backend/tests/test_items_ordenacao.py
git commit -m "feat(estoque): ordenação por coluna no servidor em GET /api/items

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

### Task 2: Revenda em CAIXA ALTA no CSV do relatório mensal

**Files:**
- Modify: `backend/app/routers/reports.py` (`export_monthly_report`)
- Test: `backend/tests/test_monthly_report.py`

**Interfaces:**
- Produces: `GET /api/reports/monthly/export` grava a coluna `revenda` com `str(v).upper()`; `GET /api/reports/monthly` (JSON) continua com o valor gravado.

- [ ] **Step 1: Escrever o teste** na classe `TestRelatorioMensal`:

```python
def test_export_csv_revenda_em_caixa_alta(self, client_gestor, criar_item):
    criar_item(revenda="Revalle Juazeiro")
    agora = datetime.now()
    params = {"year": agora.year, "month": agora.month}
    csv_txt = client_gestor.get("/api/reports/monthly/export", params=params).text
    assert '"REVALLE JUAZEIRO"' in csv_txt
    assert "Revalle Juazeiro" not in csv_txt
    linhas = client_gestor.get("/api/reports/monthly", params=params).json()
    assert any(r["revenda"] == "Revalle Juazeiro" for r in linhas)
```

- [ ] **Step 2: Rodar e ver falhar:** `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh <wt> f1 tests/test_monthly_report.py -q` → FAIL no `'"REVALLE JUAZEIRO"'`.
- [ ] **Step 3: Implementar:** no laço do `writer.writerow`, o valor da chave `revenda` (quando não `None`) sai em `.upper()`; demais colunas inalteradas.
- [ ] **Step 4: Rodar** o mesmo arquivo → PASS.
- [ ] **Step 5: Commit** `feat(relatorio): revenda em caixa alta no CSV do relatório mensal` (com a linha `Co-Authored-By`).

### Task 3: Correções da revisão da F0 nos componentes compartilhados

**Files:**
- Modify: `frontend/src/components/ui/Pagination.tsx`, `frontend/src/components/ui/StatusPill.tsx`, `frontend/src/components/ui/button.tsx`, `frontend/src/lib/status.ts` (só comentário)
- Test: `frontend/src/components/ui/Pagination.test.tsx`, `frontend/src/components/ui/StatusPill.test.tsx`, `frontend/src/components/ui/button.test.tsx`, `frontend/src/lib/status.test.ts`

**Interfaces:**
- Produces (props inalteradas):
  - `Pagination`: `const atual = Math.min(Math.max(1, page), lastPage)` usado no texto "Mostrando X–Y de Z", em `{atual} / {lastPage}`, nos `disabled` e em Anterior/Próxima (`onPageChange(atual - 1)` / `atual + 1`). "Ir para" só aceita `/^\d+$/` depois de `trim()`; o resto não chama `onPageChange`.
  - `StatusPill`: vazio = `!status?.trim()` → "—" neutro.
  - `Button`: desestrutura `'aria-busy': ariaBusy` e aplica `aria-busy={loading ? true : ariaBusy}` **depois** de `{...props}`.
  - `statusTone`: comentário JSDoc "casamento exato (sem acento, caixa e espaços nas pontas); 'Pendente …' desconhecido não herda o tom de 'Pendente' e fica neutro — status novo entra na tabela `TONE_BY_STATUS`".

- [ ] **Step 1: Escrever os testes**

```tsx
// Pagination.test.tsx
it('página além da última mostra a última (nunca "126–120")', () => {
  const { onPageChange } = setup({ page: 6 })
  expect(screen.getByText('Mostrando 101–120 de 120')).toBeInTheDocument()
  expect(screen.getByText('5 / 5')).toBeInTheDocument()
  expect(screen.getByRole('button', { name: 'Próxima' })).toBeDisabled()
  fireEvent.click(screen.getByRole('button', { name: 'Anterior' }))
  expect(onPageChange).toHaveBeenCalledWith(4)
})
it.each(['1e2', '2.5', '-1', '0x3'])('"Ir para" com %s não chama onPageChange', (valor) => {
  const { onPageChange } = setup()
  fireEvent.change(screen.getByLabelText('Ir para a página'), { target: { value: valor } })
  fireEvent.click(screen.getByRole('button', { name: 'Ir' }))
  expect(onPageChange).not.toHaveBeenCalled()
})
// StatusPill.test.tsx
it('status só com espaços mostra travessão neutro', () => {
  render(<StatusPill status="   " />)
  expect(screen.getByText('—').style.color).toBe('var(--status-neutral)')
})
// button.test.tsx
it('aria-busy do consumidor não sobrescreve o loading', () => {
  render(<Button loading aria-busy={false}>Salvar</Button>)
  expect(screen.getByRole('button', { name: 'Salvar' })).toHaveAttribute('aria-busy', 'true')
})
it('asChild + loading: o link não chama onClick e a navegação é cancelada', () => {
  const onClick = vi.fn()
  render(<Button asChild loading onClick={onClick}><a href="/x">Ir</a></Button>)
  const link = screen.getByRole('link', { name: 'Ir' })
  expect(fireEvent.click(link)).toBe(false)
  expect(onClick).not.toHaveBeenCalled()
  expect(link).toHaveAttribute('aria-busy', 'true')
})
// status.test.ts
it('"Pendente …" desconhecido fica neutro (casamento exato)', () => {
  expect(statusTone('Pendente Assinatura')).toBe('neutral')
  expect(statusTone('Pendente Devolução')).toBe('return')
})
```

- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/components/ui src/lib/status.test.ts` → FAIL em "126–120", "1e2", espaços e `aria-busy` (o teste de `asChild` e o de `statusTone` são de regressão e podem já passar).
- [ ] **Step 3: Implementar** conforme Interfaces.
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes (os testes antigos de `Pagination` — "Ir para 0 → 1", "999 → 5" — continuam passando).
- [ ] **Step 5: Commit** `fix(ui): paginação além da última página, "ir para" só com dígitos, status vazio e aria-busy do botão`.

### Task 4: Layout — navegação, topo, gaveta e cabeçalhos

**Files:**
- Modify: `frontend/src/components/layout/Sidebar.tsx`, `TopBar.tsx`, `AppLayout.tsx`, `CommandPalette.tsx`, `PageHeader.tsx`, `frontend/src/lib/palette.test.ts`
- Create: `frontend/src/components/layout/AppLayout.test.tsx`
- Test: `frontend/src/components/layout/CommandPalette.test.tsx`

**Interfaces:**
- Produces:
  - `TopBarProps { onOpenMobileNav: () => void; onOpenCommandPalette: () => void }` — botão ghost com ícone `Search`, `aria-label="Buscar tela (Ctrl+K)"`; nome de usuário `truncate max-w-[40vw]`; papel `hidden sm:inline`; "Sair do sistema" com `loading` enquanto `logout()` pende (estado local `saindo`, `finally` volta a `false`).
  - `AppLayout`: gaveta com `role="dialog" aria-modal="true" aria-label="Menu de navegação"` e botão `aria-label="Fechar menu"` (ícone `X`, canto superior direito da gaveta); `<main>` ganha `min-w-0`; passa `onOpenCommandPalette={() => setCommandPaletteOpen(true)}`.
  - `Sidebar`: `NavLink` ativo `bg-lime text-lime-foreground font-semibold` (sai `border-l-2 border-foreground`); título de grupo `text-brand-text`; selo "Controle de Patrimônio" vira chip `bg-brand-soft text-brand-soft-foreground` sem borda.
  - `CommandPalette`: opção ativa `bg-brand-soft text-brand-soft-foreground`.
  - `PageHeader`: eyebrow `text-brand-text`; `PanelHeader`: título `text-brand-text`. Props inalteradas.
  - `palette.test.ts`: acrescenta o par `['--brand-text', '--canvas']` em `pares` (o eyebrow azul fica sobre o fundo da página).

- [ ] **Step 1: Escrever os testes.** `AppLayout.test.tsx`: `vi.mock('@/contexts/AuthContext')` com `user: { id: 1, username: 'gestor.teste', role: 'Gestor' }` e `logout` = `vi.fn(() => pendente)` (promessa liberada no fim do teste); render em `<ThemeProvider><MemoryRouter initialEntries={['/stock']}><Routes><Route element={<AppLayout />}><Route path="/stock" element={<p>conteúdo</p>} /></Route></Routes></MemoryRouter></ThemeProvider>`.
  - `it('item ativo da navegação usa o destaque verde-limão')`: `expect(screen.getAllByRole('link', { name: 'Estoque' })[0]).toHaveClass('bg-lime', 'text-lime-foreground')`.
  - `it('gaveta abre como diálogo e fecha pelo botão')`: clicar "Abrir menu de navegação" → `getByRole('dialog', { name: 'Menu de navegação' })`; clicar "Fechar menu" → `queryByRole('dialog', { name: 'Menu de navegação' })` ausente.
  - `it('botão de busca abre a paleta de comandos')`: clicar "Buscar tela (Ctrl+K)" → `getByRole('dialog', { name: 'Paleta de comandos' })`.
  - `it('sair com duplo clique chama logout uma vez')`: dois cliques em "Sair do sistema" → `logout` chamado 1×; botão com `aria-busy="true"`.
  - `CommandPalette.test.tsx`: `it('opção ativa usa variação de azul')` → primeira `option` tem `bg-brand-soft`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/components/layout src/lib/palette.test.ts` → FAIL (classes e botões novos ausentes; o par novo da paleta deve passar).
- [ ] **Step 3: Implementar** conforme Interfaces.
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes (`PageHeader.test.tsx` não depende de cor).
- [ ] **Step 5: Commit** `feat(layout): destaque verde-limão no menu, busca no topo e gaveta acessível`.

### Task 5: Módulo do Estoque — cliente da listagem e cabeçalho ordenável

**Files:**
- Create: `frontend/src/features/estoque/api.ts`, `frontend/src/features/estoque/api.test.ts`, `frontend/src/features/estoque/SortableHeader.tsx`, `frontend/src/features/estoque/SortableHeader.test.tsx`

**Interfaces:**
- Consumes: `GET /api/items` com `sort`/`dir`/`colaborador_id` (Task 1); `api` e `Paginated` de `@/api/client`; `Item` de `@/api/items`.
- Produces:
  - `export const ESTOQUE_SORTS = ['id', 'tipo', 'equipamento', 'status', 'colaborador', 'setor', 'cargo', 'unidade', 'cadastro'] as const`; `export type EstoqueSort = (typeof ESTOQUE_SORTS)[number]`; `export type SortDir = 'asc' | 'desc'`.
  - `export interface EstoqueParams { tipo?: string; status?: string; revenda?: string; search?: string; colaborador_id?: number; sort: EstoqueSort; dir: SortDir; limit: number; offset: number }`
  - `export async function listEstoque(params: EstoqueParams): Promise<Paginated<Item>>` (`api.get('/items', { params })`).
  - `export function proximaOrdenacao(atual: { sort: EstoqueSort; dir: SortDir }, coluna: EstoqueSort): { sort: EstoqueSort; dir: SortDir }` — coluna nova → `asc`; mesma coluna → inverte.
  - `export function SortableHeader({ label, coluna, sort, dir, onSort, className }: { label: string; coluna: EstoqueSort; sort: EstoqueSort; dir: SortDir; onSort: (coluna: EstoqueSort) => void; className?: string })` — `<th scope="col" aria-sort={ativa ? (dir === 'asc' ? 'ascending' : 'descending') : 'none'}>` com `<button type="button">{label}<ícone aria-hidden /></button>`; ícones lucide `ArrowUp`/`ArrowDown` (ativa) e `ArrowUpDown` (inativa).

- [ ] **Step 1: Escrever os testes.**
  - `api.test.ts`: MSW captura `/api/items`; `listEstoque({ colaborador_id: 7, sort: 'setor', dir: 'asc', limit: 25, offset: 50 })` → `params.get('sort') === 'setor'`, `'dir' === 'asc'`, `'colaborador_id' === '7'`, `'limit' === '25'`, `'offset' === '50'`. `proximaOrdenacao({ sort: 'id', dir: 'desc' }, 'tipo')` → `{ sort: 'tipo', dir: 'asc' }`; `({ sort: 'tipo', dir: 'asc' }, 'tipo')` → `{ sort: 'tipo', dir: 'desc' }`; `({ sort: 'tipo', dir: 'desc' }, 'tipo')` → `{ sort: 'tipo', dir: 'asc' }`.
  - `SortableHeader.test.tsx` (renderizado dentro de `<table><thead><tr>`): coluna ativa `asc` → `columnheader` com `aria-sort="ascending"`; inativa → `"none"`; clicar o botão "Tipo" chama `onSort('tipo')`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/estoque` → FAIL (módulo inexistente).
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run src/features/estoque` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(estoque): cliente da listagem com ordenação e cabeçalho ordenável`.

### Task 6: Tela de Estoque

**Files:**
- Modify: `frontend/src/pages/StockPage.tsx`
- Create: `frontend/src/pages/StockPage.test.tsx`

**Interfaces:**
- Consumes: `listEstoque`, `proximaOrdenacao`, `SortableHeader`, `EstoqueSort`, `SortDir` (Task 5); `Pagination` (Task 3); `usePageSize`, `useDebouncedValue`, `StatusPill`, `upper`, `ColaboradorSelect`, `PageHeader`, `ItemDetailsModal`, `listUnidades`, `useConstants`.
- Produces (comportamento):
  - Estado: `page` (1-based), `[pageSize, setPageSize] = usePageSize('estoque')`, `{ sort, dir }` inicial `{ 'id', 'desc' }`, filtros `tipo`, `status`, `revenda` (nome original), `colaboradorId: number | null`, `search` com `useDebouncedValue(search.trim(), 400)`. Toda mudança de filtro, busca, ordenação ou tamanho → `page = 1`. `useEffect`: com `total > 0` e `page > lastPage` → `setPage(lastPage)`.
  - Query: `queryKey ['items', 'estoque', { tipo, status, revenda, colaboradorId, busca, sort, dir, page, pageSize }]`, `listEstoque({ ..., limit: pageSize, offset: (page - 1) * pageSize })`, `placeholderData: keepPreviousData`.
  - `STATUS_OPTIONS = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução', 'Em Manutenção']`.
  - Filtros: Radix `Select` com `SelectTrigger aria-label` "Tipo", "Status", "Unidade" (itens de unidade: `value={u.nome}`, texto `upper(u.nome)`); `ColaboradorSelect value={colaboradorId} onChange={(c) => ...(c?.id ?? null)} somenteAtivos={false} placeholder="Todos os colaboradores"`; "Limpar" zera filtros e busca (não a ordenação).
  - Colunas (`SortableHeader`, nesta ordem e com estes rótulos): Patrimônio (`id`), Tipo (`tipo`), Equipamento (`equipamento`), Status (`status`), Colaborador (`colaborador`), Setor (`setor`), Cargo (`cargo`), Unidade (`unidade`), Cadastro (`cadastro`), e a coluna de ações sem rótulo. Células: `#id` em `<Link to={`/items/${id}`}>` (com `stopPropagation`) + identificador; `StatusPill`; colaborador/setor/cargo pela Decisão 3; `upper(revenda) || '—'`.
  - Tabela pela receita (`min-w-[960px]`, `thead` azul); ações `opacity-100 lg:opacity-0 lg:group-hover:opacity-100 lg:focus-within:opacity-100`; `Pagination` sempre visível abaixo da tabela com `onPageSizeChange` e `isFetching`; cabeçalho `PageHeader` (eyebrow "Inventário", título "Estoque de Equipamentos", descrição `${total} ativos registrados`, ações "Atualizar" com `loading={isFetching}` e "Novo Ativo" para Gestor/Técnico).

- [ ] **Step 1: Escrever os testes** em `StockPage.test.tsx`: `vi.mock('@/contexts/AuthContext')` (Gestor); stubs de Pointer Events/`scrollIntoView` em `beforeAll` (mesmo bloco de `RemovePage.test.tsx`); `beforeEach(() => localStorage.clear())`; render em `<MemoryRouter>` com `renderWithClient`; MSW para `/api/items` (guarda cada `URLSearchParams` em `pedidos`), `/api/unidades` (`[{ id: 1, nome: 'Revalle Juazeiro', ... }]`), `/api/colaboradores` (`busca=mar` → `[maria]` com `id: 7, nome: 'MARIA SILVA'`). Itens base: `{ id: 10, tipo: 'Notebook', status: 'Indisponível', revenda: 'Revalle Juazeiro', colaborador_nome: 'MARIA SILVA', colaborador_setor: 'Comercial', colaborador_cargo: 'Vendedora' }` e `{ id: 11, tipo: 'Desktop', status: 'Disponível', revenda: 'Revalle Juazeiro', colaborador_nome: null, assigned_to: null }`.
  - `it('mostra Setor e Cargo do colaborador, unidade em CAIXA ALTA e "—" sem colaborador')`: na linha do #10, `'Comercial'`, `'Vendedora'`, `'REVALLE JUAZEIRO'` e `getByText('Indisponível')` com `font-bold`; na linha do #11, Setor e Cargo `'—'`.
  - `it('pede 25 por página e lembra a troca para 50')`: primeiro pedido `limit=25`, `offset=0`, `sort=id`, `dir=desc`; escolher 50 em "Itens por página" → pedido `limit=50`, `offset=0`; `localStorage.getItem('revalle_page_size:estoque') === '50'`.
  - `it('ordenação por coluna vai para o servidor e volta à página 1')`: com `total: 60`, ir para a página 2 ("Próxima"); clicar o botão "Tipo" → último pedido `sort=tipo`, `dir=asc`, `offset=0`; clicar de novo → `dir=desc`; `getByRole('columnheader', { name: /Tipo/ })` com `aria-sort="descending"`.
  - `it('filtro de unidade mostra CAIXA ALTA mas envia o nome original')`: abrir `combobox` "Unidade", escolher `option` "REVALLE JUAZEIRO" → último pedido `revenda === 'Revalle Juazeiro'`, `offset === '0'`.
  - `it('filtro de colaborador envia colaborador_id')`: abrir o seletor "Todos os colaboradores", digitar `mar`, escolher "MARIA SILVA" → último pedido `colaborador_id === '7'`, `offset === '0'`; sem `situacao` no pedido de `/api/colaboradores`.
  - `it('página além da última depois de o total encolher volta para a última')`: handler gera `total` de uma variável (`60` no início, itens `1000+i` conforme `offset`/`limit`); trocar a variável para `30`; "Ir para a página" `3` + Enter → resposta `{ items: [], total: 30 }` → aguardar pedido com `offset=25` e `findByText('Mostrando 26–30 de 30')`.
  - `it('#id leva à página do equipamento e a tabela rola dentro do painel')`: `getByRole('link', { name: /#10/ })` com `href="/items/10"`; `getByRole('table').parentElement` com `overflow-x-auto`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/StockPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar** conforme Interfaces (remover `PAGE_SIZE = 7`, `MonoBadge` e a paginação manual).
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(estoque): setor, cargo, filtros, ordenação e paginação 25/50/100 no estoque`.

### Task 7: Periféricos (`PeripheralsPage` e `LinkPeripheralPage`)

**Files:**
- Modify: `frontend/src/pages/PeripheralsPage.tsx`, `frontend/src/pages/LinkPeripheralPage.tsx`
- Create: `frontend/src/pages/PeripheralsPage.test.tsx`, `frontend/src/pages/LinkPeripheralPage.test.tsx`

**Interfaces:**
- Consumes: `Pagination`, `usePageSize`, `StatusPill`, `Button loading`, `upper`, `PageHeader`.
- Produces (comportamento):
  - `PeripheralsPage`: sai o `StatusBadge` local e `TABLE_PAGE_SIZE`; paginação no cliente (o endpoint não pagina) com `usePageSize('perifericos')`, `page` 1-based, recuo para a última página válida; `StatusPill` na tabela e na ficha; tabela pela receita (`min-w-[640px]`); `SelectTrigger` "Tipo"/"Status" com `aria-label` e `w-full sm:w-36`; botões com `loading`: "Cadastrar Periférico" (`createMutation`), "Confirmar Vínculo" (`linkMutation`), "Desvincular do Ativo" (`unlinkMutation`), "Inativar Periférico" (`deleteMutation`, continua `disabled` quando `status === 'Em Uso'`); subtítulo do seletor de equipamento `upper(i.revenda)`; cabeçalho `PageHeader`.
  - `LinkPeripheralPage`: cabeçalho `PageHeader` (eyebrow "Ativos de Suporte", detalhe "Vínculo de Periféricos", título "Vincular Periféricos"); `PeripheralCard` troca `Badge` por `StatusPill` e ganha props `loading?: boolean; disabled?: boolean` repassadas ao botão; "Desvincular" com `loading={unlinkMutation.isPending && unlinkMutation.variables === p.link_id}` e "Vincular" com `loading={linkMutation.isPending && linkMutation.variables === p.id}`, todos `disabled` enquanto qualquer uma das duas pende; linha do equipamento escolhido com `StatusPill` e `Unidade: upper(revenda)`; subtítulos do seletor `[upper(i.revenda), i.identificador]`; painel de substituição `bg-surface` (não `bg-surface-alt`) e "Confirmar Substituição" com `loading={replaceMutation.isPending}`.

- [ ] **Step 1: Escrever os testes.**
  - `PeripheralsPage.test.tsx` (Gestor; MSW `/api/peripherals` com 30 periféricos, o de id 1 `status: 'Em Uso', link_id: 33`; `/api/items` padrão): `it('pagina no cliente com 25 por padrão')` → `'Mostrando 1–25 de 30'` e, depois de "Próxima", `'Mostrando 26–30 de 30'`; `it('status em pílula')` → `getAllByText('Em Uso')[0]` com `font-bold` e `style.color === 'var(--status-loaned)'`; `it('desvincular com duplo clique envia uma requisição só')` → clicar a linha do #1, clicar "Desvincular do Ativo" duas vezes com o `DELETE /api/peripherals/links/33` pendurado numa promessa → 1 chamada e botão com `aria-busy="true"`.
  - `LinkPeripheralPage.test.tsx` (MSW: `/api/items` devolve `{ id: 5, tipo: 'Desktop', brand: 'Dell', model: 'OptiPlex', status: 'Indisponível', revenda: 'Revalle Juazeiro' }` só para `tipo=Desktop`; `/api/items/5/peripherals` → `[{ id: 9, tipo: 'Mouse', status: 'Em Uso', link_id: 33 }]`; `/api/peripherals` → `[]`): escolher o equipamento no seletor ("Selecione ou busque um equipamento…" → opção `#5 — Desktop Dell OptiPlex`) → `getByText('REVALLE JUAZEIRO')`, `getByText('Em Uso')` com `font-bold`; clicar "Desvincular" duas vezes com o `DELETE` pendurado → 1 chamada.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/PeripheralsPage.test.tsx src/pages/LinkPeripheralPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar** conforme Interfaces.
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(perifericos): StatusPill, paginação 25/50/100 e botões com carregamento`.

### Task 8: Histórico

**Files:**
- Modify: `frontend/src/pages/HistoryPage.tsx`
- Test: `frontend/src/pages/HistoryPage.test.tsx`

**Interfaces:**
- Consumes: `listHistoryPaginated`, `Pagination`, `usePageSize`, `upper`, `Button loading`, `PageHeader`.
- Produces (comportamento): sai `PAGE_SIZE = 7` e o bloco lateral "Paginação"; `page` 1-based com `usePageSize('historico')`, `limit: pageSize`, `offset: (page - 1) * pageSize`; `Pagination` (com `onPageSizeChange` e `isFetching`) no fim do painel da linha do tempo; nova busca ou novo tamanho → página 1; recuo para a última página válida quando o total encolhe. "Sumário — Página Atual" vira grade `grid-cols-2 gap-2` de caixas pequenas da receita. Revenda nos detalhes `upper(entry.revenda)`; CSV exporta `history.map((h) => ({ ...h, revenda: upper(h.revenda) }))`. "Confirmar" do estorno: `loading={reverseMutation.isPending}`, `disabled={!password}`, rótulo fixo "Confirmar". Botões "Comprovante"/"Termo": `loading={downloadingKey === <chave>}`, rótulo fixo. Linha-resumo de cada evento com `flex-wrap`. Cabeçalho `PageHeader` (eyebrow "Auditoria e Rastreabilidade", detalhe "Linha Temporal", ação "Exportar CSV").

- [ ] **Step 1: Escrever os testes** (acrescentar ao arquivo, mesmo `vi.mock` de Gestor):
  - `it('pede 25 por página e mostra a paginação padrão')`: primeiro pedido `limit=25`, `offset=0`; `findByText('Mostrando 1–1 de 1')`.
  - `it('revenda em CAIXA ALTA nos detalhes')`: entrada com `revenda: 'Revalle Juazeiro'`; clicar "Expandir detalhes" → `getByText('REVALLE JUAZEIRO')`.
  - `it('confirmar estorno com duplo clique envia uma requisição só')`: `POST /api/history/10/reverse` pendurado numa promessa; digitar a senha, clicar "Confirmar" duas vezes → `chamadasReverse === 1` e botão com `aria-busy="true"`.
  - Os 4 testes existentes continuam passando sem mudança.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/HistoryPage.test.tsx` → FAIL nos novos.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(historico): paginação 25/50/100, revenda em caixa alta e estorno com clique único`.

### Task 9: Relatório mensal e Indicadores

**Files:**
- Modify: `frontend/src/pages/ReportPage.tsx`, `frontend/src/pages/ChartsPage.tsx`
- Create: `frontend/src/pages/ReportPage.test.tsx`, `frontend/src/pages/ChartsPage.test.tsx`

**Interfaces:**
- Consumes: `Pagination`, `usePageSize`, `upper`, `Button loading`, `PageHeader`.
- Produces (comportamento):
  - `ReportPage`: paginação no cliente com `usePageSize('relatorio')` (`page` 1-based; página efetiva `Math.min(page, lastPage)`); sai `PAGE_SIZE = 7`, o `Badge` "Pág X/Y" e o import de `badge`; filtro "Revenda / Filial" com `SelectTrigger aria-label="Revenda"`, itens `value={r}` e texto `upper(r)`; coluna Revenda `upper(row.revenda) || '—'`; resumo `• ${upper(filterRevenda)}`; quatro indicadores como caixas pequenas da receita; "Gerar" `loading={isFetching}`; "CSV" `loading={isExporting}` com rótulo fixo; título em `PageHeader` (eyebrow "Relatórios Corporativos", título "Relatório Mensal de Operações") e filtros num painel próprio, controles `w-full sm:w-44`.
  - `ChartsPage`: KPIs como caixas pequenas da receita; `CORES_GRAFICO.serie1` lê `--brand-text` (fallback `#0131CD`) e a legenda de "Empréstimos" usa `bg-brand-text`; seletor de unidade com `SelectTrigger aria-label="Unidade"`, itens `value={u.nome}` e texto `upper(u.nome)`; `periodoLegenda` com `upper(params.unidade)`; "Aplicar Filtro" `loading={loansQuery.isFetching || regQuery.isFetching || monthlyQuery.isFetching}`; controles `w-full sm:w-52`.

- [ ] **Step 1: Escrever os testes.**
  - `ReportPage.test.tsx` (stubs de Pointer Events; MSW `/api/unidades` → `[]`, `/api/constants` padrão, `/api/reports/monthly` → 30 linhas: 20 com `revenda: 'Revalle Juazeiro'`, 10 com `'Revalle Petrolina'`, `operation_type: 'Cadastro'`): `it('pagina no cliente com 25 por padrão')` → `'Mostrando 1–25 de 30'`; `it('revenda em CAIXA ALTA e filtro pelo nome original')` → célula `'REVALLE JUAZEIRO'`; escolher a opção "REVALLE PETROLINA" no `combobox` "Revenda" → `'Mostrando 1–10 de 10'`; `it('indicadores em caixas azuis')` → `getByText('Total de Registros').closest('div')` com `bg-brand-soft`.
  - `ChartsPage.test.tsx` (`<ThemeProvider>`; `vi.mock` de `useAuth` como Gestor; MSW `/api/reports/charts/loans` e `/registrations` com `days: [1..30]` e valores zerados — os gráficos ficam no estado vazio e o recharts não renderiza; `/api/unidades` → `[{ id: 1, nome: 'Revalle Juazeiro', ... }]`; stubs de Pointer Events): abrir o `combobox` "Unidade" → `getByRole('option', { name: 'REVALLE JUAZEIRO' })`; `getByText('Empréstimos no mês').closest('div')` com `bg-brand-soft`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/ReportPage.test.tsx src/pages/ChartsPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(relatorios): paginação, revenda em caixa alta e indicadores em azul`.

### Task 10: Ficha rápida do equipamento (`ItemDetailsModal`)

**Files:**
- Modify: `frontend/src/components/equipment/ItemDetailsModal.tsx`
- Test: `frontend/src/components/equipment/ItemDetailsModal.test.tsx`

**Interfaces:**
- Consumes: campos `colaborador_nome`, `colaborador_setor`, `colaborador_cargo`, `superior_nome`, `local_nome`, `colaborador_id` de `Item` (F0); `StatusPill`, `upper`, `Button loading`.
- Produces (comportamento): `StatusBadge` → `StatusPill` (import de `@/components/ui/StatusPill`); `ContextCard` vira caixa pequena da receita, com Unidade `upper(item.revenda) || 'Não definida'` e Usuário Atual `item.colaborador_nome ?? item.assigned_to ?? 'Nenhum'`; chip `#id` pela receita e chip "Editando" em verde-limão. Seção "Dados de Alocação": "Funcionário Alocado" = `colaborador_nome ?? assigned_to`, e entram "Setor do colaborador" (`colaborador_setor`), "Cargo" (`colaborador_cargo`), "Superior imediato" (`superior_nome || (temColaborador ? 'Não informado' : null)`, com `temColaborador = item.colaborador_id != null || !!item.colaborador_nome`, mesma regra da aba Dados da F0) e "Local" (`local_nome`). Grade das seções `grid-cols-1 sm:grid-cols-2 lg:grid-cols-3` (`col-span-2` → `sm:col-span-2`; `col-span-2 sm:col-span-3` → `sm:col-span-2 lg:col-span-3`). "Salvar Alterações" com `loading={saveMutation.isPending}` e ícone `Save` fixo (sai o `Loader2` manual); "Ver Termo Assinado" com `loading={downloadingTerm}` e rótulo fixo; rodapé `flex flex-wrap` sobre `bg-surface`; botão de fechar com `aria-label="Fechar"`.

- [ ] **Step 1: Escrever os testes** (acrescentar):
  - `it('mostra colaborador, setor, cargo, superior, local e unidade em CAIXA ALTA')`: item com `revenda: 'Revalle Juazeiro', colaborador_id: 3, colaborador_nome: 'MARIA SILVA', colaborador_setor: 'Comercial', colaborador_cargo: 'Vendedora', superior_nome: 'JOÃO GESTOR', local_nome: 'Almoxarifado'` → textos `'Superior imediato'`, `'JOÃO GESTOR'`, `'Comercial'`, `'Vendedora'`, `'Almoxarifado'`, `'REVALLE JUAZEIRO'`.
  - `it('superior "Não informado" quando há colaborador sem superior')`: mesmo item com `superior_nome: null` → `'Não informado'`.
  - `it('status em pílula')`: `getByText('Indisponível')` com `font-bold`.
  - `it('salvar com duplo clique envia um PUT só')`: `PUT /api/items/:id` pendurado numa promessa; "Editar" → dois cliques em "Salvar Alterações" → 1 chamada e botão com `aria-busy="true"`.
  - Os testes existentes continuam passando sem mudança.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/components/equipment` → FAIL nos novos.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(equipamento): colaborador, superior e local na ficha rápida e salvar com clique único`.

### Task 11: Guarda contra paginação fixa e selos de status locais

**Files:**
- Create: `frontend/src/test/listagens.guard.test.ts`

**Interfaces:**
- Produces: `AGUARDANDO_OUTRAS_FASES: Record<string, string>` = `{ 'pages/LoanPage.tsx': 'F3', 'pages/ReturnPage.tsx': 'F3', 'pages/TermsPage.tsx': 'F3' }` (caminhos relativos a `src/`).

- [ ] **Step 1: Escrever o teste** (lê os arquivos com `node:fs`, como `lib/tokens.test.ts`; percorre `src/pages`, `src/components` e `src/features` recursivamente, ignorando `*.test.ts(x)`):
  - `it('nenhuma tela usa tamanho de página fixo')`: nenhum arquivo fora de `AGUARDANDO_OUTRAS_FASES` casa `/\b[A-Z_]*PAGE_SIZE\s*=\s*\d+/`.
  - `it('a lista de espera não tem entradas vencidas')`: cada arquivo de `AGUARDANDO_OUTRAS_FASES` ainda casa o padrão; senão falha com `` `${arquivo} já não usa tamanho fixo: remova-o de AGUARDANDO_OUTRAS_FASES` ``.
  - `it('telas da F1 não têm selo de status próprio')`: `pages/StockPage.tsx`, `pages/PeripheralsPage.tsx`, `pages/LinkPeripheralPage.tsx`, `pages/HistoryPage.tsx`, `pages/ReportPage.tsx`, `pages/ChartsPage.tsx`, `components/equipment/ItemDetailsModal.tsx` não casam `/function (MonoBadge|StatusBadge)\b/` nem `/from '@\/components\/ui\/badge'/`.
  - `it('listagens da F1 usam Pagination e usePageSize')`: `StockPage`, `PeripheralsPage`, `HistoryPage`, `ReportPage` contêm `from '@/components/ui/Pagination'` e `usePageSize(`.
- [ ] **Step 2: Rodar:** `npx vitest run src/test/listagens.guard.test.ts` → PASS (as Tasks 6–10 já trocaram as telas; se falhar, corrigir a tela apontada, não o teste).
- [ ] **Step 3: Commit** `test(visual): guarda contra paginação fixa e selos de status locais`.

### Task 12: Verificação final e checklist da fase

**Files:** nenhum novo (só correções que as suítes apontarem, cada uma com seu commit `fix(...)`).

- [ ] **Step 1: Backend completo:** `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh <wt> f1 -q -m "not mudanca_esperada"` → todos passam.
- [ ] **Step 2: Frontend completo** em `<wt>/frontend`: `npx vitest run`, `npx tsc --noEmit`, `npm run build` → verdes, sem erro de tipo.
- [ ] **Step 3: Varredura:** `grep -rnE "PAGE_SIZE\s*=\s*[0-9]" frontend/src/pages frontend/src/components` → só `LoanPage`, `ReturnPage`, `TermsPage` (F3); `grep -rn "MonoBadge\|bg-primary/" frontend/src` → nada.
- [ ] **Step 4: Conferência visual** com o app rodando (skill `run`, ou `npm run dev` + backend local): Estoque, Periféricos, Vincular, Histórico, Relatórios, Indicadores e a ficha rápida em 390, 768 e 1440 px, nos temas claro e escuro — sem rolagem horizontal da página (só dentro das tabelas), gaveta funcionando abaixo de 1024 px, formulários em uma coluna em 390 px. Se o ambiente não conseguir subir o app, registrar "pendente → teste de fumaça do Passo 6 do plano mestre" no relatório, com esta lista de telas.
- [ ] **Step 5: Preencher o checklist** no relatório da fase (evidência = teste ou arquivo):

| Item | Evidência |
|---|---|
| F1.1 / 7.1 / 7.6 / 7.7 paleta, painéis brancos/pretos, caixas azuis | receita (Decisão 10); `palette.test.ts` (par `--brand-text`×`--canvas`); testes `bg-brand-soft` (Tasks 9) e `bg-lime` (Task 4) |
| F1.2 / 2.1 / 2.2 / 7.11 paginação 25/50/100 e "ir para" | `StockPage`, `PeripheralsPage`, `HistoryPage`, `ReportPage` tests; `listagens.guard.test.ts`; `Pagination.test.tsx` |
| F1.2 / 2.3 / 7.9 status em pílula, cor por status, negrito | asserções `font-bold` nas Tasks 6, 7, 10; guarda de selos locais |
| F1.2 / 7.8 carregamento e clique único | testes de duplo clique nas Tasks 4, 7, 8, 10; `button.test.tsx` |
| F1.3 / 7.10 unidade/revenda em CAIXA ALTA (tela e exportação) | Tasks 6, 7, 8, 9, 10; `test_export_csv_revenda_em_caixa_alta`; CSV do histórico |
| F1.4 / 2.4 / 7.3 Setor, Cargo, filtros, ordenação no servidor | `test_items_ordenacao.py`; `StockPage.test.tsx` |
| F1.5 / 7.4 / 8.6 detalhe com colaborador, setor, cargo, superior, local, unidade | `ItemDetailPage.test.tsx` (F0, aba Dados + Histórico); `ItemDetailsModal.test.tsx` (Task 10); linha do tempo = aba da F4 |
| F1.6 / 7.5 responsividade | `AppLayout.test.tsx`; teste `overflow-x-auto` do Estoque; Step 4 |
| F1.7 consistência (títulos, espaçamentos, botões) | `PageHeader` em todas as telas da F1; receita |
| Aceite "nenhuma tela com 7" | guarda (pendências nomeadas da F3) |
| Aceite "contraste AA nos dois temas" | `palette.test.ts` |

- [ ] **Step 6: Limpeza:** o script de testes usa `docker run --rm` (nenhum container fica); o database `stock_sys_test_f1` no MySQL de teste fica para o Passo 6 do plano mestre. Nada de `docker volume prune`.
- [ ] **Step 7: Commit** só se houve correção (`fix(...)`), com a linha `Co-Authored-By`.
