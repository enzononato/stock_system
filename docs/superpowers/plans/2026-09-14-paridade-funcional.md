# Paridade funcional com a branch redesign-frontend — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Fechar as lacunas funcionais entre o frontend atual e `origin/redesign-frontend` — filtros, paginação, gráficos e um Dashboard — sem alterar o backend.

**Architecture:** A branch `redesign-frontend` NÃO será mesclada; serve só como referência de leitura (`git show origin/redesign-frontend:<caminho>`). Descoberta decisiva do levantamento: os nove arquivos `frontend/src/api/*.ts` são byte a byte idênticos nas duas branches, e nenhuma tela da referência chama `api.get`/`api.post` direto. Ela foi escrita contra o backend que já temos. Portanto **tudo neste plano é trabalho de frontend**.

**Tech Stack:** React 18, TypeScript, Vite 5, Tailwind 3, react-router-dom 6, TanStack Query/Table, Radix UI, recharts, vitest + Testing Library.

**Spec:** Este documento. Referência visual e funcional: `origin/redesign-frontend`, telas em `frontend/src/features/<área>/`.

## Global Constraints

- **O BACKEND NÃO PODE SER ALTERADO.** Nenhum arquivo em `backend/` é tocado. Se algo exigir endpoint ou parâmetro novo, não se faz — registra-se como impossível.
- Não mesclar, cherry-pickar ou dar checkout de nada da `redesign-frontend`.
- **Não adicionar dependência npm.** Em particular, o Ctrl+K NÃO usa `cmdk` — é implementado à mão.
- Não remover funcionalidade nem coluna existente. Em vários pontos o build atual é MELHOR que a referência (ordenação e busca no `DataTable`, 18 colunas no histórico contra 6, 19 no relatório contra 11, links de `/remove`, `/terms` e `/link` na sidebar). Nada disso regride.
- Todo texto de interface e todo comentário em português do Brasil.
- Monocromático, exceto cor de status (via `Badge`/`StatusBadge`, com o ponto colorido) e cor destrutiva.
- **Nunca usar modificador de opacidade sobre token customizado** (`bg-primary/50`) — não gera CSS nenhum aqui; o teste-guarda em `src/lib/tokens.test.ts` reprova.
- Sem sintaxe Tailwind 4, sem variantes `dark:`, e `shadow-overlay` é a única sombra permitida (exceção já existente: a tela de login).
- Cada task termina com `npx tsc --noEmit`, `npm run test:run` e `npx vite build` verdes.
- Comandos rodam de `frontend/`.

**Linha de base:** 149 testes em 15 arquivos, tsc limpo, build limpo.

## Contexto que todas as tasks precisam

`DataTable` JÁ suporta paginação no servidor — `HistoryPage` usa. API:
`DataTable<TData>({ data, columns, searchPlaceholder?, className?, pagination? })` com
`pagination: { total, pageIndex, pageSize, onPageChange, search?, onSearchChange? }`.
Adotar esse padrão é o caminho para paginar as outras telas; não reinventar.

`GET /api/items` aceita exatamente `tipo`, `status`, `revenda`, `search`, `limit`, `offset`
(`backend/app/routers/items.py:22-38`). **Não existe parâmetro `setor`** — filtro de setor
no servidor é impossível e não está neste plano.

`settings.MAX_PAGE_SIZE = 500` é TETO, não folga. Nenhuma tela pode depender de buscar tudo.

---

### Task 1: Cinco defeitos do build atual

**Files:**
- Modify: `frontend/src/components/equipment/ItemDetailsModal.tsx`
- Modify: `frontend/src/pages/LoanPage.tsx`
- Modify: `frontend/src/pages/ReturnPage.tsx`
- Modify: `frontend/src/pages/RegisterItemPage.tsx`
- Modify: `frontend/src/components/ui/FileUpload.tsx`

**Interfaces:** consome os componentes atuais; não muda API pública de nada.

- [ ] **Step 1: "Computador" não existe no enum**

`ItemDetailsModal.tsx:46` testa `['Computador', 'Notebook'].includes(item.tipo)`. O enum real
(`backend/app/core/config.py:131-141`) tem **"Desktop"**, nunca "Computador". Resultado: a
seção Hardware & Rede é invisível para todo desktop da frota. Corrigir para os valores reais
do enum. Ler o enum antes e usar exatamente os valores que ele define.

- [ ] **Step 2: Escrever um teste que trava isso**

Criar/ampliar um teste que renderize `ItemDetailsModal` com `tipo: 'Desktop'` e afirme que
um campo da seção Hardware & Rede aparece. Tem que falhar antes da correção do Step 1.

- [ ] **Step 3: Painel de confirmação que não fecha**

`LoanPage.tsx` renderiza `ConfirmacaoTermo` sem `onCancel`, e o botão Cancelar só existe
quando essa prop é passada (`ConfirmacaoTermo.tsx:138-142`). O usuário fica preso até subir
o PDF. Passar um `onCancel` que limpe o estado de confirmação.

- [ ] **Step 4: Devolução confirma às cegas**

`ReturnPage.tsx:124` mostra só `#id` no painel de confirmação. O operador confirma a
devolução sem ver qual equipamento nem de quem. Mostrar tipo, marca, modelo, colaborador e
unidade — todos já presentes no objeto do item.

- [ ] **Step 5: Hidratação do formulário descarta digitação**

`RegisterItemPage.tsx:47-63` faz `setState` em fase de render guardado por `!brand`: limpar
o campo Marca faz o formulário inteiro recarregar do servidor. Trocar por hidratação
uma-única-vez com flag em `useEffect`.

- [ ] **Step 6: FileUpload mente sobre formatos**

`FileUpload.tsx:74` diz "Suporta apenas documentos em formato PDF", mas Remoção e Vínculo
passam tipos de imagem. Derivar o texto dos tipos aceitos que o componente realmente recebe.

- [ ] **Step 7: Verificação e commit**

`npx tsc --noEmit && npm run test:run && npx vite build`. Commit.

---

### Task 2: Estoque — busca no servidor, filtro de unidade e paginação

**Files:** `frontend/src/pages/StockPage.tsx`

**Contexto:** hoje busca `limit: 500` (= o teto do backend) e renderiza tudo. Acima de 500
itens a tabela trunca em silêncio, e os cartões de indicador passam a discordar entre si
porque três são calculados sobre o array buscado e um vem do `total` do servidor.

- [ ] **Step 1:** Adotar paginação no servidor com o padrão que `HistoryPage` já usa:
`pagination={{ total, pageIndex, pageSize, onPageChange, search, onSearchChange }}`,
passando `limit`/`offset` para `listItemsPaginated`. Tamanho de página: 10.
- [ ] **Step 2:** Ligar a busca ao parâmetro `search` do servidor, em vez do filtro global
sobre a página já carregada.
- [ ] **Step 3:** Acrescentar o filtro de Unidade, alimentado por `listUnidades()`, passando
`revenda` ao endpoint.
- [ ] **Step 4:** Corrigir os cartões de indicador para não dependerem do array paginado.
Se um número não puder ser obtido sem varrer tudo, buscar esse total com uma chamada
dedicada de `limit: 1` lendo só o `total` — nunca somando o array da página.
- [ ] **Step 5:** Botão "Limpar filtros" que zere todos, e estado vazio que diferencie
"nenhum item cadastrado" de "nenhum resultado para este filtro".
- [ ] **Step 6:** Verificação e commit.

---

### Task 3: Paginação e busca nas demais telas

**Files:** `LoanPage.tsx`, `ReturnPage.tsx`, `TermsPage.tsx`, `RemovePage.tsx`, `LinkPeripheralPage.tsx`

**Contexto e a distinção que importa:** as cinco buscam `limit: 500`. Mas elas têm dois usos
diferentes, que pedem soluções diferentes:
- **Listas de SELEÇÃO** (escolher o equipamento a emprestar/devolver/remover/vincular):
  paginar é errado — o usuário quer procurar, não folhear. Ligar a busca do
  `SearchableSelect` ao parâmetro `search` do servidor, com `limit` pequeno.
- **Tabelas de NAVEGAÇÃO** (pendentes de confirmação, termos emitidos): paginar no servidor
  com o padrão do `DataTable`.

- [ ] **Step 1:** Classificar cada uma das cinco telas nos dois usos acima e registrar a
classificação em comentário no próprio arquivo.
- [ ] **Step 2:** Aplicar busca no servidor às listas de seleção.
- [ ] **Step 3:** Aplicar paginação às tabelas de navegação.
- [ ] **Step 4:** Remover toda constante `FETCH_ALL_LIMIT = 500`. Ao fim,
`grep -rn "FETCH_ALL_LIMIT" src/` deve voltar vazio.
- [ ] **Step 5:** Em `ReturnPage`, acrescentar ao filtro a condição de ter `assigned_to`
preenchido, além do status — hoje itens órfãos aparecem com Usuário vazio e um "Gerar Termo"
que responde 400.
- [ ] **Step 6:** Em `TermsPage`, acrescentar o controle segmentado Pendentes/Assinados
(re-fatiamento da mesma busca, sem requisição nova).
- [ ] **Step 7:** Verificação e commit.

---

### Task 4: Periféricos — filtros e painel de detalhe

**Files:** `frontend/src/pages/PeripheralsPage.tsx`

**Contexto:** 175 linhas aqui contra 704 na referência. É a maior diferença de tela única.
`GET /api/peripherals` aceita `status` e `tipo` (`backend/app/routers/peripherals.py:19-31`).

**NÃO PORTAR:** a referência tem, em `:127-146`, um `useQueries` que dispara uma requisição
por item com periférico para montar um `peripheralRevendaMap` — que não é usado em lugar
nenhum, só aparece na lista de dependências de um `useMemo`. É uma tempestade N+1 calculando
valor descartado. A query `listUnidades()` em `:115-118` também é buscada e nunca lida.

- [ ] **Step 1:** Filtros de Tipo e Status (via parâmetros do servidor) e busca.
- [ ] **Step 2:** Paginação.
- [ ] **Step 3:** Formulário de criação recolhível, com botão "Novo Periférico".
- [ ] **Step 4:** Painel de detalhe ao clicar na linha: ficha do periférico com as ações
disponíveis. Não duplicar o que a página `/link` já faz — se a ação de vincular já existe
lá, o painel só precisa levar até ela.
- [ ] **Step 5:** Verificação e commit.

---

### Task 5: Gráficos — terceiro gráfico, filtro de unidade e forma de área

**Files:** `frontend/src/pages/ChartsPage.tsx`

**Contexto:** `/api/reports` só expõe `/charts/loans` e `/charts/registrations`.

- [ ] **Step 1:** Terceiro gráfico, "Distribuição Semanal de Empréstimos", barras por dia da
semana (Dom–Sáb). Derivar da resposta de `getLoansChart` que já está na tela — sem busca nova.
- [ ] **Step 2:** Primeiro gráfico passa de barras agrupadas para área com preenchimento em
gradiente, como a referência. Manter a `<Legend>` do recharts, que a referência não tem.
- [ ] **Step 3:** Estado vazio por gráfico ("Nenhuma movimentação neste período") em vez de
eixo vazio, e estado de erro com botão de repetir — hoje nenhuma página lê `error` de
`useQuery`.
- [ ] **Step 4:** Validar o campo de ano (2000–2100, só dígitos, Aplicar desabilitado
enquanto inválido). Hoje `Number('abc')` vira `NaN` e vai para a query.
- [ ] **Step 5: Filtro de Unidade — LER ISTO ANTES DE IMPLEMENTAR.**
Os endpoints de gráfico não aceitam `revenda`. A referência contorna trocando a fonte para
`GET /reports/monthly` e reagregando no cliente. **Dois defeitos reais na versão dela, que
NÃO devem ser reproduzidos:**
  (a) ela agrupa `data_devolucao` por dia sem conferir o mês. O relatório mensal traz só
  empréstimos EMITIDOS no mês, e `data_devolucao` vem de subconsulta sem limite de data —
  então uma máquina emprestada em março e devolvida em abril é contada como devolução de um
  dia de abril NO GRÁFICO DE MARÇO. Filtrar a devolução pelo mês selecionado.
  (b) `GET /reports/monthly` exige papel gestor ou técnico (`reports.py:18`), enquanto os
  endpoints de gráfico não têm guarda de papel. Um Jovem Aprendiz abre Indicadores
  normalmente e recebe 403 ao escolher uma filial. Ou esconder o filtro para esse papel, ou
  tratar o 403 com mensagem clara — decidir e justificar no relatório.
  Se qualquer das duas não puder ser resolvida de forma limpa, **não entregar o filtro** e
  registrar o motivo.
- [ ] **Step 6:** Verificação e commit.

---

### Task 6: Relatório — filtros e blocos de totais

**Files:** `frontend/src/pages/ReportPage.tsx`

- [ ] **Step 1:** Filtro de Unidade e filtro de Tipo de Operação, ambos no cliente sobre as
linhas do mês já buscadas, com as opções derivadas dos próprios dados.
- [ ] **Step 2:** Quatro blocos de totais (Total, Empréstimos, Devoluções, Colaboradores
únicos) usando o `StatBlock` que já existe.
- [ ] **Step 3:** Paginação e estado vazio ciente dos filtros.
- [ ] **Step 4: Dois avisos que precisam de tratamento explícito.**
  (a) Na referência, a exportação CSV ignora os filtros: filtra-se para uma filial, exporta,
  e vem tudo. Aqui a exportação chama o endpoint com ano/mês apenas. **Deixar explícito na
  interface** que o CSV é do mês inteiro, ou desabilitar a exportação enquanto houver filtro
  ativo. Decidir e justificar.
  (b) Filtrar por unidade descarta todas as operações de periférico, porque o ramo de
  periféricos do relatório seleciona `NULL` em revenda
  (`backend/app/db/inventory_manager_db.py:980`). Não dá para corrigir sem mexer no backend.
  Sinalizar isso na interface quando o filtro de unidade estiver ativo.
- [ ] **Step 5:** Verificação e commit.

---

### Task 7: Dashboard em `/`, estoque para `/stock`

**Files:**
- Create: `frontend/src/pages/DashboardPage.tsx`
- Modify: `frontend/src/App.tsx`
- Modify: `frontend/src/components/layout/Sidebar.tsx`

**Contexto:** a referência (`features/dashboard/DashboardPage.tsx`, 402 linhas) compõe:
quatro blocos de totais com sub-rótulo; gráfico de área de 15 dias (Empréstimos ×
Devoluções); lista "Atenção Operacional" com pendentes de confirmação e de devolução, cada um
com botão de link direto; feed "Últimas Movimentações" (`listHistoryPaginated({limit: 6})`);
e um painel de links rápidos.

**Armadilha da referência:** ela agrega os indicadores sobre `listItemsPaginated({limit:500})`
no cliente, então passa de 500 itens os números ficam errados em silêncio — o mesmo teto da
Task 2. Aqui os totais devem vir do campo `total` das respostas paginadas, nunca de somar
um array.

- [ ] **Step 1:** Criar a página com as seções acima.
- [ ] **Step 2:** Mover o estoque para `/stock` e colocar o dashboard em `/`. Acrescentar o
item do estoque na sidebar apontando para a rota nova, e conferir que todo link interno para
o estoque foi atualizado — `grep -rn "to=\"/\"" src/` e `grep -rn "navigate(" src/`.
- [ ] **Step 3:** Respeitar o filtro de papel da sidebar. O dashboard mostra dados de
operação; se algum papel não puder ver alguma seção, esconder a seção, não a página.
- [ ] **Step 4:** Verificação e commit.

---

### Task 8: Shell — menu mobile, Ctrl+K e breadcrumbs

**Files:** `frontend/src/components/layout/AppLayout.tsx`, `Sidebar.tsx`, `TopBar.tsx`, e um
componente novo para a paleta de comandos.

**Contexto:** o app é hoje inutilizável em celular — `AppLayout.tsx:25` sempre renderiza uma
sidebar fixa de `w-64`, sem gaveta. E não existe nenhum manipulador de teclado em todo o
frontend (`grep -rn "keydown|metaKey|ctrlKey" src/` não retorna nada).

- [ ] **Step 1:** Gaveta mobile — abaixo de `lg` a sidebar sai do fluxo e passa a abrir por
botão na TopBar, sobre um overlay. Fechar ao navegar, ao clicar fora e com Escape.
- [ ] **Step 2:** Paleta de comandos com Ctrl+K / ⌘K. **Sem `cmdk`** — implementar à mão:
`keydown` global com `preventDefault`, campo de busca, lista filtrável de destinos de
navegação, navegação por setas, Enter para ir, Escape para fechar. Alimentar a lista a partir
da mesma estrutura de navegação da sidebar, respeitando o filtro por papel, para não oferecer
destino que o usuário não pode abrir.
- [ ] **Step 3:** Breadcrumbs na TopBar (`Grupo › Página`), derivados da mesma estrutura de
navegação.
- [ ] **Step 4:** Substituir o `<Navigate to="/">` silencioso da guarda de papel
(`App.tsx:23`) por uma tela de "Acesso restrito" explicando o que houve.
- [ ] **Step 5:** Verificação e commit.

---

### Task 9: Varredura final

- [ ] **Step 1:** `grep -rn "FETCH_ALL_LIMIT" src/` vazio; nenhuma tela busca 500 linhas.
- [ ] **Step 2:** Varreduras de sempre vazias: paleta antiga, opacidade sobre token,
`dark:`, sombra decorativa fora do login.
- [ ] **Step 3:** `npx tsc --noEmit && npm run test:run && npx vite build`.
- [ ] **Step 4:** `docker compose build frontend` para confirmar que o caminho de deploy não
quebrou.
- [ ] **Step 5:** Conferir que nada regrediu: ordenação e busca do `DataTable`, contagem de
colunas de histórico e relatório, e os links de `/remove`, `/terms` e `/link` na sidebar.

## Notas para quem executa

- A `redesign-frontend` é um RASCUNHO, não evangelho. Já se confirmou nela: um módulo de
  offboarding inteiro que não pode rodar (consulta `status: "emprestado"`, valor que o enum
  não tem), arquivos mortos importados por ninguém, três arquivos que não compilam por ler
  `item.cargo` (campo que não existe), e erros engolidos sistematicamente. Onde ela estiver
  errada, fazer certo e registrar.
- Onde o build atual for melhor, ele vence.
