# PROMPT MASTER — ANTIGRAVITY
## Redesign Total do Stock System (Gestão de Patrimônio / IT Asset Management)
### Direção de Arte: "Enterprise Editorial" — v3 (adaptado de Figma para execução direta em código)

> Este é o prompt único e definitivo para o Antigravity implementar o redesign completo do sistema **diretamente no repositório real**, e não em uma ferramenta de design. Ele deve ser lido e executado integralmente, na ordem apresentada. Nenhuma seção pode ser ignorada, resumida, comprimida ou substituída por interpretação própria do agente.
>
> **Sobre esta versão:** esta é uma adaptação do "Prompt Master — Figma v2". A versão anterior tratava o resultado como artefato de design (frames, componentes Figma, páginas de biblioteca). Esta versão trata o resultado como **produto de software funcional**: componentes React reais, tokens em código (Tailwind + CSS variables), rotas reais no TanStack Router, estados reais de dados (loading/error/empty), acessibilidade real testável no navegador. Onde a v2 dizia "criar um frame no Figma", esta versão diz "criar/editar um arquivo no repositório". Todos os tokens concretos de cor, tipografia, forma e as 8 referências de estilo estudadas foram preservados e traduzidos para artefatos de código.
>
> **Formato de execução:** o Antigravity deve tratar este documento como um plano de trabalho multi-fase. Antes de escrever qualquer componente de tela, ele deve produzir e persistir no repositório os artefatos das Fases 1 a 4 (Seção 19). Pular fases, ou entregar apenas Dashboard + Estoque, é uma execução reprovada (ver Seção 2 e Seção 20).

---

## 0. PAPEL DO AGENTE

O Antigravity deve operar, simultaneamente, como:

- Lead Product Designer com fluência em código de produção;
- Art Director responsável por consistência visual entre todas as telas;
- Engenheiro Front-end Sênior (React, TypeScript, Tailwind, Radix/shadcn);
- Especialista em Design Systems tokenizados (não em bibliotecas de componentes soltas);
- Especialista em produtos Enterprise B2B;
- Especialista em interfaces de gestão de dados densas (tabelas, filtros, formulários longos, fluxos operacionais).

O Antigravity está implementando um **produto real, que será vendido a um cliente**, dentro de um repositório real. Não é uma prova de conceito, não é uma tela de demonstração isolada, não é um componente solto para "mostrar a ideia". É um sistema corporativo completo de Gestão de Patrimônio / IT Asset Management, e o resultado final — o código, rodando no navegador — precisa parecer construído por uma equipe de produto especializada, nunca gerado automaticamente por IA.

**Definição de "concluído" para este trabalho:** o critério de sucesso não é "o componente compila". É: (a) todas as 17 rotas da Seção 8 existem e renderizam; (b) nenhuma delas reutiliza a composição de outra fora do que está descrito na Seção 9; (c) os tokens da Seção 6 estão implementados como fonte única de verdade (nunca valores hardcoded espalhados pelos componentes); (d) o checklist da Seção 20 passa tela por tela.

---

## 1. CONTEXTO DO PRODUTO E STACK REAL

O sistema gerencia: computadores, notebooks, monitores, smartphones, periféricos, equipamentos, patrimônio, usuários, unidades, empréstimos, devoluções, movimentações, histórico, relatórios, offboarding e inventário — para uso de equipes corporativas de TI, patrimônio e administração.

A interface precisa transmitir: confiança, controle, precisão, organização, segurança, eficiência, maturidade, profissionalismo.

**Stack real do produto (o Antigravity deve trabalhar dentro desta stack, não propor alternativas):**

| Camada | Tecnologia |
|---|---|
| Framework | React + TypeScript |
| Build | Vite |
| Roteamento | TanStack Router |
| Estilo | Tailwind CSS |
| Componentes base | Shadcn UI + Radix UI |
| Ícones | Lucide React |
| Dados assíncronos | React Query |
| Formulários | React Hook Form + Zod |
| Gráficos | Recharts |
| 3D (somente tela de Login) | Three.js |

**Regra de sobrescrita:** os componentes deste design system **substituem os tokens padrão do Shadcn** (radius, cor, sombra, tipografia) — a estrutura, o comportamento e os nomes de componentes do Shadcn/Radix devem ser mantidos (acessibilidade, composição, variantes via `class-variance-authority`), mas todo valor visual segue exclusivamente os tokens da Seção 6. Na prática: editar `tailwind.config.ts`, o arquivo de tokens/CSS variables (`globals.css` ou equivalente) e os arquivos em `components/ui/*` — nunca criar uma segunda camada de estilo paralela ao Shadcn.

**Modo dark-first com suporte a tema claro:** o produto é dark-first (canvas escuro como padrão), com tema claro completo como alternativa — implementado via `next-themes`/`ThemeProvider` (ou equivalente já presente no repositório) e CSS variables com fallback. Os dois temas são cidadãos de primeira classe: nenhum componente pode ser implementado "só para o dark" e remendado depois para o light. A Seção 6 traz os tokens para os dois modos, prontos para virar CSS variables.

---

## 2. O QUE ACONTECEU ATÉ AGORA (PROBLEMA A CORRIGIR)

Uma tentativa anterior de implementação **não foi aprovada**. Os problemas identificados foram:

- apenas **Dashboard** e **Estoque** foram de fato desenvolvidos — o restante do sistema ficou incompleto;
- a tela de **Cadastro** reutilizou quase integralmente a estrutura visual e de componentes do **Estoque** (mesma composição, mesma lógica de tabela/lista, sem adaptação à sua própria função de formulário);
- composição previsível, genérica, com cara de SaaS de template;
- pouca personalidade, pouca identidade visual própria, tokens de cor/radius não centralizados (valores hardcoded espalhados nos componentes);
- padrões visuais típicos de dashboards gerados por IA (grade de 4 KPIs, ícones em círculos coloridos, gradientes decorativos).

**Este prompt existe justamente para impedir que isso se repita.** Não é uma continuação incremental do trabalho anterior — é um recomeço com direção de arte, sistema de tokens e escopo obrigatórios, verificáveis a cada fase.

**O trabalho só está concluído quando TODAS as rotas da Seção 8 existirem e renderizarem, cada uma com composição própria — nunca apenas Dashboard + Estoque, e nunca telas que são cópias estruturais umas das outras.**

---

## 3. REFERÊNCIAS DE QUALIDADE (ESTUDAR, NUNCA COPIAR)

**Product design** — estudar precisão, densidade, hierarquia, navegação, estados, acabamento, tipografia, espaçamento, minimalismo:
- Linear — linear.app
- Vercel — vercel.com
- Notion — notion.com
- Figma — figma.com

**Enterprise / ITAM** — estudar ciclo de vida de ativos, tabelas, filtros, operações, densidade de dados, workflows corporativos:
- ServiceNow (ITAM) — servicenow.com/br/products/it-asset-management.html
- Snipe-IT — snipeitapp.com
- Freshservice — freshworks.com/freshservice
- ManageEngine — manageengine.com
- Lansweeper — lansweeper.com

**Referências de padrão de componente (estudadas nesta rodada — ver Seção 6.2 para o que foi aproveitado de cada uma):** shadcn/ui, Shop, Airtree (Home), Hyperstudio, Cursor, Air, Gsap, Factory.

**Regra absoluta:** nunca copiar layout, componentes, cores, logos, páginas ou identidade dessas referências. Extrair apenas comportamento, qualidade e organização de dados — no caso das 8 referências de estilo, extrair apenas **mecânica de composição** (como um stat block se estrutura, como uma sidebar se organiza, como um card aninhado separa cabeçalho de corpo), nunca cor, radius grande, tipografia decorativa ou efeitos. **A identidade final deve ser 100% original** e específica para gestão patrimonial — não uma identidade genérica "de startup".

---

## 4. DIREÇÃO VISUAL — "ENTERPRISE EDITORIAL"

Características obrigatórias: elegante, sóbria, técnica, sofisticada, corporativa, autoral, precisa, informacional.

A sofisticação vem de **tipografia, proporção, grid, composição, alinhamento, densidade, contraste, detalhe e consistência** — nunca de efeitos visuais decorativos, e nunca de bibliotecas de animação usadas como muleta estética.

**Teste de validação da direção** (aplicar mentalmente antes de codar qualquer tela, e de novo depois de renderizada): *"Se eu remover o logo, essa interface ainda parece ter identidade própria e específica para ITAM — ou poderia pertencer a qualquer SaaS genérico gerado por IA?"* Se a resposta for a segunda opção, a implementação está errada e o componente precisa ser refeito antes de seguir para a próxima tela.

---

## 5. LISTA NEGRA — PROIBIDO EM QUALQUER TELA OU COMPONENTE

Proibições visuais e de classe Tailwind:

- `rounded-xl`, `rounded-2xl`, `rounded-3xl` ou qualquer arredondamento exagerado (inclui radius 18–37px e `rounded-full`/pill vistos em Shop, Home/Airtree e Gsap — **não usar em botão, input, badge, card**);
- glassmorphism, `backdrop-blur-*`;
- gradientes decorativos (`bg-gradient-to-*` com roxo/azul, mesh gradients, gradiente verde estilo Gsap);
- glow, neon, `box-shadow` com cor saturada, efeitos futuristas genéricos;
- sombras grandes / excesso de elevação (nada equivalente a `shadow-lg`/`shadow-xl` violeta ou colorido);
- cards flutuantes em excesso, card dentro de card sem função;
- excesso de `Badge`/pill decorativos sem significado de status;
- ícones dentro de círculos coloridos sem função semântica;
- **4 cards de KPI idênticos** (`grid grid-cols-4` de stat cards) como estrutura principal de qualquer tela;
- gráficos Recharts puramente decorativos, sem eixo/legenda funcional;
- títulos gigantes / textos genéricos de placeholder (nenhum `text-6xl`/`text-7xl`+ como o Prody 131px de Airtree ou o Mori 224px de Gsap — teto absoluto: 28–32px de título de produto, ver Seção 6.2);
- espaços vazios artificiais sem função editorial real;
- visual de landing page, de template, de Lovable, de Shadcn genérico "fora da caixa", de "AI dashboard";
- elementos decorativos sem função, formas geométricas soltas no SVG/CSS (inclui "sistema de pontos" decorativo estilo Airtree);
- cores de acento saturadas fora da paleta navy institucional definida na Seção 6 (nada de violeta `#5433eb`, amarelo elétrico `#ffff48`, laranja `#f54e00`/`#ee6018`, verde `#0ae448` como cor de marca ou CTA — semânticas funcionais continuam permitidas, ver Seção 6);
- transformar **todo** elemento de página em `<Card>`;
- reutilizar a mesma composição/JSX estrutural em telas com funções diferentes (ver Seção 9 — se o código de duas telas é praticamente o mesmo componente com props trocadas, está errado).

**Verificação automatizável:** ao final de cada tela, rodar uma busca textual no arquivo por `rounded-xl|rounded-2xl|rounded-full|backdrop-blur|shadow-lg|shadow-xl|gradient` — qualquer ocorrência fora de uma exceção documentada (Seção 6.3) é falha de execução.

---

## 6. FOUNDATIONS — TOKENS CONCRETOS EM CÓDIGO (DARK-FIRST + LIGHT)

Esta seção define os valores que devem existir **como CSS variables/tema Tailwind**, nunca como valores soltos em `className`. Local sugerido: `src/styles/globals.css` (variáveis) + `tailwind.config.ts` (mapeamento semântico).

### 6.1 Cor

**Tema escuro (padrão / dark-first)** — inspirado na disciplina monocromática de Hyperstudio e Factory (canvas quase-preto, tipografia clara, bordas hairline, contraste por camadas — nunca por blur ou glow):

| Token semântico | Valor | Papel |
|---|---|---|
| `--canvas` | `#0b1120` | Fundo de página, base de todo o produto no modo escuro |
| `--surface` | `#111827` | Sidebar, painéis elevados, cabeçalho de tabela |
| `--surface-alt` | `#161f33` | Card, popover, linha de tabela em hover |
| `--hairline` | `#232d42` | Bordas, divisores, contorno de input |
| `--text-primary` | `#e7ebf3` | Títulos, valores, texto de alta ênfase |
| `--text-secondary` | `#94a0b8` | Texto de apoio, labels, placeholders |
| `--navy` | `#3b5bdb` | Cor de identidade — logo, estado ativo, ação primária. Uso pontual, nunca dominante |
| `--navy-soft` | `#1c2c52` | Fundo de estado ativo/selecionado, hover em item de navegação |

**Tema claro** — off-white/grafite, mesmo princípio de camadas por tom, não por sombra:

| Token semântico | Valor | Papel |
|---|---|---|
| `--canvas` | `#f6f7f9` | Fundo de página |
| `--surface` | `#ffffff` | Sidebar, cards, tabelas |
| `--surface-alt` | `#eef1f6` | Hover de linha, badge neutro |
| `--hairline` | `#dde2ea` | Bordas, divisores |
| `--text-primary` | `#111827` | Títulos, valores |
| `--text-secondary` | `#5b6472` | Texto de apoio, labels |
| `--navy` | `#2f4bd0` | Idêntico ao papel do dark, ajustado para contraste em fundo claro |
| `--navy-soft` | `#e5eaff` | Fundo de estado ativo/selecionado em fundo claro |

**Semânticas funcionais (idênticas nos dois temas, ajustando apenas a versão "soft" de fundo):**

| Token semântico | Cor forte | Uso |
|---|---|---|
| `--success` | `#2f9e5c` | Status "Ativo", confirmação, disponível |
| `--warning` | `#c98a1c` | Manutenção, pendente, validade próxima |
| `--destructive` | `#d1453d` | Baixa, remoção, erro de formulário |
| `--info` | igual a `--navy` | Estado ativo, link, ação primária |
| `--neutral` | igual a `--text-secondary` | Inativo, arquivado, sem status |

**Implementação obrigatória:** todas as classes Tailwind de cor usadas em componentes devem referenciar esses tokens semânticos (`bg-canvas`, `bg-surface`, `text-primary`, `border-hairline`, `bg-navy`, etc., mapeados em `tailwind.config.ts` a partir das CSS variables acima) — nunca `bg-[#0b1120]` hardcoded dentro de um componente de tela. Isso garante que trocar de tema seja uma troca de variável CSS, não uma reescrita de componente.

### 6.2 Tipografia

- **Plus Jakarta Sans** — interface geral (nav, labels, corpo de texto, botões, headings de tela).
- **IBM Plex Mono** — patrimônio, serial, ID, código de barras, datas técnicas, qualquer número que precise ser lido/comparado com precisão em tabela.

Ambas devem ser carregadas via `next/font`, `@fontsource` ou `<link>` de Google Fonts com `font-display: swap`, e expostas como tokens Tailwind (`font-sans`, `font-mono`) — nunca importadas ad-hoc por tela.

Escala tipográfica (aplicável aos dois temas — apenas a cor de texto muda), a implementar como classes utilitárias ou componentes `<Text variant="...">`:

| Token | Tamanho | Peso | Uso |
|---|---|---|---|
| `caption` | 12px | 500 | Labels de status, legendas, uppercase de metadado |
| `body-sm` | 13px | 400 | Texto de tabela, texto secundário |
| `body` | 14px | 400 | Corpo padrão |
| `body-lg` | 16px | 500 | Texto de destaque em formulário |
| `heading-sm` | 18px | 600 | Título de seção dentro de tela |
| `heading` | 22px | 600 | Título de página |
| `heading-lg` | 28px | 600 | Título de dashboard executivo / stat block principal |

Nunca ultrapassar 28–32px em nenhum título de produto — display type editorial gigante está fora de escopo; a hierarquia aqui é sóbria e funcional, não uma peça de marca.

### 6.3 Forma, espaçamento e sombra

- **Radius (`tailwind.config.ts` → `borderRadius`):** `2px` (inputs pequenos, badges), `4px` (padrão — botões, inputs, linhas de tabela), `6px` (cards), **máximo absoluto `8px`** com justificativa explícita em comentário no código (ex.: avatar circular). Nunca `9999px`/pill, nunca 18–37px.
- **Bordas:** hairline `1px` como principal recurso de separação visual (herdado de Cursor, Hyperstudio, Factory) — preferir `border border-hairline` + espaçamento + tipografia a `shadow-*` + elevação.
- **Sombra:** reservada a overlays reais (`Dialog`, `DropdownMenu`, `Toast` do Radix) — `0 4px 16px rgba(0,0,0,0.24)` no dark, `0 2px 12px rgba(16,24,40,0.08)` no light, definidas como token `shadow-overlay` em `tailwind.config.ts`. Nunca sombra decorativa em card de conteúdo estático.
- **Spacing base:** unidade `4px` — escala `4/8/12/16/20/24/32/48/64`, usando a escala padrão do Tailwind (`p-1`…`p-16`) sem valores arbitrários fora dela.
- **Grid/largura máxima:** `max-w-[1440px]` em telas de dados densos (Estoque, Relatórios), `max-w-[1120px]` em telas de leitura/formulário (Termos, Detalhes), aplicado no container de página, nunca em cada componente filho individualmente.

---

## 6.1-B COMPONENTES — ADAPTADOS DO SHADCN/RADIX, COM MECÂNICA EMPRESTADA DAS REFERÊNCIAS

O produto já usa Shadcn UI + Radix UI como base técnica — portanto os componentes abaixo mantêm os **nomes de arquivo, props e comportamento** do Shadcn (`components/ui/button.tsx`, `components/ui/badge.tsx`, etc.), mas todos os valores visuais (radius, cor, sombra, densidade) devem ser sobrescritos para seguir os tokens da Seção 6, nunca os defaults gerados pelo `shadcn add`. A mecânica de composição de cada componente foi adaptada de uma referência específica — ver origem entre parênteses.

### Sidebar Nav Rail — `components/layout/sidebar.tsx` (mecânica: shadcn/ui)
Fundo `bg-surface`, sem gradiente, sem ícone dentro de círculo colorido. Item ativo: fundo `bg-navy-soft`, `rounded` (4px), texto/ícone em `text-navy`. Item em hover: `bg-surface-alt`. Estrutura em grupos (Principal / Inventário / Operações / Gestão) com label de grupo em `caption` uppercase, `text-secondary` — nunca divisor pesado, apenas espaçamento (`space-y-*`).

### Breadcrumb Trail — `components/layout/breadcrumb.tsx` (mecânica: shadcn/ui)
Texto inline, separadores em chevron (Lucide `ChevronRight`), 14px, sem fundo, sem borda — segmento atual em `text-primary`, demais em `text-secondary`. Hierarquia puramente tipográfica.

### Search Trigger / Command Palette — `components/layout/command-palette.tsx` (mecânica: shadcn/ui)
Fundo `bg-surface-alt`, `rounded` (4px), `px-2.5 py-2`, indicador de atalho (⌘K) alinhado à direita em `caption`/`text-secondary`. Abre um `CommandDialog` (Radix `cmdk`) com busca global de ativos, usuários e ações — nunca decorativo.

### Stat Block — `components/data/stat-block.tsx` (mecânica: shadcn/ui — usado com moderação, nunca como grade de 4 cards idênticos)
Label em `caption` uppercase (`text-secondary`), valor em `heading-lg` com `font-mono` quando for número de patrimônio/quantidade, `font-sans` quando for métrica textual. Sem chrome de card — a hierarquia tipográfica sozinha comunica o dado. Usado no Dashboard Executivo e nos cabeçalhos de Relatórios, nunca como grade repetitiva (`grid-cols-4` de cards idênticos é proibido — ver Seção 5).

### Nested Card Header/Footer — `components/data/nested-card.tsx` (mecânica: shadcn/ui)
Radius assimétrico (topo `rounded-t-md`, base `rounded-b-md`), fundo transparente sobre o card, usado para separar zona de ação (ex.: "Exportar", "Filtros ativos") do corpo de dados sem introduzir nova cor.

### Badge — Solid / Soft / Outline — `components/ui/badge.tsx` (mecânica: shadcn/ui, radius reduzido)
Radius `rounded` (4px, nunca pill). `variant="solid"` usa a cor semântica forte com texto invertido; `variant="soft"` usa a versão de fundo tonal (10–15% de opacidade sobre `surface`); `variant="outline"` é hairline sem fundo, para status informativo neutro. Usado para Status de ativo, papel de usuário, tipo de movimentação.

### Botão Primário / Secundário / Outline — `components/ui/button.tsx` (mecânica: shadcn/ui, radius reduzido)
`variant="default"` (primário): fundo `bg-navy`, texto branco/off-white, `rounded` (4px). `variant="secondary"`: fundo `bg-surface-alt`, `text-primary`. `variant="outline"`: transparente, `border border-hairline`. Nunca pill; altura `h-9`/`h-10` (36–40px); `font-medium`.

### Figure/Ground Panel — `components/patterns/figure-ground-panel.tsx` (mecânica: Factory)
Para fluxos de ação crítica (Remoção de Ativo, confirmação de Offboarding), usar um painel `bg-surface` destacado sobre o `bg-canvas` escuro (ou `surface` sobre `canvas` claro no tema light) — contraste de figura/fundo em vez de sombra pesada, reservado exclusivamente a esses momentos de decisão crítica, nunca em telas de navegação comum.

### Camadas por tom, não por sombra (mecânica: Cursor)
`canvas → surface → surface-alt` formam a escala de elevação do produto inteiro (nos dois temas). Um elemento "sobe" de nível trocando de classe de fundo, não ganhando `shadow-*` — reservar sombra real apenas para overlays flutuantes do Radix (Seção 6.3).

---

## 6.2-B O QUE FOI APROVEITADO DE CADA REFERÊNCIA (E O QUE FOI DESCARTADO)

| Referência | Aproveitado (mecânica) | Descartado (não usar) |
|---|---|---|
| **shadcn/ui** | Nomenclatura e comportamento de componente (Stat Block, Nested Card, Search Trigger, Breadcrumb, Badge, Sidebar), disciplina monocromática, hairline border como definidor de card | Radius 18–24px, geometria pill em botão/input/badge |
| **Shop** | — (avaliado e rejeitado quase por completo) | Radius 20–28px, pills `9999px`, acento violeta saturado, cards flutuantes de imagem, densidade de e-commerce |
| **Airtree / Home** | — (avaliado e rejeitado quase por completo) | Display serif gigante (131px), acento amarelo elétrico, radius 37px, sistema de pontos decorativos, layout de site institucional |
| **Hyperstudio** | Canvas quase-preto com contraste por tipografia e hairline, componentes "esqueleto" (outline, ghost, sem sombra), acento pontual em vez de cor dominante | Tom editorial-tech de marca de estúdio, uso de dourado/gold como acento de ícone |
| **Cursor** | Cantos retos/quase retos (4px), camadas por tom (canvas → bone → linen), bordas hairline warm-neutral, restrição tipográfica | Paleta parchment/cream, acento ember laranja, serif editorial em subheadings |
| **Air** | — (avaliado e rejeitado quase por completo) | Fotografia full-bleed, formas de vidro 3D, itálico cursivo decorativo |
| **Gsap** | — (avaliado e rejeitado quase por completo) | Display type 224px, cor por taxonomia saturada (verde/laranja/rosa/violeta/azul), botões pill 100px |
| **Factory** | Contraste figura/fundo (card claro sobre canvas escuro) para momentos de alta atenção, disciplina monocromática com 2 acentos funcionais (nunca decorativos), tipografia condensada e precisa | Uso de laranja como acento visual recorrente fora de estados de dado vivo |

Este mapeamento é vinculante: qualquer elemento listado na coluna "Descartado" que aparecer em qualquer tela do código é uma falha de execução e deve ser corrigido antes de prosseguir (ver Seção 20).

---

## 7. NAVEGAÇÃO

**Sidebar** (`components/layout/sidebar.tsx`, Seção 6.1-B; discreta, sem fundos pesados, ícones grandes, badges ou gradientes; estado ativo claro porém elegante), organizada em grupos:

- **Principal:** Dashboard
- **Inventário:** Estoque, Cadastro, Histórico
- **Operações:** Empréstimos, Devoluções, Periféricos, Offboarding
- **Gestão:** Usuários, Unidades, Relatórios

**Topbar** (`components/layout/topbar.tsx`): Breadcrumb Trail, nome da seção, Search Trigger (⌘K), atalhos, alternância de tema claro/escuro, avatar/menu de usuário — minimalista, sem competir com o conteúdo.

Ambos devem ser implementados como layout compartilhado no TanStack Router (rota raiz `__root.tsx` ou equivalente), nunca duplicados por página.

---

## 8. MAPA COMPLETO DE ROTAS (OBRIGATÓRIAS — NENHUMA PODE FICAR DE FORA)

| # | Tela | Rota | Arquivo sugerido |
|---|---|---|---|
| 01 | Dashboard Executivo | `/` | `routes/index.tsx` |
| 02 | Estoque | `/stock` | `routes/stock.tsx` |
| 03 | Cadastro de Ativo | `/register` | `routes/register.tsx` |
| 04 | Edição de Ativo | `/edit/:id` | `routes/edit.$id.tsx` |
| 05 | Remoção de Ativo | `/remove` | `routes/remove.tsx` |
| 06 | Empréstimos | `/loan` | `routes/loan.tsx` |
| 07 | Devoluções | `/return` | `routes/return.tsx` |
| 08 | Termos | `/terms` | `routes/terms.tsx` |
| 09 | Periféricos | `/peripherals` | `routes/peripherals.tsx` |
| 10 | Vincular Periférico | `/link` | `routes/link.tsx` |
| 11 | Gráficos | `/charts` | `routes/charts.tsx` |
| 12 | Histórico | `/history` | `routes/history.tsx` |
| 13 | Relatórios | `/report` | `routes/report.tsx` |
| 14 | Unidades | `/unidades` | `routes/unidades.tsx` |
| 15 | Usuários | `/users` | `routes/users.tsx` |
| 16 | Offboarding TI | `/offboarding-ti` | `routes/offboarding-ti.tsx` |
| 17 | Detalhes do Ativo | modal/painel | componente `AssetDetailsPanel`, aberto sobre qualquer rota relevante |

O sistema **não está concluído** se existirem apenas Dashboard e Estoque, e **não está concluído** se várias dessas rotas forem apenas variações copiadas do componente de Estoque com poucas props trocadas.

---

## 9. CADA TELA TEM SUA PRÓPRIA EXPERIÊNCIA — REGRA CRÍTICA

Componentes podem (e devem) ser reutilizados via design system (Seção 6.1-B). **A composição inteira, não.** É proibido o padrão:

> Estoque = tabela · Cadastro = tabela · Usuários = tabela · Unidades = tabela · Periféricos = tabela · Relatórios = cards

Isso é CRUD genérico e é exatamente o que este prompt existe para evitar. Cada rota abaixo precisa de composição adequada à sua função:

- **Dashboard** (`routes/index.tsx`) → visão executiva (ver Seção 10), usando Stat Block com moderação, nunca grade de KPIs.
- **Estoque** (`routes/stock.tsx`) → operação e consulta; tabela como protagonista, alta densidade (ver Seção 11).
- **Cadastro** (`routes/register.tsx`) → formulário profissional dividido em seções lógicas (ver Seção 12); **jamais** reaproveitar a composição de tabela do Estoque.
- **Edição** (`routes/edit.$id.tsx`) → variação do cadastro com estado atual pré-preenchido via `defaultValues` do React Hook Form, diffs claros quando aplicável.
- **Remoção** (`routes/remove.tsx`) → fluxo de ação destrutiva com confirmação explícita, usando o Figure/Ground Panel (Seção 6.1-B), contexto do item, aviso claro de irreversibilidade.
- **Empréstimos / Devoluções** (`routes/loan.tsx`, `routes/return.tsx`) → fluxo operacional passo a passo (stepper), com contexto do ativo e do responsável visível durante todo o processo.
- **Termos** (`routes/terms.tsx`) → composição de documento corporativo (leitura, não tabela).
- **Periféricos** (`routes/peripherals.tsx`) → relação entre equipamentos, visualização de vínculo.
- **Vincular Periférico** (`routes/link.tsx`) → fluxo de associação ativo↔periférico, busca e confirmação (reaproveitar o `CommandPalette`/Search Trigger).
- **Gráficos** (`routes/charts.tsx`) → análise visual funcional (não decorativa) com Recharts, foco em tendência e comparação.
- **Histórico** (`routes/history.tsx`) → linha do tempo de eventos e movimentações, não tabela genérica.
- **Relatórios** (`routes/report.tsx`) → consulta e geração, com preview e exportação.
- **Unidades** (`routes/unidades.tsx`) → estrutura organizacional (hierarquia/localização), não apenas lista.
- **Usuários** (`routes/users.tsx`) → administração, papéis, permissões, vínculo com ativos.
- **Offboarding** (`routes/offboarding-ti.tsx`) → processo em etapas (checklist/stepper), não formulário solto.
- **Detalhes do Ativo** (`AssetDetailsPanel`) → ficha técnica completa, organizada em seções com divisores hairline (não card por seção).

---

## 10. DASHBOARD EXECUTIVO — `routes/index.tsx`

**Proibido:** estrutura de "KPI + KPI + KPI + KPI" (`grid-cols-4` de stat cards) como composição principal.

A tela deve responder, com hierarquia editorial (Stat Block pontual, não grade de cards):

- Como está o patrimônio?
- O que exige atenção?
- O que mudou recentemente?
- Como está a operação?
- Quais problemas existem?
- Quais tendências existem?

Sugestão de composição: uma faixa superior com 2–3 Stat Blocks tipográficos (não cards), seguida de um bloco editorial assimétrico combinando um gráfico Recharts funcional (tendência) com uma lista compacta de itens que exigem atenção (`text-sm`, hairline dividers), e uma seção de "mudanças recentes" como linha do tempo compacta — nunca três grades de cards empilhadas.

---

## 11. ESTOQUE — `routes/stock.tsx`

Tabela como protagonista absoluto. Construir com React Query + Tanstack Table (ou equivalente já usado no repositório): busca (`Search Trigger`), filtros, filtros avançados, ordenação por coluna, seleção múltipla (checkbox), indicadores de status semânticos discretos (`Badge` — Seção 6.1-B), ações contextuais por linha (dropdown), paginação real (não infinita, para manter previsibilidade em contexto ITAM). Alta densidade informacional — esta tela deve parecer uma ferramenta profissional de gestão de ativos, não uma listagem simples.

Colunas de referência: Patrimônio, Tipo, Descrição, Usuário, Unidade, Status, Localização, Atualizado em, Ações. Usar `font-mono` (IBM Plex Mono) para dados técnicos (patrimônio, serial). Edição via `Dialog`/`Sheet` (modal), tabela padronizada em 7 itens visíveis por página + scroll horizontal quando necessário em telas menores.

---

## 12. CADASTRO DE ATIVO — `routes/register.tsx`

Formulário profissional, denso e organizado, dividido em seções lógicas (`<fieldset>`/blocos com `heading-sm`): **identificação, equipamento, localização, responsabilidade, situação, observações.** Implementar com React Hook Form + Zod (schema de validação único, reaproveitado em Cadastro e Edição). Estados de campo obrigatórios: label, input, texto de apoio (`description`), validação inline, mensagem de erro, `disabled`, `loading` (submit em andamento) — em ambos os temas.

---

## 13. TELA DE LOGIN — NÃO REIMPLEMENTAR

A tela de login está **fora do escopo deste redesign** e deve ser preservada exatamente como está no repositório atual, incluindo: identidade visual, animação 3D (logotipo "revalle" em `logo.glb`, Three.js), logo, composição, câmera, iluminação, materiais e timing das animações. Requisito adicional obrigatório, se ainda não estiver satisfeito: **zero scroll vertical e zero scroll horizontal**, em qualquer resolução. O Antigravity não deve tocar nos arquivos da rota de login além do estritamente necessário para corrigir esse requisito de scroll, se aplicável.

---

## 14. DESIGN SYSTEM NO CÓDIGO

**Foundations (arquivos de configuração, não telas):**
- `tailwind.config.ts` — cores semânticas (dark + light via CSS variables — Seção 6.1), radius (Seção 6.3), tipografia (`fontFamily`), spacing.
- `src/styles/globals.css` (ou equivalente) — declaração das CSS variables por tema (`:root` e `.dark`), import das fontes.
- `src/lib/tokens.ts` (opcional, recomendado) — constantes TypeScript espelhando os tokens para uso em Recharts/SVG onde classes Tailwind não se aplicam diretamente.

**Components (`src/components/ui/*`, com variantes de tamanho, estado e ênfase via `cva`, mapeados 1:1 para Shadcn/Radix — ver Seção 6.1-B):** Button, Input, Select, Search, Table, Badge, Status, Tabs, Modal (Dialog), Drawer (Sheet), Dropdown, Pagination, Toast, Empty State, Loading State (Skeleton), Error State, Stat Block, Nested Card Header/Footer, Sidebar Nav Rail, Breadcrumb Trail, Command Palette, Figure/Ground Panel.

**Iconografia:** Lucide Icons (já usado no stack) — `size-4`/`size-5`, consistentes, funcionais, nunca decorativos isolados sem rótulo ou tooltip.

**Regra de dependência:** nenhuma tela (`routes/*`) deve declarar cor, radius ou sombra própria fora dos tokens acima. Se uma tela "precisa" de uma cor nova, a decisão correta é adicionar um token semântico em `tailwind.config.ts`, nunca um valor arbitrário na tela.

---

## 15. ESTADOS OBRIGATÓRIOS

`hover`, `focus`, `active`, `selected`, `disabled`, `loading`, `empty`, `error`, `success`, `destructive` — implementados via variantes de componente (`cva`) para todos os componentes relevantes, como parte formal do design system, e visualmente verificados nos dois temas (dark e light) antes de considerar uma tela concluída.

Estados de dados assíncronos (React Query) devem ter tratamento explícito em cada rota: `isLoading` → Skeleton coerente com o layout final (não spinner genérico central), `isError` → Error State com ação de retry, dados vazios → Empty State com CTA relevante ao contexto (nunca uma ilustração genérica descontextualizada).

---

## 16. RESPONSIVIDADE E ACESSIBILIDADE

Implementar Desktop, Tablet e Mobile **reorganizando** a experiência conforme o espaço (sidebar colapsável/`Sheet` em mobile, tabela → lista de cards compactos em mobile, filtros → `Drawer`), nunca apenas aplicando `overflow-x-auto` no layout desktop e chamando de responsivo.

Acessibilidade: navegação completa por teclado (`Tab`/`Shift+Tab`/`Enter`/`Esc`), estados de foco visíveis (`focus-visible:ring-2 ring-navy`) em todos os componentes interativos, nos dois temas, atributos ARIA corretos herdados do Radix (não removidos por customização de estilo), contraste mínimo AA verificado para texto sobre `canvas`/`surface` em ambos os temas. Incluir o alternador de tema claro/escuro na topbar em todas as resoluções, inclusive mobile.

---

## 17. MOTION

Definir motion (CSS transitions/Framer Motion, o que já estiver disponível no repositório) para: navegação entre rotas, abertura de modal/drawer, seleção de linha/item, aplicação de filtros, transição de loading → conteúdo, feedback de sucesso, mudanças de estado de formulário, transição entre tema claro/escuro (transição de cor suave, não flash). Duração alvo: 120–200ms para micro-interações, 200–300ms para overlays.

Deve ser sutil, rápido, elegante, funcional e coordenado entre componentes (mesma curva de easing em todo o produto — definir uma única `transition-timing-function` padrão em `tailwind.config.ts`). Proibido: bounce, parallax, exagero, efeitos chamativos, animação de entrada em cada item de lista (stagger) usada apenas como decoração.

---

## 18. ORGANIZAÇÃO DO REPOSITÓRIO / DESIGN SYSTEM NO CÓDIGO

Estrutura de pastas de referência (adaptar aos nomes já existentes no repositório real, sem duplicar estrutura paralela):

```
src/
  routes/                 # 17 rotas da Seção 8 + layout raiz
  components/
    ui/                   # primitives shadcn sobrescritos (button, input, badge, table...)
    layout/               # sidebar, topbar, breadcrumb, command-palette
    data/                 # stat-block, nested-card, data-table wrappers
    patterns/             # figure-ground-panel, stepper, timeline
  lib/
    tokens.ts             # espelho TS dos tokens para Recharts/SVG
    schemas/              # Zod schemas (cadastro/edição reaproveitados)
  styles/
    globals.css           # CSS variables dark/light + fontes
```

Todos os componentes devem ser reutilizáveis e importados a partir do design system central — nenhuma tela isolada deve conter JSX de componente "base" duplicado (ex.: um segundo `<Button>` estilizado manualmente dentro de uma rota).

---

## 19. PROCESSO DE EXECUÇÃO OBRIGATÓRIO EM FASES

**Não pular direto para telas.** O Antigravity deve executar nesta ordem, sem saltar etapas, e deve reportar/commitar ao final de cada fase antes de iniciar a próxima:

1. **Análise** — do repositório atual (estrutura existente, o que já usa Shadcn/Radix corretamente, o que precisa ser removido) e das referências de estilo (incluindo a leitura da Seção 6.2-B).
2. **Direção de arte** — confirmação formal da identidade Enterprise Editorial, dark-first + light, documentada em comentário no topo de `globals.css` ou em um `docs/design-system.md` curto.
3. **Foundations** — implementar cor (dois temas), tipografia, grid, spacing, radius, sombra, ícones em `tailwind.config.ts` + `globals.css` (Seção 6).
4. **Componentes** — biblioteca completa em `components/ui` e `components/layout`/`data`/`patterns`, com variantes e estados, mapeada ao Shadcn/Radix (Seção 6.1-B).
5. **Navegação** — sidebar e topbar como layout compartilhado (Seção 7).
6. **Todas as rotas** — as 17 telas da Seção 8, cada uma com composição própria (Seção 9), nesta ordem sugerida: Dashboard → Estoque → Cadastro → Edição → Remoção → Empréstimos → Devoluções → Periféricos → Vincular Periférico → Termos → Gráficos → Histórico → Relatórios → Unidades → Usuários → Offboarding → Detalhes do Ativo.
7. **Estados** — hover, focus, loading, empty, error, success, destructive em todos os componentes, nos dois temas (Seção 15).
8. **Responsividade e acessibilidade** — desktop, tablet, mobile; teclado e foco (Seção 16).
9. **QA visual e funcional** — inspeção final de cada rota contra os testes da Seção 20 e contra a tabela da Seção 6.2-B; rodar a busca textual da Seção 5 por classes proibidas.

Nenhuma fase deve ser considerada concluída sem que o resultado exista de fato no repositório e renderize sem erros.

---

## 20. TESTE ANTI-IA — APLICAR A CADA TELA/ROTA INDIVIDUALMENTE

- Parece SaaS genérico?
- Parece template?
- Parece Lovable ou Shadcn padrão (sem os ajustes de radius/cor da Seção 6)?
- Parece gerada automaticamente por IA?
- Existem cards demais?
- Existem elementos decorativos sem função?
- Algum elemento da coluna "Descartado" da Seção 6.2-B vazou para o código?
- Alguma classe da Lista Negra (Seção 5) aparece no arquivo?
- A rota tem personalidade própria?
- A composição é específica para ITAM ou poderia ser de qualquer produto?
- A informação está bem hierarquizada?
- Os estados de loading/error/empty foram implementados de verdade ou só o caminho feliz?
- Funciona corretamente nos dois temas e em pelo menos três larguras de viewport?
- Parece um software corporativo real, que alguém pagaria para usar?

**Se qualquer resposta falhar o critério, a rota deve ser corrigida antes de prosseguir para a próxima.**

---

## 21. OBJETIVO FINAL

O resultado deve ser: **elegante, corporativo, técnico, autoral, sofisticado, funcional, denso, preciso — e, acima de tudo, código real rodando no navegador**, não uma peça de design estática.

O resultado **não pode** ser: genérico, SaaS padrão, template, "AI slop", nem uma implementação parcial (Dashboard + Estoque apenas).

Critério de sucesso: alguém abrir o produto no navegador e pensar **"isso é um sistema real, pronto para ser vendido"** — nunca **"isso parece gerado automaticamente"**.
