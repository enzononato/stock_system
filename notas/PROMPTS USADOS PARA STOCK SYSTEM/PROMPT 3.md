# PROMPT MASTER — ANTIGRAVITY
# STOCK SYSTEM / REVALLE — REDESIGN TOTAL ENTERPRISE EDITORIAL v5

> **EXECUÇÃO COMPLETA EM UMA ÚNICA TAREFA**
>
> Este prompt deve ser executado sobre o projeto existente `stock_system`.
> O objetivo é transformar o frontend inteiro em um produto corporativo real de Gestão de Patrimônio / IT Asset Management, preservando a lógica existente e reconstruindo a experiência visual.
>
> **NÃO entregar apenas Dashboard + Estoque.**
>
> **NÃO criar um template visual e reaplicá-lo em todas as páginas.**
>
> **NÃO parar no meio.**
>
> A entrega final deve contemplar todas as rotas e experiências existentes descritas neste documento.

---

## 1. OBJETIVO

O produto deve parecer um sistema corporativo maduro, utilizado diariamente por equipes de:

- TI;
- patrimônio;
- suporte;
- operações;
- administração;
- gestão.

A interface deve transmitir:

- controle;
- precisão;
- rastreabilidade;
- patrimônio;
- governança;
- operação;
- confiabilidade;
- densidade de informação.

O resultado NÃO deve parecer:

- SaaS genérico;
- dashboard de IA;
- template administrativo;
- Lovable;
- Shadcn default;
- landing page;
- CRUD estilizado;
- aplicação criada a partir de um único template.

### Princípio central

**Informação > decoração**

**Clareza > efeitos**

**Operação > apresentação**

**Composição > componentes**

**Identidade > tendência**

---

# 2. AUDITE O PROJETO ANTES DE ALTERAR

Primeiro percorra o projeto real.

Analise:

1. `frontend/src`;
2. todas as rotas;
3. componentes;
4. hooks;
5. serviços;
6. tipos;
7. chamadas API;
8. React Query;
9. autenticação;
10. permissões;
11. sistema de tema;
12. componentes compartilhados;
13. páginas existentes;
14. estilos globais;
15. arquivos de configuração.

Use o código existente como fonte da verdade para:

- rotas;
- endpoints;
- contratos;
- campos;
- tipos;
- regras;
- permissões;
- comportamento;
- integrações.

Não invente uma arquitetura paralela sem necessidade.

---

# 3. FRONTEND ONLY

## NÃO MODIFICAR O BACKEND

Não alterar:

- FastAPI;
- routers;
- endpoints;
- banco;
- MySQL;
- migrations;
- modelos backend;
- autenticação backend;
- regras de negócio backend;
- contratos da API.

Se algo puder ser resolvido no frontend, resolva no frontend.

Não criar API fake para substituir a existente.

Não remover funcionalidades para facilitar o redesign.

---

# 4. LOGIN — INTOCÁVEL

A rota:

`/login`

e os elementos relacionados:

- `LoginPage`;
- `ThreeLogoCanvas`;
- `LogoFallback`;
- `logo.glb`;

devem permanecer visualmente preservados.

Não:

- redesenhar;
- substituir o logo;
- remover a animação 3D;
- adicionar efeitos;
- adicionar scroll vertical;
- adicionar scroll horizontal.

O login é uma exceção visual e funcional.

---

# 5. REGRA MAIS IMPORTANTE — 17 EXPERIÊNCIAS DIFERENTES

O Design System é compartilhado.

O layout NÃO é.

Cada página deve possuir:

- protagonista próprio;
- wireframe próprio;
- distribuição espacial própria;
- hierarquia própria;
- ritmo próprio;
- densidade própria;
- composição adequada à operação.

### Regra absoluta

Se duas páginas possuem praticamente o mesmo wireframe, uma delas está errada.

Não fazer:

```text
HEADER
CARDS
TABELA
CARDS
```

em todas as páginas.

---

# 6. NÃO COMECE PELOS COMPONENTES

Antes de escrever JSX:

1. defina o problema operacional;
2. defina a informação protagonista;
3. defina a ação principal;
4. defina a informação persistente;
5. defina a hierarquia;
6. desenhe mentalmente o wireframe;
7. escolha a composição;
8. somente depois escolha os componentes.

### Ordem obrigatória

**Problema → hierarquia → wireframe → composição → componentes → estilo**

Nunca:

**Card → Card → Card → Dashboard genérico**

---

# 7. TESTE DOS 5 SEGUNDOS

Cada página deve ser identificável em aproximadamente 5 segundos sem depender apenas do título.

O usuário deve perceber:

- Estoque;
- Empréstimo;
- Histórico;
- Relatório;
- Unidades;
- Usuários;

pela própria composição.

---

# 8. TESTE EM PRETO E BRANCO

Remova mentalmente:

- cores;
- ícones;
- textos;
- badges.

A estrutura ainda precisa diferenciar cada página.

Se todas parecem iguais:

**REFaÇA.**

---

# 9. TESTE SEM EFEITOS

Remova mentalmente:

- sombras;
- gradientes;
- animações;
- glow;
- efeitos.

A interface ainda deve parecer sofisticada.

A sofisticação vem de:

- tipografia;
- grid;
- alinhamento;
- proporção;
- densidade;
- espaço;
- contraste;
- informação.

---

# 10. DIREÇÃO DE COR — MONOCROMÁTICA

## REGRA ABSOLUTA

A interface deve ser predominantemente:

**PRETO + BRANCO + CINZAS.**

Não utilizar cores típicas de interfaces SaaS/IA como identidade visual.

### NÃO usar como cores de identidade:

- azul;
- azul-marinho;
- navy;
- roxo;
- violeta;
- lilás;
- ciano;
- turquesa;
- verde;
- laranja;
- terracota;
- marrom;
- rosa;
- amarelo;
- neon;
- fluorescentes;
- gradientes multicoloridos.

### A identidade NÃO deve depender de uma cor de accent.

O contraste deve fazer esse papel.

---

# 11. LIGHT MODE

Paleta:

```text
Canvas:          #F7F7F5
Surface:         #FFFFFF
Surface Alt:     #F0F0ED

Border:          #D9D9D4
Border Strong:   #BDBDB7

Text Primary:    #111111
Text Secondary:  #5F5F5A
Text Muted:      #85857E

Black:           #000000
White:           #FFFFFF
```

O Light Mode deve parecer:

- editorial;
- limpo;
- técnico;
- sóbrio;
- corporativo.

Não deve parecer um dashboard branco genérico.

---

# 12. DARK MODE

Paleta:

```text
Canvas:          #050505
Surface:         #0C0C0C
Surface Alt:     #131313

Border:          #272727
Border Strong:   #3A3A3A

Text Primary:    #F2F2F0
Text Secondary:  #A4A4A0
Text Muted:      #70706C

Black:           #000000
White:           #FFFFFF
```

O Dark Mode deve ser realmente preto.

Não usar azul-marinho como substituto.

Não usar “dark blue”.

---

# 13. CONTRASTE É A IDENTIDADE

A hierarquia deve vir de:

- preto sobre branco;
- branco sobre preto;
- cinza sobre branco;
- cinza sobre preto;
- linhas;
- divisores;
- peso tipográfico;
- tamanho;
- espaço;
- preenchimento;
- inversão de superfície.

Não utilizar cores apenas para chamar atenção.

---

# 14. BOTÕES

### Light — primário

```text
background: #111111
text: #FFFFFF
```

### Dark — primário

```text
background: #F2F2F0
text: #080808
```

Secundários:

- transparentes;
- borda;
- texto monocromático.

Não usar botão azul, verde, roxo ou amarelo.

---

# 15. ESTADOS SEM DEPENDER DE COR

Não fazer:

```text
verde = ativo
vermelho = erro
amarelo = atenção
azul = informação
```

Utilizar:

- texto;
- símbolos;
- bordas;
- contraste;
- peso;
- preenchimento;
- posição.

Exemplos:

```text
● ATIVO
○ INATIVO
! ATENÇÃO
× ERRO
✓ CONCLUÍDO
```

A interface deve continuar compreensível em escala de cinza.

---

# 16. GRÁFICOS

Gráficos são funcionais, não decorativos.

Usar principalmente:

- preto;
- branco;
- cinzas;
- opacidade;
- diferentes espessuras;
- linhas sólidas;
- linhas tracejadas;
- pontos;
- áreas monocromáticas.

Não criar gráficos arco-íris.

Os gráficos precisam responder perguntas reais.

---

# 17. TIPOGRAFIA

### Interface

Plus Jakarta Sans.

### Dados técnicos

IBM Plex Mono.

Usar IBM Plex Mono para:

- patrimônio;
- serial;
- códigos;
- identificadores;
- números técnicos;
- datas;
- dados tabulares quando apropriado.

Escala:

```text
12
13
14
16
18
22
28
```

Títulos principais:

aproximadamente 28–32px no máximo.

Não criar títulos gigantes.

---

# 18. BORDER RADIUS

A interface não deve ter aparência “rounded SaaS”.

Preferir:

```text
2px
4px
6px
8px
```

Painéis:

máximo aproximado de 6px quando possível.

Inputs:

aproximadamente 4px.

Evitar:

```text
rounded-xl
rounded-2xl
rounded-3xl
rounded-full
```

Não transformar tudo em pill.

---

# 19. PROIBIÇÕES VISUAIS

Não usar:

- glassmorphism;
- backdrop blur decorativo;
- gradientes;
- glow;
- neon;
- sombras coloridas;
- backgrounds com grids decorativos;
- blobs;
- círculos decorativos;
- formas abstratas;
- excesso de sombras;
- excesso de cards;
- KPI grid repetitivo;
- pills em excesso;
- estética de landing page;
- estética Lovable;
- estética SaaS;
- estética AI dashboard;
- estética Shadcn default.

---

# 20. NÃO TRANSFORMAR TUDO EM CARD

Cards são componentes, não estrutura universal.

Preferir:

- tabelas;
- listas;
- fieldsets;
- divisores;
- painéis;
- drawers;
- timelines;
- árvores;
- documentos;
- toolbars;
- áreas de trabalho;
- grids;
- split views;
- workflows.

---

# 21. APP SHELL

Reconstruir o Shell para parecer produto corporativo.

## Sidebar

### Principal

- Dashboard

### Inventário

- Estoque
- Cadastro
- Histórico

### Operações

- Empréstimos
- Devoluções
- Periféricos
- Offboarding

### Gestão

- Usuários
- Unidades
- Relatórios
- Indicadores

Ações contextuais como:

- Termos;
- Vincular;
- Remover;
- Editar;

não precisam ser itens permanentes da sidebar.

---

# 22. PERIFÉRICOS

“Vincular Periféricos” não deve ser uma seção principal da navegação.

Usar:

**Periféricos**

como entrada principal.

A ação de vincular deve existir dentro do fluxo de periféricos ou em ações contextuais.

---

# 23. EXPERIÊNCIA 01 — DASHBOARD `/`

Protagonista:

**estado operacional do patrimônio.**

Composição:

- assimétrica;
- editorial;
- hierárquica.

Pode conter:

- resumo operacional;
- números relevantes;
- tendência;
- atenção;
- atividade recente;
- distribuição patrimonial;
- timeline compacta.

Não usar:

```text
KPI KPI KPI KPI
GRAPH GRAPH
CARD CARD CARD
```

O Dashboard deve responder:

> Como está o patrimônio de TI agora e onde preciso prestar atenção?

---

# 24. EXPERIÊNCIA 02 — ESTOQUE `/stock`

Protagonista:

**data grid operacional.**

A tela deve parecer uma ferramenta usada diariamente.

Estrutura:

```text
CONTEXTO
↓
TOOLBAR OPERACIONAL
↓
FILTROS
↓
DATA GRID
↓
PAGINAÇÃO
```

Recursos:

- busca;
- filtros;
- ordenação;
- seleção múltipla;
- ações contextuais;
- status;
- patrimônio;
- equipamento;
- categoria;
- unidade;
- responsável;
- situação;
- dados técnicos.

Paginação:

**7 itens por página**, preservando o comportamento esperado do projeto.

Dados técnicos com IBM Plex Mono.

Ao selecionar um ativo:

abrir `AssetDetailsPanel` lateral.

Não criar modal gigante.

---

# 25. EXPERIÊNCIA 03 — CADASTRO `/register`

Protagonista:

**formulário estruturado.**

Não usar formulário genérico em Card.

Organizar em:

- identificação;
- patrimônio;
- equipamento;
- características;
- localização;
- responsabilidade;
- informações complementares.

Composição:

```text
FORMULÁRIO PRINCIPAL       CONTEXTO
70%                        30%
```

A coluna lateral pode conter:

- resumo;
- regras;
- estado;
- informações contextuais;
- próximos passos.

---

# 26. EXPERIÊNCIA 04 — EDIÇÃO `/edit/:id`

A edição deve ser visualmente diferente do cadastro.

Protagonista:

**mudança.**

Mostrar:

```text
ESTADO ATUAL       ALTERAÇÕES
────────────       ──────────
valor antigo       novo valor
valor antigo       novo valor
valor antigo       novo valor
```

Destacar:

- campos alterados;
- valores anteriores;
- novos valores;
- impacto.

Não simplesmente reutilizar a tela de cadastro.

---

# 27. EXPERIÊNCIA 05 — REMOÇÃO `/remove`

Protagonista:

**decisão irreversível.**

Usar composição figure/ground.

Estrutura:

```text
ATIVO
↓
CONSEQUÊNCIA
↓
CONFIRMAÇÃO
↓
AÇÃO
```

Não transformar em CRUD.

A confirmação deve ocorrer em duas etapas.

Não depender de vermelho para transmitir perigo.

---

# 28. EXPERIÊNCIA 06 — EMPRÉSTIMOS `/loan`

Protagonista:

**fluxo operacional.**

Stepper de quatro fases:

```text
1. Equipamento
2. Usuário
3. Condições
4. Confirmação
```

O contexto do ativo permanece visível.

Não criar wizard genérico centralizado.

---

# 29. EXPERIÊNCIA 07 — DEVOLUÇÕES `/return`

Protagonista:

**conferência física.**

Mostrar:

- equipamento;
- responsável;
- checklist;
- periféricos;
- estado físico;
- observações;
- documento/termo assinado;
- confirmação.

Deve parecer uma operação de conferência de patrimônio.

---

# 30. EXPERIÊNCIA 08 — TERMOS `/terms`

Protagonista:

**documento.**

Não usar cards.

Estrutura:

```text
PENDENTES / ASSINADOS
        |
        |
   DOCUMENTO
        |
   VISUALIZAÇÃO
```

A área principal deve parecer um documento corporativo.

Painel lateral:

- seleção;
- status;
- informações;
- ações.

---

# 31. EXPERIÊNCIA 09 — PERIFÉRICOS `/peripherals`

Protagonista:

**relação entre ativos.**

Usar Master/Detail.

Esquerda:

lista.

Direita:

detalhes.

Mostrar:

- tipo;
- patrimônio;
- serial;
- status;
- equipamento relacionado;
- usuário;
- unidade;
- histórico.

---

# 32. EXPERIÊNCIA 10 — VINCULAR `/link`

Protagonista:

**associação.**

Fluxo:

```text
SELECIONAR ATIVO
↓
SELECIONAR PERIFÉRICO
↓
CONFERIR RELAÇÃO
↓
CONFIRMAR
```

Mostrar claramente:

```text
ATIVO
  ↕
PERIFÉRICO
```

Não tratar como CRUD.

---

# 33. EXPERIÊNCIA 11 — INDICADORES `/charts`

Protagonista:

**análise comparativa.**

Não copiar a estrutura do Dashboard.

Usar filtros funcionais.

Gráficos devem responder perguntas reais sobre:

- distribuição;
- evolução;
- situação;
- unidades;
- categorias;
- movimentações.

Usar Recharts funcionalmente.

---

# 34. EXPERIÊNCIA 12 — HISTÓRICO `/history`

Protagonista:

**tempo.**

Criar timeline editorial.

Estrutura:

```text
FILTROS       EVENTOS
30%           70%
```

Cada evento pode mostrar:

- data;
- hora;
- ação;
- ativo;
- usuário;
- unidade;
- contexto.

Não usar tabela CRUD genérica.

---

# 35. EXPERIÊNCIA 13 — RELATÓRIOS `/report`

Protagonista:

**configuração + resultado.**

Estrutura:

```text
CONFIGURAÇÃO
────────────
parâmetros

PREVIEW
────────────
relatório / tabela
```

Ações:

- gerar;
- exportar;
- filtrar;
- revisar.

A prévia deve ser densa.

---

# 36. EXPERIÊNCIA 14 — UNIDADES `/unidades`

Protagonista:

**hierarquia organizacional.**

Usar árvore.

Mostrar:

- unidades;
- filiais;
- distribuição;
- densidade patrimonial;
- equipamentos;
- responsáveis.

Edição através de drawer contextual.

Não transformar em tabela administrativa genérica.

---

# 37. EXPERIÊNCIA 15 — USUÁRIOS `/users`

Protagonista:

**governança.**

Mostrar:

- usuários;
- funções;
- permissões;
- responsabilidades;
- ativos associados.

Criar matriz de permissões quando suportada pela estrutura existente.

Não tratar como simples CRUD.

---

# 38. EXPERIÊNCIA 16 — OFFBOARDING `/offboarding-ti`

Protagonista:

**checklist operacional.**

Estrutura:

```text
DESLIGAMENTO
↓
EQUIPAMENTOS
↓
PERIFÉRICOS
↓
ACESSOS / CONFERÊNCIA
↓
DEVOLUÇÃO
↓
FINALIZAÇÃO
```

Estados:

- pendente;
- em andamento;
- concluído;
- bloqueado.

Não criar dashboard.

---

# 39. EXPERIÊNCIA 17 — ASSET DETAIL PANEL

Protagonista:

**ficha técnica do ativo.**

Painel lateral.

Seções:

```text
IDENTIFICAÇÃO
PATRIMÔNIO
HARDWARE
PERIFÉRICOS
RESPONSÁVEL
LOCALIZAÇÃO
HISTÓRICO
TERMOS
```

Deve funcionar principalmente a partir do Estoque.

Não criar modal gigante.

---

# 40. COMPONENTES COMPARTILHADOS

Reutilizar componentes funcionais:

- Button;
- Input;
- Select;
- Dialog;
- Drawer;
- Table;
- Tooltip;
- Command;
- Tabs;
- etc.

Mas:

## REUTILIZAR COMPONENTES ≠ REUTILIZAR COMPOSIÇÃO

A mesma tabela pode existir em vários contextos.

A mesma estrutura de página não.

---

# 41. PAGE HEADERS

Não usar o mesmo cabeçalho em todas as páginas.

Evitar:

```text
Título
Descrição
Botão
```

universalmente.

Cada entrada deve refletir a função.

Exemplos:

- Estoque → toolbar operacional;
- Empréstimo → etapa do fluxo;
- Histórico → período + filtros;
- Termos → documento + status;
- Relatórios → configuração + preview;
- Unidades → hierarquia;
- Periféricos → seleção master/detail.

---

# 42. ESTADOS OBRIGATÓRIOS

Implementar e revisar:

- loading;
- skeleton;
- empty;
- error;
- success;
- destructive;
- disabled;
- hover;
- focus;
- active;
- selected.

Não implementar somente o estado feliz.

---

# 43. RESPONSIVIDADE

Desktop é prioridade.

Mobile não deve ser apenas:

```text
overflow-x-auto
```

Quando necessário:

- Sidebar → Sheet;
- filtros → Drawer;
- tabelas → cards compactos;
- detalhes → Drawer;
- formulários → uma coluna;
- workflows → vertical.

---

# 44. MOTION

Microinterações:

`120–200ms`

Overlays:

`200–300ms`

Evitar:

- bounce;
- parallax;
- stagger decorativo;
- efeitos chamativos.

---

# 45. REFERÊNCIAS

Use referências para estudar comportamento e organização:

### Produtos

- Linear
- Vercel Dashboard
- Notion
- Figma

### Enterprise / ITAM

- ServiceNow ITAM
- Snipe-IT
- Freshservice
- ManageEngine AssetExplorer
- Lansweeper
- Asset Panda
- InvGate
- Reftab
- GoCodes
- GLPI
- EZOfficeInventory
- AssetTiger

### Componentes

- shadcn/ui
- Radix UI

## NÃO COPIAR

Não copiar:

- layouts;
- cores;
- logos;
- identidade;
- páginas;
- componentes inteiros.

Extrair apenas:

- comportamento;
- densidade;
- organização;
- hierarquia;
- composição;
- padrões operacionais.

---

# 46. REGRA ANTI-IA

Para cada tela, pergunte:

> Isso poderia estar em qualquer SaaS?

Se sim:

**REFAÇA.**

Pergunte:

> Isso parece um sistema usado diariamente por uma equipe de patrimônio/TI?

Se não:

**REFAÇA.**

Pergunte:

> Essa tela possui identidade própria?

Se não:

**REFAÇA.**

---

# 47. MATRIZ DE DIVERSIDADE

| Rota | Estrutura obrigatória |
|---|---|
| `/` | Dashboard executivo assimétrico |
| `/stock` | Data grid operacional |
| `/register` | Formulário editorial |
| `/edit/:id` | Comparação / diff |
| `/remove` | Figure / ground |
| `/loan` | Stepper operacional |
| `/return` | Workflow de conferência |
| `/terms` | Documento |
| `/peripherals` | Master / Detail |
| `/link` | Associação progressiva |
| `/charts` | Análise comparativa |
| `/history` | Timeline |
| `/report` | Configuração + preview |
| `/unidades` | Árvore organizacional |
| `/users` | Governança / permissões |
| `/offboarding-ti` | Checklist operacional |
| `AssetDetailsPanel` | Ficha técnica lateral |

Essa matriz é uma **restrição de implementação**, não apenas inspiração.

---

# 48. NÃO USAR A TELA ANTERIOR COMO BASE VISUAL

Ao construir uma nova rota:

NÃO copie a estrutura da rota anterior apenas porque ela já existe.

NÃO copie:

- posição do título;
- posição dos botões;
- quantidade de cards;
- estrutura do conteúdo;
- distribuição de colunas;
- mesmo header;
- mesmo footer;
- mesma toolbar.

O componente pode ser reutilizado.

A composição deve ser decidida pela função da página.

---

# 49. ORDEM DE IMPLEMENTAÇÃO

Execute tudo em uma única tarefa.

## Fase 1 — Auditoria

Mapear o projeto.

## Fase 2 — Fundação visual

Reconstruir:

- tokens;
- cores;
- tipografia;
- radius;
- borders;
- spacing;
- estados;
- tema.

## Fase 3 — App Shell

Reconstruir:

- sidebar;
- navegação;
- mobile;
- command palette;
- área de conta.

## Fase 4 — Dashboard

Implementar `/`.

## Fase 5 — Estoque

Implementar `/stock` completamente.

## Fase 6 — Operações

Implementar:

- `/register`;
- `/edit/:id`;
- `/remove`;
- `/loan`;
- `/return`;
- `/terms`;
- `/peripherals`;
- `/link`;
- `/offboarding-ti`.

## Fase 7 — Gestão

Implementar:

- `/charts`;
- `/history`;
- `/report`;
- `/unidades`;
- `/users`.

## Fase 8 — Asset Details

Finalizar `AssetDetailsPanel`.

## Fase 9 — Estados

Revisar todas as rotas.

## Fase 10 — Responsividade

Revisar desktop/tablet/mobile.

## Fase 11 — QA

Executar build, lint e auditoria visual.

---

# 50. NÃO PARAR APÓS DUAS TELAS

É expressamente proibido terminar a tarefa depois de:

- Dashboard;
- Estoque.

A implementação só está concluída quando todas as experiências estiverem revisadas.

Não entregar:

> “primeira versão”

> “protótipo”

> “base para continuar”

> “as outras páginas podem ser feitas depois”

A tarefa é o redesign completo.

---

# 51. NÃO SIMPLIFICAR O SISTEMA

Não:

- apagar funcionalidades;
- remover filtros;
- remover ações;
- remover permissões;
- trocar API por mock;
- substituir dados reais por placeholders;
- remover React Query;
- simplificar fluxos para facilitar CSS.

O redesign deve trabalhar sobre a aplicação real.

---

# 52. QUALIDADE DO CÓDIGO

Manter:

- React;
- TypeScript;
- Vite;
- TanStack Router;
- TanStack Query;
- Tailwind;
- Radix;
- Shadcn;
- React Hook Form;
- Zod;
- Recharts;
- Lucide.

Evitar:

- duplicação;
- hacks;
- CSS desnecessário;
- `!important` sem justificativa;
- inline styles indiscriminados;
- componentes gigantes novos;
- lógica de negócio na apresentação.

---

# 53. AUDITORIA DE CLASSES / ESTILOS

Depois da implementação, procurar por padrões como:

```text
rounded-xl
rounded-2xl
rounded-3xl
rounded-full
backdrop-blur
bg-gradient
gradient
glow
shadow-[color
bg-grid
glass
```

Revisar todas as ocorrências.

Remover qualquer elemento que contradiga a direção monocromática e editorial.

---

# 54. QA VISUAL

Revisar cada rota individualmente.

Para cada uma:

### 5 segundos

É reconhecível?

### Preto e branco

A composição continua diferente?

### Sem efeitos

Continua sofisticada?

### Sem logo

Ainda parece parte do mesmo produto?

### Sem cor

A hierarquia continua clara?

### Sem cards

A página ainda possui uma boa composição?

---

# 55. CHECKLIST FINAL

## Arquitetura

- [ ] todas as rotas funcionam;
- [ ] navegação funciona;
- [ ] permissões continuam funcionando;
- [ ] APIs continuam funcionando;
- [ ] backend não foi alterado.

## Visual

- [ ] monocromático;
- [ ] preto e branco presentes;
- [ ] sem azul;
- [ ] sem roxo;
- [ ] sem verde;
- [ ] sem laranja;
- [ ] sem amarelo;
- [ ] sem terracota;
- [ ] sem neon;
- [ ] sem gradientes;
- [ ] sem glow;
- [ ] sem glassmorphism;
- [ ] sem SaaS genérico;
- [ ] sem estética AI;
- [ ] sem Shadcn genérico;
- [ ] sem excesso de cards;
- [ ] sem excesso de radius.

## Diferenciação

- [ ] Dashboard assimétrico;
- [ ] Estoque data grid;
- [ ] Cadastro editorial;
- [ ] Edição comparação;
- [ ] Remoção decisão;
- [ ] Empréstimo workflow;
- [ ] Devolução conferência;
- [ ] Termos documento;
- [ ] Periféricos master/detail;
- [ ] Link associação;
- [ ] Indicadores análise;
- [ ] Histórico timeline;
- [ ] Relatório config + preview;
- [ ] Unidades árvore;
- [ ] Usuários governança;
- [ ] Offboarding checklist;
- [ ] Asset Details ficha lateral.

## Login

- [ ] Login preservado;
- [ ] logo 3D preservado;
- [ ] animação preservada;
- [ ] sem scroll.

## Estados

- [ ] loading;
- [ ] skeleton;
- [ ] empty;
- [ ] error;
- [ ] success;
- [ ] destructive;
- [ ] disabled;
- [ ] hover;
- [ ] focus;
- [ ] active;
- [ ] selected.

## Responsividade

- [ ] desktop;
- [ ] tablet;
- [ ] mobile;
- [ ] sidebar mobile;
- [ ] filtros mobile;
- [ ] tabelas adaptadas;
- [ ] drawers adaptados.

## Qualidade

- [ ] `npm run build`;
- [ ] `npm run lint`;
- [ ] console revisado;
- [ ] todas as rotas testadas;
- [ ] light mode revisado;
- [ ] dark mode revisado;
- [ ] auditoria anti-template realizada.

---

# 56. PRINCÍPIO FINAL

Não tente fazer o sistema parecer “bonito”.

Faça parecer:

**PRECISO.**

**CORPORATIVO.**

**OPERACIONAL.**

**MADURO.**

**CONFIÁVEL.**

**REAL.**

A interface não deve impressionar através de efeitos.

Ela deve impressionar através de:

**informação + estrutura + hierarquia + densidade + contraste + precisão.**

O produto final deve parecer:

> **um software corporativo real de Gestão de Patrimônio / IT Asset Management.**

Não uma coleção de telas geradas por IA.

Não um template.

Não um SaaS genérico.

---

# RESULTADO ESPERADO

Uma única aplicação coerente chamada:

**REVALLE — CONTROLE DE PATRIMÔNIO**

com:

**1 identidade visual monocromática**

+

**17 experiências visualmente distintas**

+

**fluxos operacionais reais**

+

**alta densidade de informação**

+

**preto e branco como identidade**

+

**zero aparência de template SaaS/IA**

+

**frontend preservando integralmente a lógica e o backend existentes.**

## EXECUTE O REDESIGN COMPLETO AGORA.

Não entregue somente uma proposta.

Não entregue somente imagens.

Não pare após duas telas.

Não faça mockup no lugar da implementação.

**Implemente o frontend completo.**
