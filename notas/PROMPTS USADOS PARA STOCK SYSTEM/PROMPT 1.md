# PROMPT MASTER — REDESIGN FRONTEND STOCK SYSTEM (GOOGLE ANTIGRAVITY)

VOCÊ ESTÁ TRABALHANDO SOBRE UM PROJETO EXISTENTE. NÃO COMECE CRIANDO UM NOVO SISTEMA. PRIMEIRO ANALISE O CÓDIGO REAL.

O BACKEND JÁ ESTÁ FUNCIONANDO. NÃO TRATE O BACKEND COMO UM PROBLEMA A SER RESOLVIDO DURANTE O REDESIGN.

QR CODE NÃO FAZ PARTE DO PRODUTO E NÃO DEVE SER IMPLEMENTADO EM NENHUMA HIPÓTESE.

Você atuará como um **Engenheiro de Software Sênior / Frontend Architect / UX-UI Senior** trabalhando sobre um produto SaaS corporativo real, existente e em produção/desenvolvimento ativo. Você NÃO está prototipando, NÃO está criando um projeto acadêmico e NÃO está reconstruindo o backend.

---

## 1. CONTEXTO DO PROJETO

- **Nome do projeto:** Sistema corporativo de Controle de Patrimônio / Estoque / Ativos.
- **Repositório:** https://github.com/enzononato/stock_system
- **Branch de trabalho:** `redesign-frontend`
- **Diretório local:** `E:\STOCK SYSTEM`
- **Estado:** projeto real, em desenvolvimento, com backend funcional e módulos já implementados (Estoque, Periféricos, Relatório Mensal, exportação CSV, barra lateral de navegação).
- **Ferramenta de execução:** você (Antigravity) tem acesso direto ao código-fonte real e deve operar sobre ele.

O objetivo é uma evolução profunda de UX/UI, arquitetura de interface e consistência visual, **mantendo 100% das funcionalidades existentes operacionais**.

---

## 2. STACK TÉCNICA CONHECIDA

**Frontend:** React + TypeScript + Vite
**Backend:** Python + FastAPI + Uvicorn (JÁ FUNCIONAL — não tratar como quebrado, não recriar, não mockar)
**Banco de dados:** MySQL (real — não substituir, não mockar)
**Autenticação:** JWT (já implementada — não reconstruir)

**PROIBIDO nesta etapa:**
- criar backend fake ou mock permanente;
- criar API fake;
- substituir o banco de dados;
- recriar a autenticação;
- "resolver" a conexão MySQL como se estivesse quebrada;
- reconstruir endpoints existentes sem motivo comprovado por auditoria.

---

## 3. REGRA ABSOLUTA — ANALISAR ANTES DE ALTERAR

Ordem obrigatória de trabalho: **ENTENDER → PLANEJAR → IMPLEMENTAR**. Nunca alterar código "porque parece melhor".

Antes de qualquer alteração, faça uma auditoria completa e mapeie:

**Frontend:** estrutura de diretórios, páginas, rotas, layouts, componentes, componentes reutilizáveis, hooks, contextos, providers, serviços, chamadas HTTP, gerenciamento de estado, autenticação, guards, formulários, tabelas, filtros, paginação, modais, dropdowns, notificações, loading/error/empty states, responsividade.

**Backend:** estrutura, routers, endpoints, schemas, models, serviços, autenticação, autorização, regras de negócio, integrações, contratos de API.

**Banco:** como o frontend consome os dados reais, sem presumir estruturas que não existem.

---

## 4. MATRIZ DE FUNCIONALIDADES (OBRIGATÓRIA ANTES DO REDESIGN)

Para cada funcionalidade encontrada, documentar: nome, página, rota, componente, endpoint, método HTTP, dados recebidos/enviados, permissões, dependências, estado atual, problemas encontrados, risco de alteração.

Classificar cada uma como: **funcionando** / **parcialmente funcionando** / **com problema** / **somente visual** / **dependente de backend** / **dependente de dados** / **desconhecida**.

Módulos já conhecidos que devem estar nessa matriz (confirmar existência real no código, sem presumir que precisam ser criados):
- Estoque / Patrimônio
- Periféricos (e "Vincular Periféricos" — hierarquicamente dentro de Periféricos, que por sua vez pertence a "Operações")
- Relatório Mensal (com filtro de revenda)
- Exportação CSV (padrão de acentuação/formatação ABNT2)
- Empréstimos, Devoluções, Transferências, Baixa
- Inventário, Unidades, Usuários, Permissões, Histórico, Configurações
- Demais módulos existentes no código que ainda não estejam listados aqui

NÃO remover nenhuma funcionalidade existente só porque não está citada explicitamente neste prompt.

---

## 5. PRINCÍPIO FUNDAMENTAL

**Preservar sempre:** funcionalidades, regras de negócio, APIs, contratos, autenticação, autorização, permissões, integração com backend, banco de dados, fluxos existentes, validações, comportamento funcional, dados reais.

**Modificar principalmente:** UI, UX, layout, hierarquia visual, navegação, componentes, espaçamentos, tipografia, responsividade, estados visuais, feedbacks, consistência, acessibilidade.

Se uma mudança visual exigir mudança funcional, analisar com cautela antes de agir — nunca por padrão.

---

## 6. PROIBIDO USAR DADOS FALSOS

Nunca inventar números, usuários, patrimônio, movimentações, gráficos, estatísticas ou relatórios fictícios. Use exclusivamente dados reais retornados pelo backend. Quando não houver dados suficientes, exibir um empty state apropriado, por exemplo: *"Não há dados suficientes para exibir este gráfico."*

---

## 7. DASHBOARD

Tratar como área de gestão real, usando somente métricas comprovadamente existentes no backend (descobrir isso na auditoria, não presumir). Possíveis categorias, apenas se suportadas por dados reais: total de ativos, disponíveis, em uso, empréstimos, devoluções, movimentações, estoque, unidades, setores, usuários, status, distribuição por categoria, tendências, alertas.

Nenhum gráfico ou card deve existir só para preencher espaço — cada elemento precisa de função administrativa real.

---

## 8. ESCOPO FUNCIONAL

Manter e apresentar profissionalmente todas as funcionalidades reais já existentes relacionadas a patrimônio, estoque, equipamentos, periféricos, movimentações, empréstimos, devoluções, transferências, baixa, inventário, unidades, usuários, permissões, histórico, relatórios, configurações e operações administrativas. Descubra no código quais desses módulos realmente existem antes de assumir que precisam ser criados.

---

## 9. QR CODE — TOTALMENTE FORA DE ESCOPO

QR Code **não faz parte do produto**. Não criar geração, leitura, scanner, botão, campo, página, modal, impressão, fluxo, integração, componente ou biblioteca de QR Code. Não mencionar QR Code como melhoria futura, backlog, sugestão, feature ou roadmap, em nenhuma hipótese. Se encontrar código de QR Code já existente no projeto durante a auditoria, apenas identifique-o e avalie remoção somente se isso não quebrar nada mais — nunca introduza QR Code em qualquer parte nova do sistema.

---

## 10. DIREÇÃO VISUAL

Identidade: moderna, corporativa, profissional, elegante, tecnológica, limpa, sofisticada, funcional e consistente.

Usar como **referência conceitual de qualidade** (nunca copiar layout, cores, logotipo ou textos): InvGate, Reftab, Asset Panda, GoCodes, Snipe-IT, Freshservice, ServiceNow ITAM, ManageEngine AssetExplorer, Lansweeper, EZOfficeInventory, AssetTiger, GLPI.

Criar identidade visual própria, inspirada apenas em padrões de qualidade, hierarquia, densidade de informação e organização desses produtos.

---

## 11. EXPERIÊNCIA DE USO

Priorizar produtividade, clareza, velocidade, escaneabilidade, redução de cliques, hierarquia, previsibilidade, feedback visual e consistência. É um software corporativo de uso intenso — não uma landing page. Evitar excesso de animações, gradientes, efeitos, sombras, elementos decorativos e cards desnecessários.

---

## 12. SIDEBAR / NAVEGAÇÃO

Elemento central da experiência. Deve ter hierarquia clara, agrupamento lógico, estados ativo/inativo, expansão/recolhimento, ícones consistentes e comportamento responsivo.

**Hierarquia obrigatória a respeitar (já definida no projeto):**
- "Periféricos" pertence a "Operações".
- "Vincular Periféricos" pertence a "Periféricos" (não é módulo principal separado).

Organizar perfil, tema, logout e configurações de forma profissional dentro da navegação, sem elementos visuais desnecessários. Considerar também um modelo de navegação em árvore quando fizer sentido (ex.: Inventário > Estoque > Cadastrar Item), desde que preserve a hierarquia real do sistema.

---

## 13. LAYOUT

Estrutura de aplicação SaaS profissional: Sidebar, Header, Page Header, Breadcrumb (só quando útil), conteúdo principal, filtros, ações, tabelas, cards, modais, drawers, notificações, estados vazios. Nada deve existir só porque "sistemas SaaS costumam ter" — tudo precisa ter função.

---

## 14. TABELAS

Componentes críticos. Devem ter boa hierarquia, colunas equilibradas, espaçamento adequado, alinhamento consistente, cabeçalho claro, hover, seleção quando aplicável, ordenação, filtros, busca, paginação, controle de itens por página, loading, empty state, error state e responsividade.

**Ajuste específico já solicitado:** padronizar todas as tabelas para exibir um número fixo de itens por página (ex.: 7) com scroll lateral/vertical quando necessário, em vez de tabelas esmagadas ou colunas cortadas.

No mobile: usar cards, tabela horizontal com colunas prioritárias, ou outra estratégia adequada — nunca apenas reduzir tudo até ficar ilegível.

---

## 15. FORMULÁRIOS

Labels claras, agrupamento lógico, campos consistentes, validação, mensagens de erro, estados disabled/loading, feedback de sucesso, acessibilidade. Separar em seções quando necessário; evitar formulários gigantes sem hierarquia.

**Ajuste específico já solicitado:** a edição de itens no módulo Estoque deve ocorrer via pop-up/modal (confirmar comportamento atual do botão "Confirmar", que possui bug reportado, e corrigi-lo após identificar a causa raiz).

---

## 16. FILTROS E BUSCA

Incluir, quando aplicável: busca textual, filtros por status, unidade, categoria, tipo, revenda, período, filtros combinados, botão de limpar filtros e indicação visual de filtros ativos. Não esconder nem duplicar filtros importantes.

**Ajustes específicos já solicitados:**
- Adicionar filtros no módulo de Periféricos.
- Adicionar filtro de revenda no Relatório Mensal.

---

## 17. DETALHES DOS ATIVOS

Organizar de forma clara: identificação, descrição, categoria, status, localização, unidade, responsável, histórico, movimentações e informações adicionais — somente campos que existam de fato no backend.

---

## 18. MOVIMENTAÇÕES

Empréstimo, devolução, transferência, movimentação e baixa devem deixar claro: o que está acontecendo, qual ativo, origem, destino, responsável, status, data e consequência da ação. A interface deve reduzir erros operacionais.

---

## 19. HISTÓRICO

Apresentar eventos com data, ação, usuário, ativo, origem, destino e detalhes. Usar timeline ou tabela apenas quando fizer sentido com os dados reais disponíveis.

---

## 20. RELATÓRIOS

Sem gráficos decorativos. Melhorar filtros, organização, leitura, exportação, estados de carregamento e vazios, preservando as integrações existentes.

**Ajustes específicos já solicitados:**
- Gráficos em onda/linha onde fizer sentido com dados reais.
- Filtro de revenda.
- Corrigir acentuação/formatação da exportação CSV para o padrão ABNT2.

---

## 21. RESPONSIVIDADE

Funcionar profissionalmente em desktop, notebook, tablet e mobile, adaptando a experiência (não apenas "responsive por obrigação"):
- **Desktop:** aproveitar espaço, tabelas completas, navegação lateral.
- **Tablet:** reorganizar conteúdo mantendo ações acessíveis.
- **Mobile:** priorizar conteúdo, cards quando necessário, menus apropriados (ex.: sidebar em estilo hambúrguer, conforme já solicitado), botões acessíveis, evitar overflow horizontal desnecessário.

---

## 22. DESIGN SYSTEM

Definir e consolidar:
- **Cores:** background, surface, primary, secondary, success, warning, danger, info, texto, texto secundário, bordas.
- **Tipografia:** família, tamanhos, pesos, line-height, hierarquia.
- **Espaçamento:** escala consistente.
- **Border radius** e **shadows** padronizados, usados com moderação.
- **Componentes padrão:** Button, Input, Select, Checkbox, Switch, Badge, Card, Modal, Drawer, Dropdown, Tooltip, Tabs, Table, Pagination, Alert, Toast, Skeleton, Empty State.

Evitar múltiplas versões diferentes do mesmo componente.

---

## 23. ACESSIBILIDADE

Contraste, foco visível, navegação por teclado, aria-label, tamanho de área de clique adequado, semântica correta, mensagens de erro claras. Nunca usar cor como única forma de transmitir informação.

---

## 24. PERFORMANCE

Evitar renderizações desnecessárias, componentes pesados, imagens não otimizadas, bundle inchado. Usar lazy loading, evitar chamadas duplicadas, otimizar tabelas grandes, filtros e paginação. Não instalar bibliotecas sem necessidade comprovada.

---

## 25. SEGURANÇA

Nunca expor senhas, tokens, JWT secrets, credenciais, `.env` ou informações sensíveis. Não alterar autenticação sem necessidade, não colocar credenciais no frontend, não versionar arquivos de ambiente, não comprometer permissões existentes.

---

## 26. GIT

Branch de trabalho: `redesign-frontend`. Antes de iniciar: verificar branch atual, `git status`, alterações locais e commits recentes.

**PROIBIDO sem autorização explícita:** `git reset --hard`, `git clean -fd`, checkout destrutivo, apagar alterações do usuário, apagar stash, reescrever histórico.

Existe um stash local chamado `backup-local-antes-atualizacao` — **NÃO remover, aplicar, alterar ou descartar automaticamente**.

---

## 27. ARQUIVO .ENV

Não substituir, apagar, versionar, expor, imprimir secrets ou modificar arbitrariamente. Preservar a configuração local existente.

---

## 28. ESTRATÉGIA DE IMPLEMENTAÇÃO EM FASES

1. **Reconhecimento:** estrutura, código, rotas, APIs, componentes, funcionalidades.
2. **Auditoria:** problemas, inconsistências, duplicações, riscos, dívida visual/técnica, problemas de UX.
3. **Arquitetura visual:** layout, navegação, design tokens, componentes, hierarquia, responsividade.
4. **Shell da aplicação:** sidebar, header, navegação, layout, tema, responsividade.
5. **Páginas principais** (ordem final definida após auditoria): Dashboard, Estoque/Patrimônio, Operações, Periféricos, Empréstimos, Devoluções, Transferências, Baixas, Histórico, Relatórios, Usuários, Unidades, Configurações e demais páginas existentes.
6. **Componentes:** consolidar tabelas, filtros, formulários, modais, cards, badges, feedbacks.
7. **Responsividade:** testar em desktop, notebook, tablet, mobile.
8. **Qualidade:** rodar TypeScript, build, lint (se configurado), testes existentes, validar rotas, APIs, autenticação e principais fluxos.

---

## 29. NÃO QUEBRAR FUNCIONALIDADES

Após cada alteração relevante, verificar se rota, componente, chamada de API, exibição de dados, formulários, filtros, paginação, autenticação e permissões continuam funcionando. Uma página só está "pronta" se continuar funcional, não apenas visualmente melhor.

---

## 30. TRATAMENTO DE ESTADOS

- **Loading:** skeleton ou loading state apropriado.
- **Empty:** mensagem clara (ex.: "Não há equipamentos cadastrados."), nunca tabela vazia sem explicação.
- **Error:** mensagem compreensível.
- **Success:** confirmação adequada após operações.
- **Disabled:** deixar claro quando uma ação não está disponível.

---

## 31. MICROINTERAÇÕES

Usar apenas quando melhoram a experiência: hover, focus, transições, feedback, expansão/collapse, confirmação. Evitar excesso — o sistema deve parecer rápido.

---

## 32. MODAIS E DRAWERS

Usar quando melhoram o fluxo, não em todas as páginas. Garantir fechamento, tecla ESC, gerenciamento de foco, scroll, validação, loading e feedback.

---

## 33. ERROS DE UX A EVITAR

Dashboard cheio de cards inúteis; sidebar enorme; textos pequenos; botões sem hierarquia; tabelas apertadas; excesso de bordas, sombras, cores ou ícones; informações duplicadas; modais dentro de modais; formulários confusos; telas vazias sem orientação; dados falsos; funcionalidades fictícias; páginas desconectadas; componentes inconsistentes.

---

## 34. ARQUITETURA FRONTEND

Verificar durante a auditoria: componentes gigantes, lógica duplicada, chamadas de API espalhadas, tipos duplicados, hooks inconsistentes, estado mal distribuído, alto acoplamento, imports circulares, CSS/estilos inconsistentes ou duplicados. Melhorar quando necessário, priorizando mudanças seguras e incrementais — evitar refatoração massiva sem justificativa clara.

---

## 35. BIBLIOTECAS

Antes de adicionar qualquer dependência: verificar se já existe no projeto, avaliar se dá para resolver com o que já está disponível, justificar tecnicamente, avaliar impacto no bundle.

---

## 36. REGRA SOBRE FUNCIONALIDADES EXISTENTES

- Funcionalidade existe → **MANTER**.
- Funcionalidade existe mas está visualmente ruim → **REDESENHAR**.
- Existe duplicação → **CONSOLIDAR**.
- Existe código morto → **IDENTIFICAR antes de remover**.
- Existe funcionalidade quebrada → **CORRIGIR somente após entender a causa raiz**.
- Não tem certeza → **NÃO presumir, investigar**.

---

## 37. LIMITES DA AUTONOMIA

Autonomia total para decisões de layout, cores, tipografia, espaçamento, componentes, UX, responsividade e organização visual.

**Autonomia NÃO cobre:** quebrar APIs, remover funcionalidades, alterar regras de negócio, substituir banco de dados, criar dados falsos, modificar autenticação, ignorar permissões, introduzir QR Code, ou reconstruir o backend sem necessidade comprovada.

---

## 38. CRITÉRIO DE QUALIDADE FINAL

O resultado deve parecer um produto SaaS corporativo profissional — nunca um template, projeto universitário, dashboard genérico, protótipo, CRUD básico, ou "sistema antigo com nova cor". Deve transmitir confiança, organização, maturidade, precisão e qualidade empresarial.

---

## 39. CHECKLIST DE AUDITORIA FINAL

**Funcional**
- [ ] Login funcionando
- [ ] Autenticação funcionando
- [ ] Rotas funcionando
- [ ] APIs funcionando
- [ ] CRUDs funcionando
- [ ] Filtros funcionando
- [ ] Busca funcionando
- [ ] Paginação funcionando
- [ ] Formulários funcionando
- [ ] Permissões preservadas
- [ ] Dados reais sendo utilizados

**Visual**
- [ ] Design consistente
- [ ] Sidebar consistente
- [ ] Header consistente
- [ ] Tipografia consistente
- [ ] Espaçamento consistente
- [ ] Botões consistentes
- [ ] Tabelas consistentes
- [ ] Formulários consistentes
- [ ] Estados consistentes
- [ ] Responsividade adequada

**UX**
- [ ] Navegação intuitiva
- [ ] Hierarquia clara
- [ ] Feedbacks adequados
- [ ] Empty states
- [ ] Loading states
- [ ] Error states
- [ ] Confirmações
- [ ] Ações importantes facilmente encontradas

**Técnico**
- [ ] TypeScript sem erros
- [ ] Build funcionando
- [ ] Sem APIs fake
- [ ] Sem dados fake
- [ ] Sem secrets expostos
- [ ] Sem alterações destrutivas
- [ ] Sem dependências desnecessárias
- [ ] Sem QR Code
- [ ] Backend preservado

---

## 40. RELATÓRIO FINAL OBRIGATÓRIO

Ao concluir, apresente um relatório objetivo contendo:

1. **Auditoria** — arquitetura encontrada, principais módulos, rotas, APIs, componentes relevantes.
2. **Problemas encontrados** — separados por UX, UI, arquitetura, responsividade, performance, funcionalidade.
3. **Alterações realizadas** — listadas por página/módulo.
4. **Componentes criados ou modificados.**
5. **APIs preservadas** — quais integrações continuam funcionando.
6. **Funcionalidades preservadas** — confirmação dos principais fluxos.
7. **Validações** — comandos executados, resultados, erros encontrados e corrigidos. NÃO afirmar que algo foi testado se não foi realmente testado.
8. **Pendências** — listar apenas o que realmente ficou pendente.

---

## 41. RESUMO DE EXECUÇÃO

Você deve: analisar o projeto real → compreender a arquitetura → mapear funcionalidades, frontend, backend e APIs → preservar o funcionamento → criar estratégia de redesign → criar identidade visual própria → implementar progressivamente → usar dados reais → preservar autenticação, banco e APIs → melhorar UX/UI/responsividade → consolidar componentes → validar continuamente → documentar o resultado.

Aja como um engenheiro sênior trabalhando em um produto empresarial existente. NÃO como alguém criando um protótipo, uma landing page, um CRUD fictício, ou reconstruindo o backend.
