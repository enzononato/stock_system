# Core Logic Renew

REBUILD COMPLETO DO FRONTEND — PRESERVAR TODA A LÓGICA E RECONSTRUIR A EXPERIÊNCIA

MISSÃO

Quero que você reconstrua completamente o frontend deste projeto, mas utilizando o código atual como fonte de verdade para descobrir e preservar toda a lógica existente.

A interface atual NÃO precisa ser preservada.

O design atual NÃO precisa ser preservado.

A arquitetura visual atual NÃO precisa ser preservada.

Porém, toda a lógica, integração e funcionalidade importante existente no projeto precisa ser cuidadosamente identificada antes da reconstrução.

O objetivo é:

preservar o cérebro do sistema e reconstruir completamente a experiência.



1. REGRA MAIS IMPORTANTE

NÃO comece criando telas novas imediatamente.

Antes de modificar ou substituir o frontend, faça uma engenharia reversa completa do projeto existente.

Leia e analise o código.

Identifique tudo que já existe e que seja importante para o funcionamento do sistema.

Você deve procurar explicitamente por:

APIs;

endpoints;

URLs de backend;

chamadas HTTP;

fetch;

Axios;

React Query;

mutations;

queries;

hooks;

services;

repositories;

autenticação;

autorização;

tokens;

refresh token;

interceptors;

headers;

cookies;

localStorage;

sessionStorage;

variáveis de ambiente;

tipos TypeScript;

interfaces;

enums;

models;

schemas;

validações;

regras de negócio;

estados;

contextos;

providers;

stores;

gerenciamento global de estado;

rotas;

guards;

permissões;

upload de arquivos;

tratamento de erros;

paginação;

filtros;

busca;

ordenação;

relacionamentos entre entidades;

integrações externas;

configurações;

constantes;

utilitários;

funções compartilhadas.

Não invente nada que já possa ser descoberto no código existente.



2. INVENTÁRIO DO SISTEMA EXISTENTE

Antes do rebuild, faça um inventário interno de:

Frontend

páginas;

componentes;

layouts;

rotas;

hooks;

contexts;

providers;

services;

utils;

types;

schemas;

stores.

Backend/API utilizada pelo frontend

Identifique:

base URLs;

endpoints;

métodos HTTP;

parâmetros;

query parameters;

request body;

headers;

autenticação;

respostas;

códigos de erro;

paginação;

filtros;

ordenação.

Dados

Identifique todas as entidades utilizadas.

Por exemplo, se existirem:

usuários;

equipamentos;

periféricos;

patrimônio;

empréstimos;

devoluções;

unidades;

categorias;

movimentações;

manutenção;

descubra como cada entidade é representada no código.

Não assuma nomes.

Não invente estruturas.

Use o código existente como referência.



3. PRESERVAÇÃO DAS APIs

ISSO É CRÍTICO

Todas as APIs existentes devem continuar funcionando.

Não substitua chamadas reais por mocks.

Não crie dados falsos apenas para conseguir montar a interface.

Não invente endpoints.

Não invente respostas.

Não invente IDs.

Não invente regras de negócio.

Se o projeto já possui integração com um backend, mantenha essa integração.

Ao reconstruir uma página, conecte a nova interface às APIs existentes.



4. MAPA DE API

Durante a análise, identifique todos os endpoints utilizados atualmente.

Para cada endpoint, compreenda:

Método

Endpoint

Autenticação

Parâmetros

Body

Resposta

Erros

Onde é utilizado

Qual funcionalidade depende dele

Não precisa necessariamente criar um documento visual se isso não for necessário, mas você deve compreender completamente essas relações antes de alterar o frontend.

Se encontrar algo como:

GET /...

POST /...

PUT /...

PATCH /...

DELETE /...

preserve o comportamento existente.



5. AUTENTICAÇÃO

Analise profundamente o sistema de autenticação existente.

Identifique:

login;

logout;

sessão;

token;

refresh;

persistência;

expiração;

redirecionamento;

proteção de rotas;

usuário atual;

permissões;

roles.

Não recrie autenticação do zero se ela já existir e estiver funcionando.

A nova interface deve utilizar a infraestrutura existente.



6. AUTORIZAÇÃO E ROLES

Descubra todas as roles/permissões existentes no código.

Por exemplo:

Gestor;

Técnico;

Administrador;

Usuário;

ou qualquer outra existente.

Não invente novas roles.

Mapeie quais funcionalidades cada perfil pode utilizar.

A nova UI deve refletir corretamente essas permissões.

IMPORTANTE:

O frontend apenas representa permissões.

A segurança real continua sendo responsabilidade do backend.



7. VARIÁVEIS DE AMBIENTE

Procure e preserve todas as variáveis de ambiente relevantes.

Especialmente:

URLs de API;

chaves públicas;

configurações;

ambientes;

serviços externos.

Não coloque secrets diretamente no frontend.

Não exponha credenciais.

Não substitua uma variável de ambiente por uma URL hardcoded.



8. SERVIÇOS E CAMADA DE API

Se já existir algo como:

services/

api/

lib/

hooks/

repositories/

analise cuidadosamente antes de substituir.

Determine se é melhor:

reutilizar;

refatorar;

reorganizar;

encapsular;

em vez de duplicar.

O objetivo é evitar terminar com:

API antiga

+

API nova

+

serviços duplicados

+

hooks duplicados

Quero uma única arquitetura coerente.



9. TIPOS E MODELOS

Identifique todos os:

interfaces;

types;

enums;

DTOs;

schemas;

modelos de resposta;

modelos de request.

Preserve-os quando forem relevantes.

Não transforme tudo em any.

Não utilize casts apenas para fazer o TypeScript parar de reclamar.

Se houver tipos ruins:

melhore os tipos.



10. REGRAS DE NEGÓCIO

Descubra as regras existentes no frontend.

Exemplos:

quando um equipamento pode ser emprestado;

quando pode ser devolvido;

quais status existem;

quais campos são obrigatórios;

quais ações são permitidas;

quais dados dependem de outros;

quais estados são possíveis.

Não invente regras.

Se uma regra estiver implementada no backend e apenas refletida no frontend:

não replique desnecessariamente a regra no frontend.

O backend continua sendo a autoridade.



11. FLUXOS EXISTENTES

Mapeie todos os fluxos importantes.

Exemplos:

Login

Login → autenticação → sessão → dashboard.

Cadastro

Formulário → validação → API → resposta → atualização → feedback.

Empréstimo

Seleção → validação → API → atualização do item → feedback.

Devolução

Identificação → validação → API → atualização → feedback.

Consulta

Busca → filtros → API → tabela → detalhes.

Exclusão

Ação → confirmação → API → atualização → feedback.

Preserve esses fluxos mesmo que a interface seja completamente diferente.



12. REBUILD VISUAL

Depois de compreender completamente o funcionamento do sistema:

reconstrua o frontend.

Você possui liberdade total de design.

Pode mudar completamente:

layout;

navegação;

sidebar;

topbar;

dashboard;

tabelas;

formulários;

cards;

dialogs;

filtros;

organização;

tipografia;

cores;

ícones;

animações;

responsividade.

Não copie o frontend atual.

Use-o como referência funcional.



13. NOVA EXPERIÊNCIA DE PRODUTO

Projete como um produto corporativo de grande porte.

Pense como:

Staff Frontend Engineer;

Senior Product Designer;

UX Engineer;

Design Systems Engineer;

Accessibility Engineer.

A interface deve ser:

profissional;

moderna;

clara;

rápida;

consistente;

escalável;

responsiva.



14. DESIGN SYSTEM

Crie uma nova linguagem visual consistente.

Defina padrões para:

cores;

tipografia;

spacing;

radius;

borders;

shadows;

buttons;

inputs;

selects;

badges;

cards;

tables;

dialogs;

dropdowns;

tooltips;

navigation.

Evite componentes duplicados.



15. DASHBOARD

Redesenhe o dashboard com foco em informação útil.

Utilize os dados REAIS fornecidos pelas APIs existentes.

Não crie números fictícios.

Os indicadores devem representar dados reais do sistema.

Priorize:

patrimônio;

estoque;

disponibilidade;

empréstimos;

devoluções;

movimentações;

alertas;

manutenção;

informações relevantes para gestão.

Não adicione gráficos apenas para preencher espaço.



16. TABELAS

Reconstrua completamente as tabelas se necessário.

Utilize os dados reais da API.

Suporte adequadamente:

busca;

filtros;

ordenação;

paginação;

status;

ações;

seleção;

loading;

empty;

erro.

Se a API possuir paginação, utilize a paginação da API.

Não carregue milhares de registros apenas para paginar visualmente no frontend sem necessidade.



17. FORMULÁRIOS

Reconstrua os formulários mantendo os contratos atuais da API.

Garanta:

validação;

mensagens de erro;

loading;

disabled;

sucesso;

prevenção de submit duplicado;

campos obrigatórios;

feedback.

Não altere nomes de campos enviados à API sem entender o impacto.



18. RESPONSIVIDADE

Projete responsividade desde o início.

Não transforme desktop em mobile apenas diminuindo elementos.

Crie experiências apropriadas para:

320px;

375px;

390px;

430px;

tablet;

notebook;

desktop;

ultrawide.

Especialmente:

navegação;

tabelas;

formulários;

filtros;

dialogs;

cards;

dashboard.



19. MOBILE

O mobile deve ser tratado como uma experiência própria.

Pode alterar:

navegação;

posição dos elementos;

ordem das informações;

quantidade de dados visíveis;

filtros;

ações.

Uma tabela pode se transformar em:

cards;

linhas expansíveis;

detalhes;

scroll horizontal controlado;

quando isso melhorar a experiência.



20. ESTADOS

Todas as páginas devem lidar corretamente com:

loading;

empty;

error;

success;

disabled;

pending.

Nunca deixe uma API carregando sem feedback.

Nunca mostre uma página vazia sem contexto.

Nunca mostre erro técnico bruto ao usuário quando uma mensagem melhor puder ser apresentada.



21. ACESSIBILIDADE

Garanta:

contraste;

foco;

keyboard navigation;

aria;

semântica;

labels;

dialogs acessíveis;

tabelas acessíveis.

Respeite:

prefers-reduced-motion.



22. PERFORMANCE

Não faça chamadas desnecessárias à API.

Evite:

requests duplicados;

loops de requests;

renders desnecessários;

listeners vazando;

efeitos incorretos.

Analise cuidadosamente useEffect.

Não faça atualização de estado durante o render.



23. COMPATIBILIDADE COM A API

Antes de considerar cada funcionalidade concluída, verifique:

GET

Os dados aparecem corretamente?

POST

O body enviado corresponde ao contrato?

PUT/PATCH

A atualização funciona?

DELETE

A remoção funciona?

Errors

Os erros são tratados?

Loading

O usuário recebe feedback?

Refresh

Os dados são atualizados corretamente depois de uma mutação?



24. NÃO USE MOCKS

Esta regra é obrigatória.

Não substitua dados reais por mocks.

Não use:

fakeUsers

fakeEquipment

fakeLoans

mockData

dummyData

para simular funcionalidades que já possuem API.

Se o projeto existente já possui mocks exclusivamente para desenvolvimento, preserve-os apenas onde realmente fizerem parte do ambiente de desenvolvimento.

A aplicação final deve utilizar os dados reais.



25. NÃO INVENTE FUNCIONALIDADES

Você possui liberdade total para design.

Não possui liberdade para inventar regras de negócio.

Pode melhorar:

como o usuário faz algo.

Não pode inventar:

o que o sistema deveria fazer.



26. ARQUITETURA FINAL

O resultado deve evitar:

componente gigante

→ API

→ regra de negócio

→ UI

→ estado

→ validação

→ tudo no mesmo arquivo

Prefira separar responsabilidades de maneira coerente.

Exemplo conceitual:

UI

↓

Hooks

↓

Services/API

↓

Backend

Mas não crie camadas desnecessárias apenas para seguir um padrão.



27. MIGRAÇÃO

Durante a reconstrução:

Preserve a infraestrutura funcional.

Preserve APIs.

Preserve autenticação.

Preserve tipos relevantes.

Preserve regras.

Preserve integrações.

Substitua a camada visual.

Refatore a arquitetura quando necessário.

Teste cada fluxo.

Remova código antigo somente quando tiver certeza de que não é mais utilizado.

Não deixe código morto espalhado pelo projeto.



28. VALIDAÇÃO FINAL

Depois da reconstrução:

API

Verifique todas as integrações.

Autenticação

Teste login/logout/sessão.

Rotas

Teste todas as rotas.

Permissões

Teste os diferentes perfis.

CRUD

Teste:

criar;

visualizar;

editar;

excluir;

quando aplicável.

Fluxos

Teste:

empréstimo;

devolução;

cadastro;

consulta;

movimentação.

UI

Teste:

desktop;

tablet;

mobile.

Estados

Teste:

loading;

empty;

error;

success.

Build

Execute o build.

TypeScript

Corrija erros.



29. REGRA DE OURO

Antes de substituir qualquer código importante, faça esta pergunta:

“Esse código contém alguma informação necessária para o funcionamento do sistema?”

Se sim:

preserve ou migre a lógica.

Se for apenas apresentação visual:

você possui liberdade para substituir.



30. DEFINIÇÃO DE “DO ZERO”

Quando digo:

“faça o frontend do zero”

estou dizendo:

NÃO reutilize obrigatoriamente a interface atual.

Não estou dizendo:

“ignore o código existente”.

Muito pelo contrário.

O código existente deve ser tratado como fonte de conhecimento sobre o sistema.

Primeiro entenda.

Depois preserve a lógica.

Depois reconstrua a experiência.



31. PROIBIDO

É proibido:

inventar APIs;

inventar endpoints;

inventar dados;

substituir API real por mock;

remover autenticação existente;

remover permissões;

quebrar contratos;

hardcodar URLs que deveriam vir de ambiente;

expor secrets;

transformar tipos em any sem necessidade;

apagar funcionalidades para simplificar o rebuild;

alterar regras de negócio sem justificativa.



32. CRITÉRIO DE QUALIDADE

O novo frontend deve ser simultaneamente:

FUNCIONAL

●

CONECTADO À API REAL

●

VISUALMENTE EXCELENTE

●

RESPONSIVO

●

ACESSÍVEL

●

PERFORMÁTICO

●

MANUTENÍVEL

●

ESCALÁVEL



33. TESTE DE SENIORIDADE

Depois de terminar, faça uma revisão crítica do próprio trabalho.

Pergunte:

Eu mantive todas as integrações existentes?

Alguma API foi perdida?

Algum endpoint foi substituído por mock?

Alguma regra de negócio foi alterada acidentalmente?

Algum fluxo ficou pior?

Algum componente ficou complexo demais?

A arquitetura está realmente melhor?

O mobile foi projetado de verdade?

O dashboard utiliza dados reais?

As tabelas utilizam a API corretamente?

Os estados de erro estão tratados?

A interface parece um produto empresarial real?

Eu aprovaria esse projeto em um Pull Request de uma empresa grande?

Se a resposta for não para qualquer ponto importante:

corrija antes de considerar o trabalho concluído.



OBJETIVO FINAL

Quero que você faça algo equivalente a:

ENGENHARIA REVERSA DO SISTEMA EXISTENTE

↓

MAPEAMENTO DE APIs, DADOS, REGRAS E FLUXOS

↓

PRESERVAÇÃO DA LÓGICA E INTEGRAÇÕES

↓

RECONSTRUÇÃO COMPLETA DO FRONTEND

↓

NOVO DESIGN SYSTEM

↓

NOVA UX

↓

RESPONSIVIDADE REAL

↓

TESTES E VALIDAÇÃO

O resultado final deve dar a sensação de que:

uma equipe profissional recebeu um sistema existente, entendeu profundamente seu funcionamento e depois reconstruiu sua experiência de frontend de maneira muito mais madura, sem perder as funcionalidades ou integrações existentes.

Não copie o frontend antigo.

Não descarte a lógica antiga.

Entenda primeiro. Preserve o que importa. Reconstrua o que precisa ser reconstruído.

This project was built with [Lovable](https://lovable.dev).

## Build with Lovable

Continue developing this project in the [Lovable editor](https://lovable.dev/projects/6d5da3ed-f7f4-4e92-adb4-508f687c3872).

- **Ship faster**: describe what you want to build and Lovable handles the code.
- **Stay in sync**: every change made in Lovable is committed straight to this repository.
- **Full ownership**: this code is yours. Push to `main` on GitHub and your changes sync back into Lovable, ready for your next prompt.

## Development

Prefer working locally? You need Node.js and npm — [install with nvm](https://github.com/nvm-sh/nvm#installing-and-updating).

```sh
git clone <this-repository-url>
cd <repository-name>
npm i
npm run dev
```
