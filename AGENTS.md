# AGENTS.md — contexto do projeto Patrimônio Revalle

Este arquivo é para quem chega ao projeto sem contexto: uma pessoa nova ou um agente de código (Codex, Claude Code etc.). Leia inteiro antes de mudar qualquer coisa. Os detalhes estão no `README.md`, em `docs/` e em `backend/migrations/README.md`; aqui fica o mapa e as regras.

Estado descrito: `main` em 26/11/2026 (PRs #2 a #5 mesclados e em produção).

---

## 1. Regras de trabalho (obrigatórias)

1. **Nunca trabalhe direto na `main`.** Para cada mudança, crie um branch a partir da `main` atualizada: `feat/<assunto>`, `fix/<assunto>`, `docs/<assunto>` ou `chore/<assunto>`.
2. **Commit a cada mudança, por menor que seja.** Uma mudança lógica por commit, com mensagem em português no padrão `tipo(escopo): descrição` (ex.: `fix(coletor): mensagem do servidor sem JSON cru`). Tipos usados: `feat`, `fix`, `docs`, `test`, `chore`, `ci`.
3. **Termine com um pull request para a `main`.** Quem faz o merge é o responsável pelo sistema, no GitHub. Não faça merge nem push na `main`.
4. **O CI precisa passar** (GitHub Actions, `.github/workflows/ci.yml`: testes do backend contra MySQL de teste e tipos, testes e build do frontend). Rode localmente antes de abrir o PR (seção 6).
5. **Teste antes de implementar** (TDD): escreva o teste que falha, depois o código. Mudança sem teste não entra.
6. **Segredos nunca vão para o repositório.** Senhas, tokens e chaves ficam só no `.env` do servidor e nas variáveis do EasyPanel. `backend/.env` está no `.gitignore`; use `backend/.env.example` como modelo, com valores de exemplo.
7. **Nunca rode testes nem scripts contra o banco de produção.** A suíte se recusa a rodar se as variáveis apontarem para produção. Não contorne essa trava.
8. **A base de RH (Postgres) é somente leitura.** O sistema só lê os colaboradores de lá. Nunca escreva nela.
9. **Mudanças de banco só por migração aditiva** (seção 5). Nunca apague nem renomeie coluna ou tabela existente.
10. **Mensagens para o usuário sempre em português**, claras, sem termos técnicos.
11. **Atualize este arquivo** quando mudar algo que ele descreve: módulo novo, regra nova, variável nova.

---

## 2. O que é o sistema

O **Patrimônio** é o controle de equipamentos de TI da Revalle (revendas e unidades na BA e em PE). Ele registra:
- cada equipamento (notebook, desktop, celular, impressora, tablet, switch, HD, nobreak, access point) e seus periféricos;
- com qual colaborador o equipamento está, com termo de responsabilidade assinado;
- manutenções, transferências entre unidades, baixa com apagamento de dados (LGPD), inventário físico com etiqueta QR, linhas telefônicas e os acessos do colaborador a sistemas (admissão, desligamento, mudança de função).

Ele substituiu um sistema desktop em Tkinter. `dist/Controle de Estoque.exe` é desse sistema antigo, guardado só como histórico; não mexa.

Endereço em produção: **https://patrimonio.revalle.com.br**

### Papéis
| Papel | Pode |
|---|---|
| **Gestor** | Tudo, mais: baixa, estorno, usuários, unidades, auditoria, catálogo de acessos e sincronização manual de colaboradores. 2FA obrigatório. |
| **Técnico** | Operação do dia a dia: cadastro, empréstimo, devolução, termos, manutenções, inventário, etiquetas, linhas, colaboradores, relatórios. |
| **Jovem Aprendiz** | Só leitura: painel, estoque, gráficos e linhas. Recebe o CPF mascarado e não vê e-mail nem telefone. |

As regras ficam no backend (`gestor_only`, `gestor_or_tecnico`, `require_role` em `backend/app/dependencies.py`). O frontend só esconde botões; **a permissão de verdade é a do backend**, e toda regra nova precisa de teste de API.

---

## 3. Arquitetura

```
/
├── backend/                 API FastAPI (Python 3.12)
│   ├── app/
│   │   ├── main.py          cria o app, registra routers, roda migrações e agendador no startup
│   │   ├── core/            config (Settings via pydantic-settings), security (JWT, hash)
│   │   ├── dependencies.py  autenticação, papéis, IP do cliente, rate limit (slowapi)
│   │   ├── routers/         um arquivo por módulo da API (/api/...)
│   │   ├── schemas/         modelos Pydantic de entrada e saída (validação de domínio fica aqui)
│   │   ├── db/              acesso ao MySQL com SQL puro (PyMySQL, get_cursor), um manager por módulo
│   │   ├── services/        regras de negócio e integrações (e-mail, RH, RevSign, PDF, exportações, auditoria...)
│   │   ├── templates/       e-mails (HTML) e o script do coletor (PowerShell)
│   │   └── jobs.py          tarefas agendadas
│   ├── migrations/          NNNN_*.sql + NNNN_*.down.sql
│   ├── modelos/             modelos .docx dos termos (versionados)
│   ├── scripts/             migrate, reset_2fa, marcar_assinatura, seed e diagnósticos
│   └── tests/               pytest (unitários e de integração contra MySQL de teste)
├── frontend/                SPA React 18 + TypeScript + Vite
│   └── src/
│       ├── api/             um cliente por módulo (axios em api/client.ts)
│       ├── pages/           telas (uma por rota)
│       ├── features/        partes de tela por módulo (estoque, item-detail, coletor, inventario...)
│       ├── components/      componentes comuns (ui/ do shadcn, layout/, equipment/...)
│       ├── contexts/        AuthContext (login, logout, sessão)
│       └── test/            MSW, render com React Query, setup do Vitest
├── docs/                    CI, testes, decisões e planos (docs/superpowers/specs e plans)
├── docker-compose.yml       produção local/self-hosted (backend + frontend com nginx)
└── docker-compose.test.yml  MySQL de teste efêmero (porta 3307)
```

**Backend:** FastAPI 0.115, Pydantic 2, PyMySQL (SQL escrito à mão, sem ORM), slowapi (limite de requisições), python-docx (termos), openpyxl e fpdf2 (Excel e PDF), psycopg (leitura da base de RH), pyotp e cryptography (2FA), boto3 (storage S3 opcional), LibreOffice headless na imagem (conversão de termos para PDF).

**Frontend:** React 18, React Router 6, TanStack Query 5, Tailwind 3 com componentes shadcn/ui, Vitest e Testing Library com MSW. O nginx do container do frontend serve o build e faz proxy de `/api/` para o backend.

**Bancos:**
- **MySQL** (produção: MySQL 9; testes: MySQL 8.0.40) é o banco do sistema. As tabelas antigas (`items`, `history`, `peripherals`, `equipment_peripherals`, `usuarios`, `unidades`, `revoked_tokens`) são criadas pelos managers em `app/db/`; todo o resto vem das migrações.
- **Postgres de RH** (tabela `"colaboradores-revalle"`) é só leitura. Uma sincronização diária copia os colaboradores para a tabela `colaboradores` do MySQL.

---

## 4. Módulos e fluxos principais

### Equipamento e situação (`items.status`)
`Disponível` → empréstimo iniciado → `Pendente` (aguarda o termo assinado) → confirmado → `Indisponível` (em uso) → devolução iniciada → `Pendente Devolução` → confirmada → `Disponível`.
Também há `Em Manutenção`. A baixa marca `is_active = 0`; equipamentos de dados (notebook, desktop, celular, tablet, HD) exigem o registro do apagamento.
Toda operação grava em `history`, que também alimenta o estorno (só Gestor, com reconfirmação de senha). Os campos que cada tipo de equipamento mostra ficam em `frontend/src/components/equipment/TypeSpecificFields.tsx` (`TYPE_FIELDS`).

### Colaboradores (`/api/colaboradores`)
Vêm da base de RH, com sincronização diária (`COLABORADORES_SYNC_HORA`) ou manual pelo Gestor. E-mail, telefone e superior imediato podem ser completados à mão, e a sincronização não sobrescreve o que foi editado. Quem sai da base vira `desligado`, sem ser apagado. O CPF é validado pelos dígitos verificadores.

### Termos de responsabilidade (`/api/termos`, `services/termos*.py`)
Tipos: `entrega`, `devolucao` e `uso`. O termo `uso` é pedido a qualquer momento e pode ser agendado. O termo é gerado em `.docx` a partir de `backend/modelos/`, preenchendo marcadores `{{campo}}`.
- **Integração com o RevSign** (assinatura eletrônica): o contrato está em `docs/integracoes/revsign-contrato-patrimonio.md`.
  - O termo é convertido para PDF com LibreOffice (`services/termo_pdf.py`) e enviado com o CPF do colaborador. O termo `uso` vai como `ENTREGA`, por decisão do responsável.
  - A volta chega por um webhook com HMAC (`/api/integrations/revsign/webhook`) e por consulta periódica.
  - Toda falha cai em **assinatura manual**, com o motivo na tela.
- **Estado atual:** integração **desligada** (`REVSIGN_ENABLED=false`). Os termos seguem o fluxo manual: baixar, assinar e anexar o PDF.
- Os modelos têm o marcador invisível `{{assinatura}}` (fonte branca), que o RevSign usa para posicionar o campo de assinatura. Se trocar um modelo, rode `python scripts/marcar_assinatura.py` a partir de `backend/`.

### Login e sessão (`/api/auth`, `/api/users`)
- **Entrada:** login com e-mail `@revalle.com.br` e senha. Contas antigas, que entravam por nome de usuário, confirmam o e-mail com um código no próximo acesso.
- **Etapas do login:** a resposta traz `etapa` = `ok`, `mfa`, `cadastro_2fa` ou `cadastro_email`, com um token temporário para cada etapa.
- **2FA:** TOTP, obrigatório para Gestor, com códigos de recuperação. O segredo é cifrado com `TOTP_ENCRYPTION_KEY`, que **não pode mudar**: trocar a chave invalida os autenticadores.
- **Contas novas:** o Gestor informa só o e-mail e o papel, e a pessoa recebe um convite para criar a senha. Ninguém define a senha de outra pessoa.
- **Sessão:** o access token fica em memória e o refresh token num cookie httpOnly. `usuarios.sessao_versao` permite encerrar todas as sessões de um usuário.

### Coletor de dados da máquina (`/api/coletor`, `features/coletor/`)
1. Na máquina que vai cadastrar, o técnico abre `/register/coletor` e copia um comando para o **Win+R**.
2. O comando baixa um script PowerShell (`backend/app/templates/coletor/coletor.ps1`). O script lê marca, modelo, serial, host, domínio, CPU, RAM, disco, sistema, licença (Product ID), AnyDesk e MAC, e envia em JSON.
3. O backend normaliza para os campos do formulário (`services/coletor_normalizacao.py`) e procura uma máquina já cadastrada por serial, MAC ou AnyDesk.
4. O cadastro abre preenchido. Ao salvar, a página **sai da conta do técnico** naquele computador.

Detalhes que importam:
- O código do comando é de uso único, vale 15 minutos e o banco guarda só o hash SHA-256.
- As rotas públicas do coletor são autorizadas só por esse código e têm limite de requisições.
- Adaptadores USB (dock) não identificam a máquina.
- **Pendente:** testar o Win+R numa máquina real da empresa. Até agora o script só rodou em modo "só mostrar", sem enviar.

### Outros módulos
- `manutencoes`: ordens de manutenção, envio à assistência, custo e retorno.
- `transferencias`: equipamento entre unidades.
- `linhas`: linhas telefônicas, com eventos e relatório de custo por unidade.
- `inventarios` e `etiquetas`: folha A4 de etiquetas QR (o QR aponta para `APP_PUBLIC_URL/e/{id}`), leitura pela câmera do celular e relatório de divergências.
- `painel`: indicadores e alertas calculados na hora (garantia, termo pendente, devolução atrasada, manutenção parada, desligado com equipamento).
- `exportacoes`: Excel e PDF das listagens.
- `acessos`: catálogo de sistemas, pacotes por cargo e tarefas de criar ou desativar acesso. Guarda só o login, **nunca senha**.
- `audit`: registro de alterações (`services/audit.py`, `audit.record(...)`) em todo fluxo de escrita. Mantenha isso em funcionalidades novas.
- `locais` e `unidades`: locais físicos por unidade; a unidade tem `alias_rh` para ligar ao código da base de RH (RV-PE, RV-JZ...).

### Tarefas agendadas (`app/jobs.py`, `services/scheduler.py`)
Um tique a cada 60 s, e cada job roda sob `GET_LOCK` (são 2 workers). Os jobs são:
- `sincronizar_colaboradores` (diário);
- `limpar_coletas` (diário, 03:30);
- `termos_envios_agendados` e `termos_consulta_revsign` (só com o RevSign ligado).

`SCHEDULER_ENABLED=false` desliga todos.

---

## 5. Banco de dados e migrações

- Arquivos em `backend/migrations/`: `NNNN_descricao.sql` e o par **obrigatório** `NNNN_descricao.down.sql`.
- Rodam sozinhas na subida da API, sob trava, e também por `python -m scripts.migrate [aplicar|pendentes|reverter]` a partir de `backend/`.
- Regras:
  - somente mudanças aditivas;
  - `IF NOT EXISTS` / `IF EXISTS`, porque a migração precisa poder ser reaplicada;
  - `CREATE TABLE` **sem** charset ou collation, para herdar o padrão do banco;
  - cada comando termina com `;` no fim da linha;
  - para alterar uma coluna condicionalmente, use o padrão `information_schema` + `PREPARE`, como em `0900_usuarios_email_login.sql`.
- Faixas já usadas: 0001–0899 (fases 0 a 8), 0900 (login por e-mail), 1000 (coletor). **Use o próximo número livre de uma faixa nova (ex.: 1100) para cada funcionalidade nova.** Detalhes em `backend/migrations/README.md`.
- A API sobe mesmo se uma migração falhar; o erro só vai para o log. Por isso, depois de todo deploy, confira `python -m scripts.migrate pendentes` → "nenhuma".

---

## 6. Rodar e testar

### Backend
```bash
cd backend
cp .env.example .env        # preencha com valores de DESENVOLVIMENTO, nunca os de produção
pip install -r requirements-dev.txt
uvicorn app.main:app --reload --port 8000      # docs em http://localhost:8000/api/docs
```

**Testes:** `pytest` contra um MySQL de teste descartável. Veja `docs/TESTES.md`.
```bash
docker compose -f docker-compose.test.yml up -d          # MySQL 8 em 127.0.0.1:3307, dados em memória
export TEST_DB_HOST=127.0.0.1 TEST_DB_PORT=3307 TEST_DB_NAME=stock_sys_test \
       TEST_DB_USER=stock_test TEST_DB_PASSWORD=stock_test_pw \
       JWT_SECRET=teste-jwt-secret-com-mais-de-32-caracteres
cd backend && python -m pytest -m "not mudanca_esperada"
docker compose -f docker-compose.test.yml down -v       # ao terminar
```
- Sem `TEST_DB_HOST`/`TEST_DB_NAME`, os testes de integração são **pulados**. Só os marcados `unit` rodam (`python -m pytest -m unit`).
- Os testes da conversão para PDF precisam do LibreOffice (`soffice`) e do `pdftotext` (poppler). Sem eles, esses testes são pulados.
- Os testes que leem a base de RH usam um Postgres temporário (`COLAB_TEST_PG_ADMIN_URL`); sem ele, são pulados.
- Fixtures úteis em `backend/tests/conftest.py`: `client`, `client_gestor`, `client_tecnico`, `client_aprendiz`, `criar_item`, `criar_usuario`, `executar_sql`. O banco é zerado entre testes.
- Ambiente sem Docker (ex.: Codex na nuvem): instale um MySQL 8 no script de setup, crie o banco e o usuário de teste e exporte as variáveis acima. Se não for possível, rode pelo menos `-m unit` e diga no PR que a integração ficou para o CI.

### Frontend
```bash
cd frontend
npm ci
npm run dev                 # http://localhost:5173, proxy /api → localhost:8000
npx tsc --noEmit && npm run test:run && npm run build   # o mesmo que o CI roda
```
O MSW simula todas as chamadas: rota sem handler é erro de teste. Use os helpers de `src/test/`.

---

## 7. Produção (só para entender; quem publica é o responsável)

- **Onde roda:** VPS com **EasyPanel** (Docker Swarm), projeto `agente_revalle`, serviços `stock-sys-backend` e `stock-sys-frontend`. O EasyPanel publica a partir da **`main`** do GitHub quando o responsável clica em Deploy.
- **Serviços que o backend usa** (pela rede interna do EasyPanel):
  - MySQL `agente_revalle_mysql`, banco `stock_sys_db`, usuário próprio do sistema;
  - Postgres de RH `agente_revalle_postgres`, com usuário de leitura;
  - RevSign `agente_revalle_revsign`, em https://assinatura.revalle.com.br.
- A porta do MySQL está fechada para a internet.
- **Variáveis:** lista completa e comentada em `backend/.env.example` e na tabela do `README.md`. Novas variáveis precisam de valor padrão seguro, para o sistema não parar de subir quando faltarem.
- **E-mail:** SMTP `mail.revalle.com.br:465`, conta `naoresponda@revalle.com.br`.
- **Depois de um deploy:** `python -m scripts.migrate pendentes` dentro do container do backend precisa responder "nenhuma".
- **Antes de migração arriscada:** o responsável faz backup do banco. Se a sua mudança tem migração que reconstrói tabela ou mexe em dados, diga isso no PR.

---

## 8. Convenções de código

- **Backend:**
  - um router, um manager em `app/db/` e schemas próprios por módulo;
  - validação de domínio nos schemas Pydantic (erro 422);
  - SQL sempre parametrizado, nunca com f-string de valor;
  - grave auditoria em toda escrita;
  - horários: o container roda com `TZ=America/Sao_Paulo`; os módulos antigos e os termos gravam horário de Brasília, e algumas tabelas novas (coletas, códigos de e-mail) gravam UTC. Siga o que o módulo que você está mexendo já usa e deixe isso claro no nome ou no comentário da coluna;
  - não logue segredos, tokens, senhas nem dados pessoais completos (CPF mascarado).
- **Frontend:**
  - telas em `pages/`, partes de módulo em `features/<modulo>/`, chamadas em `api/<modulo>.ts` com React Query;
  - use os componentes existentes: `Button` com `loading`, `StatusPill`, `Pagination`, `PageHeader`, `ErrorState`;
  - nomes de unidade e revenda em CAIXA ALTA na exibição (`upper()`).
- **Textos:** em português, curtos e claros. Erros dizem o que fazer ("Cadastre o e-mail desta conta primeiro.").
- **Segurança:**
  - rotas públicas (sem login) precisam de limite de requisições, limite de tamanho e respostas que não revelem se um código ou usuário existe;
  - toda entrada vinda de fora (webhook, coletor, RH) é dado não confiável.

---

## 9. Pendências e ideias conhecidas

- **Coletor:** testar o comando do Win+R numa máquina real da empresa (usuário sem administrador, notebook na dock, VPN, AnyDesk ausente). Ajustar se o Windows ou o antivírus bloquearem.
- **RevSign:** para ligar, é preciso:
  1. criar a caixa `patrimonio@revalle.com.br`;
  2. criar a conta remetente no RevSign com esse e-mail;
  3. configurar as variáveis dos dois lados (README, seção RevSign);
  4. fazer um termo de teste.
  
  A chave e o segredo ficam com o responsável.
- **Fora de escopo até agora** (cartão "Patrimônio – Ajustes e melhorias" no Trello): gestão de telefonia (MDM), envio de termo por WhatsApp, travar o equipamento até a assinatura, restrição de visão por unidade, revisão periódica de acessos pelo superior.
- O histórico das decisões de cada entrega está em `docs/superpowers/specs/` (desenho) e `docs/superpowers/plans/` (plano). Leia o da área que for mexer.
