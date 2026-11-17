# Patrimônio — Ajustes e melhorias (cartão do Trello) — Design

Data: 12/11/2026 · Branch: `feat/patrimonio-melhorias` · Origem: cartão "Patrimônio – Ajustes e melhorias" do quadro Projetos Revalle (14 checklists, feito por outra pessoa).

## 1. Objetivo e critério de sucesso

Atender as solicitações do cartão que fazem sentido para o sistema em produção, entregando tudo junto numa única branch, sem quebrar o que já funciona.

Pronto significa: cada item marcado como "entra" na matriz da seção 3 funcionando, com testes; suíte completa do backend e do frontend verde; build do frontend ok; migrações aplicadas do zero e sobre uma cópia do schema atual; nenhum fluxo atual (empréstimo, devolução, termos, baixa, estorno, relatórios) quebrado.

## 2. Decisões tomadas com o Enzo (12/11/2026)

| Tema | Decisão |
|---|---|
| MDM (checklist 5) e travar equipamento (6.3) | Fora do escopo |
| WhatsApp (3.2, 3.3) e canal de atendimento do TI (3.6) | Fora desta rodada |
| Script do processo de patrimônio (8.3) | Fora: ninguém sabe qual é |
| Banco MySQL | Usado só pelo sistema web: alterações aditivas (ADD COLUMN NULL, novas tabelas, novo valor de ENUM) são permitidas, sempre com migração versionada |
| Paleta | Base `#0131CD` (azul), `#FFFFFF` (branco), `#B9FC01` (verde-limão) |
| Colaboradores | Vêm da tabela `"colaboradores-revalle"` do Postgres `agente_revalle` (mesma base do RevSign), **somente leitura** |
| Superior imediato | Campo de seleção (outro colaborador) no cadastro do colaborador |
| E-mail do colaborador | Obrigatório para receber equipamento; pode ser pessoal |
| Termos | Assinatura pelo RevSign, que envia o link por e-mail. O RevSign será alterado (no repositório dele) para aceitar e-mail pessoal nos termos vindos do Patrimônio |
| E-mail do sistema | SMTP `mail.revalle.com.br:465`, usuário `naoresponda@revalle.com.br`, senha só no `.env` |
| 2FA | App autenticador (TOTP), obrigatório para Gestor, opcional para os demais, com códigos de recuperação |
| Perfis | Continuam os 3 papéis atuais (Gestor, Técnico, Jovem Aprendiz), vendo todas as unidades |
| Localização | Locais cadastrados por unidade |
| Alertas | Só no painel (sem e-mail) |
| Acessos | Admissão, desligamento e mudança de função registrados à mão |
| Revisão periódica de acessos pelo superior (14.13) | Fora |
| Etiquetas | Folha A4 adesiva |
| Entrega | Tudo de uma vez; execução multiagente a critério da implementação |
| Produção | Patrimônio roda na mesma VPS/EasyPanel do RevSign e do Postgres de RH |

## 3. Matriz do cartão (rastreabilidade)

Legenda das fases na seção 6. "Já atende" = o sistema atual cobre.

| Item do cartão | Destino |
|---|---|
| 1.1 Esqueci minha senha por e-mail | F2 |
| 1.2 2FA no login | F2 |
| 2.1 Aumentar paginação (hoje 7) | F0 (componente) + F1 (aplicar) |
| 2.2 Escolher itens por página (25/50/100) | F0 + F1 |
| 2.3 Status arredondados, cor por status, contraste | F0 (componente) + F1 |
| 2.4 Setor e Cargo na tabela de estoque | F1 (depende do cadastro de colaborador, F0) |
| 3.1 Assinatura dos termos no RevSign | F3 |
| 3.2 / 3.3 WhatsApp | Fora |
| 3.4 Telefone e e-mail no cadastro, validados | F0 (cadastro) + F3 (exigência no envio) |
| 3.5 Termo por e-mail | F3 (o RevSign envia o link) |
| 3.6 Canal de atendimento do TI | Fora |
| 4.1 / 4.2 Manutenção com histórico e custo | F4 |
| 5.1–5.11 Gestão de Telefonia (MDM) | Fora |
| 6.1 Termo de uso no equipamento (vigente, status, data) | F3 |
| 6.2 TI solicita a assinatura quando quiser | F3 |
| 6.3 Travar até assinar | Fora |
| 6.4 Agendar a solicitação | F3 (agenda o envio, sem bloqueio) |
| 6.5 Disparo automático na troca de aparelho ou usuário | F3 |
| 6.6 Histórico de quem solicitou e quando assinou | F3 |
| 7.1 Padrão visual único | F1 |
| 7.2 Painel com indicadores | F7 |
| 7.3 Busca, filtros e ordenação nas listagens | F1 |
| 7.4 Tela de detalhe com histórico completo | F0 (esqueleto) + abas nas fases 1, 3, 4, 5, 8 |
| 7.5 Responsividade | F1 |
| 7.6 / 7.7 Paleta e painéis | F0 (tokens) + F1 |
| 7.8 Carregamento nos botões e clique único | F0 (componente) + F1 |
| 7.9 Status em negrito | F0 + F1 |
| 7.10 Unidade e revenda em CAIXA ALTA | F1 |
| 7.11 Ir direto a uma página | F0 + F1 |
| 8.1 Validar CPF com mensagem clara | F0 |
| 8.2 Nome pelo CPF em fonte autorizada | F0 (base de colaboradores) |
| 8.3 Script do processo | Fora |
| 8.4 Localização | F0 (locais) + F4 (no equipamento) |
| 8.5 Cadastro do colaborador | F0 |
| 8.6 Superior imediato | F0 (campo) + F1 (exibir junto ao equipamento) |
| 8.7 Dados da Sólides/Tangerino com complemento manual | F0 (via `"colaboradores-revalle"`) |
| 9.1–9.5 Definições pendentes | Resolvidas na seção 2 |
| 10.1 / 10.2 Entrada com NF, valor, compra, garantia e anexo | F4 |
| 10.3 Transferência entre unidades | F4 |
| 10.4 Linha do tempo da máquina | F4 (endpoint) + detalhe |
| 10.5 Baixa com apagamento de dados (LGPD) | F4 |
| 11.1–11.4 Etiquetas QR, leitura, inventário, relatório | F6 |
| 12.1 Perfis | Já atende (3 papéis); restrição por unidade fora |
| 12.2 Log de alterações em todo o Patrimônio | F0 |
| 12.3 Alertas automáticos | F7 |
| 12.4 Exportar Excel e PDF | F0 (infraestrutura) + F6/F7 |
| 13.1–13.4 Linhas telefônicas | F5 |
| 14.1–14.12 Acessos por função | F8 |
| 14.13 Revisão periódica pelo superior | Fora |

## 4. Arquitetura

### 4.1 Princípio

Cada frente nova vira um módulo isolado: router próprio em `backend/app/routers/`, manager próprio em `backend/app/db/`, schemas próprios em `backend/app/schemas/`. O `InventoryDBManager` (1.093 linhas) recebe só ganchos mínimos (ex.: preencher `colaborador_id` no empréstimo). No frontend, páginas e componentes novos em pastas por módulo. Isso permite que fases rodem em paralelo com poucos conflitos.

### 4.2 Migrações versionadas

- Arquivos `backend/migrations/NNNN_descricao.sql`, aplicados em ordem numérica e registrados em `schema_migrations (versao, aplicada_em)`.
- Cada migração tem o par `NNNN_descricao.down.sql` com a reversão.
- Executor `app/db/migrations.py`: roda no startup do app (dentro de `GET_LOCK('patrimonio_migrations', 30)`, porque o uvicorn sobe com 2 workers) e também por `python -m scripts.migrate` (aplicar, listar pendentes, reverter a última).
- Os `CREATE TABLE IF NOT EXISTS` atuais dos managers continuam como baseline; as migrações rodam depois deles.
- Faixas de numeração por fase, para evitar colisão no trabalho paralelo: F0 `0001–0099`, F1 `0100–0199`, F2 `0200–0299`, F3 `0300–0399`, F4 `0400–0499`, F5 `0500–0599`, F6 `0600–0699`, F7 `0700–0799`, F8 `0800–0899`.
- Toda migração é idempotente quando possível (checagem em `information_schema`) e testada no banco de teste do zero.

### 4.3 Configuração nova (`backend/.env.example`, sem valores reais)

| Variável | Uso |
|---|---|
| `APP_PUBLIC_URL` | Endereço público do Patrimônio (links de e-mail e QR) |
| `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM` | E-mail (SSL implícito quando a porta é 465; `SMTP_FROM` padrão = `SMTP_USER`) |
| `COLABORADORES_DATABASE_URL` | Postgres de RH (somente leitura) |
| `COLABORADORES_TABLE` (padrão `colaboradores-revalle`), `COLABORADORES_SCHEMA` (padrão `public`), `COLABORADORES_COLUMNS` (opcional) | Mesmo contrato do RevSign |
| `COLABORADORES_SYNC_HORA` (padrão `05:00`, horário de Brasília) | Sincronização diária |
| `TOTP_ENCRYPTION_KEY` | Chave Fernet para cifrar os segredos TOTP no banco |
| `REVSIGN_ENABLED` (padrão `false`), `REVSIGN_BASE_URL`, `REVSIGN_API_TOKEN`, `REVSIGN_WEBHOOK_SECRET` | Integração RevSign (detalhes fecham com o contrato, seção 6.F3) |
| `ALERTA_GARANTIA_DIAS` (30), `ALERTA_TERMO_PENDENTE_DIAS` (3), `ALERTA_DEVOLUCAO_PENDENTE_DIAS` (7), `ALERTA_MANUTENCAO_DIAS` (15) | Limites dos alertas |

Produção: o container do backend precisa estar na rede do EasyPanel onde `agente_revalle_postgres` resolve. O `README.md` ganha a seção de deploy correspondente.

### 4.4 Serviços internos compartilhados (Fase 0)

- **`app/services/mail.py`**: `send_mail(to, subject, html, text)` com `smtplib.SMTP_SSL` (465) ou STARTTLS (587), timeout de 15 s, sem logar corpo nem senha. Em teste, um backend em memória captura as mensagens. Modelos HTML simples em `backend/app/templates/email/`.
- **`app/services/audit.py`**: `record(ctx, entidade, entidade_id, acao, antes, depois)` grava em `audit_log`. `AuditContext` (usuário, papel, IP) vem de uma dependência FastAPI; IP só de `X-Forwarded-For` quando a requisição vem do proxy confiável (nginx do próprio compose). CPF é mascarado nos valores gravados (`***.456.789-**`).
- **`app/services/colaboradores_source.py`**: único módulo que fala com o Postgres de RH. Conexão com `psycopg` em modo `default_transaction_read_only=on`; descobre colunas pelo nome no `information_schema` (mesmas regras do RevSign: cpf, nome, cargo, setor, unidade, matricula, email, telefone, situacao, e gestor/superior se existir), aceita CPF com máscara e override por `COLABORADORES_COLUMNS`. Nunca escreve. Não lê nem guarda data de nascimento (minimização LGPD: o Patrimônio não precisa dela).
- **`app/services/scheduler.py`**: tarefa assíncrona iniciada no lifespan do app, com tique a cada 60 s. Cada job roda dentro de `GET_LOCK('patrimonio_job_<nome>', 0)`, para que só um dos 2 workers execute. Jobs: sincronização diária de colaboradores (F0), envios agendados de termo e consulta de reserva ao RevSign (F3). Falha de um job é logada e não derruba o app.
- **`app/services/exports.py`**: helpers para gerar Excel (`openpyxl`) e PDF tabular (`fpdf2`) a partir de uma lista de colunas e linhas, com cabeçalho Revalle, data de geração, filtros aplicados e unidade/revenda em CAIXA ALTA.

## 5. Modelo de dados (MySQL)

Tabelas e colunas novas; tudo aditivo. Tipos indicativos; o plano fixa os detalhes.

**F0**
- `schema_migrations (versao PK, aplicada_em)`
- `colaboradores (id, cpf CHAR(11) UNIQUE, nome, matricula, cargo, setor, unidade_rh, unidade_id FK NULL, situacao ENUM('ativo','afastado','desligado'), email, telefone, superior_id FK colaboradores NULL, origem ENUM('rh','manual'), email_manual TINYINT, telefone_manual TINYINT, sincronizado_em, criado_em, atualizado_em)`. A sincronização atualiza nome, matrícula, cargo, setor, unidade e situação; atualiza e-mail/telefone só quando o campo não foi editado à mão (`*_manual = 0`); nunca mexe em `superior_id` (exceto se a base tiver gestor e o campo estiver vazio). Quem some da base vira `situacao='desligado'` (não é apagado). Colaborador `origem='manual'` (ex.: terceiro/PJ) pode ser criado à mão.
- `unidades.alias_rh VARCHAR NULL`: nome da unidade como vem da base de RH (ex.: `Juazeiro` ↔ `Revalle Juazeiro`). A sincronização mapeia `unidade_rh` → `unidade_id` por esse alias, com sugestão automática por nome normalizado.
- `locais (id, unidade_id FK, nome, descricao, ativo, UNIQUE(unidade_id, nome))`
- `items.colaborador_id FK NULL`, `items.local_id FK NULL`. `assigned_to` e `cpf` continuam sendo preenchidos (compatibilidade com histórico, termos e relatórios).
- `audit_log (id, ocorrido_em UTC, usuario_id, username, papel, ip, entidade, entidade_id, acao, antes JSON, depois JSON)`, com índices por `(entidade, entidade_id)` e `ocorrido_em`.

**F2**
- `usuarios.email VARCHAR NULL UNIQUE`, `usuarios.totp_secret_cifrado VARCHAR NULL`, `usuarios.totp_ativo TINYINT DEFAULT 0`
- `password_reset_tokens (id, usuario_id FK, token_hash, expira_em, usado_em, criado_ip)`
- `mfa_recovery_codes (id, usuario_id FK, code_hash, usado_em)`

**F3**
- `termo_solicitacoes (id, item_id FK, colaborador_id FK, tipo ENUM('entrega','devolucao','uso'), history_id NULL, modelo_versao, status ENUM('agendado','enviado','assinado','recusado','expirado','cancelado','erro','manual'), solicitado_por, solicitado_em, agendado_para NULL, enviado_em NULL, assinado_em NULL, revsign_documento_id NULL, arquivo_assinado_key NULL, motivo/erro NULL)`

**F4**
- `items` ganha `valor_compra DECIMAL(12,2) NULL`, `data_compra DATE NULL`, `garantia_ate DATE NULL`, `nf_anexo_key VARCHAR NULL`
- `items.status` ENUM ganha `'Em Manutenção'`
- `manutencoes (id, item_id FK, status ENUM('aberta','enviada','em_reparo','retornada','sem_conserto','cancelada'), defeito, assistencia, aberta_por, aberta_em, enviada_em NULL, previsao_retorno NULL, retornada_em NULL, custo DECIMAL(12,2) NULL, anexo_key NULL, observacoes)`
- `transferencias (id, item_id FK, de_unidade, para_unidade, de_local_id NULL, para_local_id NULL, motivo, transferido_por, transferido_em)`
- `baixas_apagamento (id, item_id FK, history_id, dados_apagados ENUM('sim','nao_aplica'), metodo ENUM('formatacao_segura','reset_fabrica','destruicao_fisica','nao_aplica'), apagado_por, apagado_em, comprovante_key NULL)`

**F5**
- `linhas (id, numero UNIQUE, operadora, plano, custo_mensal DECIMAL(10,2), status ENUM('ativa','suspensa','cancelada'), unidade_id FK, item_id FK NULL, colaborador_id FK NULL, iccid NULL, observacoes, criada_em)`
- `linha_eventos (id, linha_id FK, tipo ENUM('ativacao','troca_aparelho','troca_usuario','portabilidade','suspensao','reativacao','cancelamento','edicao'), de_valor, para_valor, ocorrido_em, por_usuario, observacao)`

**F6**
- `inventarios (id, unidade_id FK, status ENUM('aberto','encerrado','cancelado'), iniciado_por, iniciado_em, encerrado_em NULL, observacao)`
- `inventario_leituras (id, inventario_id FK, item_id FK NULL, codigo_lido, lido_por, lido_em, local_id_encontrado NULL, resultado ENUM('ok','local_errado','usuario_diferente','outra_unidade','nao_cadastrado'), observacao, UNIQUE(inventario_id, item_id))`. No encerramento, itens ativos da unidade sem leitura entram no relatório como "não encontrado".

**F8**
- `sistemas (id, nome UNIQUE, descricao, ativo)` com carga inicial: AD, E-mail, VPN, Promax, Biz Deliveries, WMS
- `pacotes_funcao (id, cargo, sistema_id FK, UNIQUE(cargo, sistema_id))`
- `colaborador_acessos (id, colaborador_id FK, sistema_id FK, login, status ENUM('ativo','desativado'), item_id FK NULL, criado_por, criado_em, desativado_por NULL, desativado_em NULL, UNIQUE(colaborador_id, sistema_id, login))`. **Não existe coluna de senha.**
- `acesso_eventos (id, colaborador_id FK, tipo ENUM('admissao','desligamento','mudanca_funcao'), cargo_anterior NULL, cargo_novo NULL, status ENUM('aberto','concluido','cancelado'), aberto_por, aberto_em, concluido_em NULL)`
- `acesso_tarefas (id, evento_id FK, sistema_id FK, acao ENUM('criar','desativar'), status ENUM('pendente','feito','nao_se_aplica'), login NULL, feito_por NULL, feito_em NULL, observacao)`

## 6. Fases

Cada fase lista o que entrega e como é aceita. Endpoints sob `/api/`.

### F0 — Base

1. Migrações versionadas (4.2) e scripts de linha de comando.
2. Serviços compartilhados (4.4): e-mail, auditoria, fonte de colaboradores, agendador, exportações.
3. **Colaboradores**: `GET/POST/PUT /colaboradores`, `GET /colaboradores/{id}`, busca por nome, CPF ou matrícula; `POST /colaboradores/sincronizar` (Gestor); `GET /colaboradores/cargos` (lista distinta, usada em F8). Validação de CPF (dígitos verificadores) com mensagens claras: "CPF inválido", "CPF não encontrado na base de colaboradores", "Base de colaboradores indisponível no momento". E-mail validado por formato; telefone normalizado para dígitos com DDD (10 ou 11 dígitos). Tela `/colaboradores` (lista, filtros, edição de e-mail, telefone e superior, criação manual) e indicação de quando foi a última sincronização.
4. **Unidades**: campo "Nome na base de RH" na tela de unidades.
5. **Locais**: CRUD `/locais` por unidade (Gestor e Técnico), tela dentro de Unidades.
6. **Auditoria**: chamadas a `audit.record` nos fluxos existentes de escrita (itens, empréstimo, devolução, baixa, estorno, periféricos, usuários, unidades) e em todos os módulos novos. `GET /audit` (Gestor) com filtros por entidade, usuário e período; tela `/auditoria`.
7. **Frontend base**:
   - Tokens da paleta (seção 7) em `index.css` e `tailwind.config.js`.
   - `StatusPill` único (pílula arredondada, texto em negrito, cor por status) substituindo `MonoBadge`, `StatusBadge` e o badge local de Periféricos.
   - `Pagination` único: anterior/próxima, campo "ir para a página", seletor 25/50/100 (padrão 25, preferência guardada por tela no `localStorage`).
   - `Button` com prop `loading` (spinner, `disabled`, `aria-busy`), impedindo clique repetido.
   - Helper `upper()` para nomes de unidade/revenda.
   - Página `/items/:id` (detalhe do equipamento) com registro de abas: cada fase adiciona sua aba num arquivo próprio, sem editar as outras.

Aceite: migrações aplicam do zero e sobre o schema atual; sincronização com a base fictícia cria/atualiza/marca desligados sem sobrescrever campos manuais; audit_log registra antes/depois de uma edição de item; componentes com testes.

### F1 — Visual e listagens

1. Aplicar a paleta em todas as telas; painéis grandes com fundo branco (claro) e preto (escuro); caixas menores (indicadores, chips, cabeçalhos de seção) em variações de azul.
2. `StatusPill`, `Pagination` e `Button loading` aplicados em todas as listagens e ações (Estoque, Periféricos, Histórico, Empréstimo, Devolução, Termos, Remoção, Relatório, Usuários, Unidades e telas novas).
3. Unidade e revenda sempre em CAIXA ALTA na exibição e nas exportações (a gravação não muda, para não disparar a cascata de renomeação de `unidade_db.py`).
4. Estoque: colunas Setor e Cargo (do colaborador vinculado); filtros por unidade, tipo, status e colaborador; ordenação por coluna no servidor (`sort` com lista branca de colunas e `dir`); busca existente mantida.
5. Detalhe do equipamento: aba "Dados" com colaborador, setor, cargo e **superior imediato**, local e unidade; aba "Linha do tempo" (consome o endpoint da F4; antes dela, mostra o `history` atual).
6. Responsividade: navegação em gaveta abaixo de 768 px, tabelas viram cartões ou rolam horizontalmente sem quebrar a página, formulários em uma coluna; conferido em 390, 768 e 1440 px.
7. Revisão de consistência visual (espaçamentos, títulos, botões) em todas as telas.

Aceite: nenhuma tela com paginação de 7 itens; status iguais em todas as telas; contraste AA verificado nos dois temas; testes das listagens atualizados.

### F2 — Login

1. **E-mail dos usuários do sistema**: campo na tela de Usuários (Gestor) e no próprio perfil.
2. **Esqueci minha senha**: `POST /auth/forgot` (resposta sempre genérica, com rate limit), e-mail com link `APP_PUBLIC_URL/redefinir-senha?token=...` válido por 30 min e de uso único (token guardado como hash); `POST /auth/reset` define a nova senha (mesmas regras de força atuais), revoga as sessões abertas do usuário e registra na auditoria.
3. **2FA (TOTP)**: cadastro com QR code e confirmação de um código; 10 códigos de recuperação exibidos uma vez; segredo cifrado com `TOTP_ENCRYPTION_KEY`. Login em duas etapas: senha → token temporário de 5 min com `mfa_pendente` → `POST /auth/mfa/verificar` (código TOTP ou de recuperação, com rate limit) → tokens normais. Gestor sem 2FA é levado ao cadastro no primeiro login após o deploy e não usa o sistema antes de concluir. Gestor pode zerar o 2FA de outro usuário (com auditoria). Se o único Gestor perder o aparelho e os códigos de recuperação, o comando `python -m scripts.reset_2fa <usuario>`, rodado no servidor, zera o 2FA dele (e registra na auditoria). Demais papéis ativam no perfil, se quiserem.

Aceite: testes de fluxo completo (token expirado, reutilizado, usuário inexistente sem vazamento, código TOTP errado, recuperação, Gestor obrigado a cadastrar).

### F3 — Termos e RevSign

1. Empréstimo e devolução passam a escolher o **colaborador do cadastro** (busca por nome/CPF). O formulário preenche nome, CPF, cargo, setor e unidade a partir do cadastro; o e-mail é **obrigatório** e validado antes de seguir. Os campos antigos continuam gravados para compatibilidade.
2. `termo_solicitacoes` (seção 5) registra cada pedido de assinatura. O equipamento mostra o **termo vigente**: tipo, versão do modelo, status (pendente/assinado), data da assinatura e quem solicitou.
3. Envio: com `REVSIGN_ENABLED=true`, gerar o termo (.docx atual) e criar o documento no RevSign com o colaborador como signatário; o RevSign envia o e-mail com o link. Com `REVSIGN_ENABLED=false` ou falha, a solicitação fica como `manual` e o fluxo atual (baixar o termo, subir o PDF assinado) continua funcionando.
4. Retorno: webhook `POST /integrations/revsign/webhook` autenticado pelo segredo e, como reserva, consulta periódica (job do agendador) das solicitações `enviado`. Ao assinar: baixar o PDF final, guardar em `termos_assinados` (ou `termos_devolucao_assinados`) e executar a mesma confirmação que o upload manual faz hoje (empréstimo confirmado / devolução confirmada).
5. "Solicitar assinatura" a qualquer momento (`POST /items/{id}/termos`, tipo `uso`), com opção de agendar data e hora (`agendado_para`); o agendador envia na hora marcada. Uma solicitação nova cancela a pendente anterior do mesmo item.
6. Disparo automático: todo empréstimo (entrega a um colaborador, inclusive troca de aparelho ou de usuário) e toda devolução criam a solicitação correspondente sem ação extra do técnico.
7. Histórico por equipamento e por colaborador: quem solicitou, quando enviou, quando assinou (aba "Termos" no detalhe do equipamento e no colaborador).
8. Integração isolada em `app/services/revsign_client.py` com a interface `RevSignClient` (`criar_documento`, `consultar`, `cancelar`, `reenviar`, `baixar_assinado`, `validar_webhook`) e uma implementação falsa para testes. A implementação HTTP real segue o **contrato de integração** que a sessão do RevSign vai entregar (prompt enviado pelo Enzo em 12/11/2026). Enquanto o contrato não chega, tudo funciona com `REVSIGN_ENABLED=false`.

Aceite: com o cliente falso, um empréstimo gera solicitação, o webhook simulado conclui o empréstimo e grava o PDF; envio agendado sai no tique certo; troca de equipamento cancela a pendente; com a integração desligada o fluxo manual atual passa em todos os testes existentes.

### F4 — Ciclo de vida e manutenção

1. Cadastro e edição do equipamento com NF (campo atual), fornecedor (atual), **valor, data da compra, fim da garantia e anexo da NF** (PDF/imagem, até 10 MB, storage `notas_fiscais`), local (da unidade).
2. **Transferência** entre unidades (`POST /items/{id}/transferir`, Gestor e Técnico): destino, local opcional, motivo; atualiza `revenda` e `local_id`; registra em `transferencias`, `history` (operação `Edição` com detalhe "Transferência X → Y") e auditoria.
3. **Manutenção**: abrir ordem (só para equipamento sem colaborador; status do item vai para `Em Manutenção`), registrar envio à assistência, previsão, retorno com custo e anexo (orçamento/NF), ou "sem conserto" (sugere a baixa). Ao retornar, o item volta a `Disponível`. Listagem `/manutencoes` com filtros e total de custo por equipamento; aba "Manutenções" no detalhe.
4. **Baixa com apagamento de dados**: para Notebook, Desktop, Celular, Tablet e HD, a baixa exige registrar se os dados foram apagados, o método, quem fez e quando (comprovante opcional); para os demais tipos, `nao_aplica`.
5. **Linha do tempo**: `GET /items/{id}/linha-do-tempo` junta cadastro, empréstimos, devoluções, termos, manutenções, transferências, inventários, linha telefônica e baixa, em ordem cronológica.

Aceite: item em manutenção não aparece como disponível para empréstimo; custo acumulado correto; linha do tempo completa com dados de todas as tabelas; baixa de notebook sem registro de apagamento é recusada.

### F5 — Linhas telefônicas

1. CRUD de linhas (`/linhas`), com número validado (DDD + 8/9 dígitos), operadora, plano, custo mensal, status, unidade e vínculos opcionais com aparelho (Celular/Tablet) e colaborador.
2. Eventos registrados automaticamente a cada troca de aparelho ou de usuário, portabilidade, suspensão, reativação e cancelamento, com histórico na tela da linha e aba "Linha" no detalhe do aparelho.
3. Relatório de custo mensal por unidade (só linhas ativas), com exportação Excel/PDF.

Aceite: trocar o aparelho de uma linha gera evento; relatório soma corretamente por unidade; número duplicado é recusado.

### F6 — Inventário físico

1. **Etiquetas**: seleção de equipamentos na listagem (ou todos de uma unidade) → folha A4 para impressão (`@media print`) com QR code, número do patrimônio e tipo/modelo. Padrão: 3 colunas × 8 linhas (70 × 37 mm, 24 por folha), com layout configurável (colunas, linhas, margens) e posição inicial, para aproveitar folhas já usadas. O QR contém só `APP_PUBLIC_URL/e/{id}`; nenhum dado pessoal.
2. `/e/{id}` abre o detalhe do equipamento (com login).
3. **Inventário**: Gestor ou Técnico abre um inventário por unidade; no celular, a tela de leitura usa a câmera (`@zxing/browser`, carregado sob demanda; exige HTTPS) ou digitação/leitor USB. Cada leitura confere unidade, local e colaborador e marca o resultado; o técnico pode informar o local onde encontrou. Encerrar o inventário calcula os "não encontrados".
4. **Relatório** por inventário (tela e exportação Excel/PDF): conferidos, não encontrados, local errado, usuário diferente, de outra unidade e não cadastrados.

Aceite: leitura duplicada não conta duas vezes; encerramento lista os não lidos; câmera negada mostra a entrada manual; impressão sai com o layout configurado.

### F7 — Painel, alertas e exportações

1. `GET /painel/resumo`: equipamentos por unidade e por status, em manutenção, termos pendentes, linhas ativas e custo mensal; números calculados no servidor (sem o limite de 500 do cliente).
2. `GET /painel/alertas`, calculado na hora: garantia vencendo (≤ `ALERTA_GARANTIA_DIAS`), termo pendente (> `ALERTA_TERMO_PENDENTE_DIAS`), devolução pendente (`Pendente Devolução` > `ALERTA_DEVOLUCAO_PENDENTE_DIAS`), manutenção parada (> `ALERTA_MANUTENCAO_DIAS`), colaborador desligado ainda com equipamento. Cada alerta leva ao registro.
3. Painel inicial redesenhado com os indicadores e a lista de alertas.
4. Exportação Excel e PDF em: Estoque (com os filtros da tela), Histórico, Manutenções, Linhas, Inventário, Termos pendentes e Relatório mensal (o CSV atual continua).

Aceite: números do painel batem com consultas diretas; cada alerta aparece com o dado de teste certo e some quando resolvido; arquivos exportados abrem e respeitam os filtros.

### F8 — Acessos

1. **Catálogo de sistemas** (`/acessos/sistemas`, Gestor): cadastrar, editar, desativar.
2. **Pacotes por função**: para cada cargo (lista vinda dos colaboradores), os sistemas padrão.
3. **Ficha de acessos** do colaborador: sistema, login, status (ativo/desativado), equipamento onde está instalado (opcional), quem criou/desativou e quando. O sistema nunca pede nem guarda senha; os formulários deixam isso explícito.
4. **Eventos registrados à mão** (Gestor e Técnico):
   - Admissão: gera tarefas "criar" a partir do pacote do cargo; o TI marca cada uma como feita (informando o login) ou "não se aplica".
   - Desligamento: gera tarefas "desativar" para todos os acessos ativos e a lista de equipamentos com o colaborador (atalho para a devolução); o evento só conclui quando todas as tarefas estiverem resolvidas e nenhum equipamento continuar com ele.
   - Mudança de função: compara o pacote do cargo antigo com o novo, gera "criar" para o que entra e "desativar" para o que sai (o TI pode marcar "manter", que vira "não se aplica" com observação).
5. Acessos é um módulo próprio no menu, usando o mesmo cadastro de colaboradores; aba "Acessos" no detalhe do colaborador e no do equipamento.

Aceite: admissão de um cargo com pacote gera as tarefas certas; desligamento não conclui com equipamento pendente; mudança de função gera a diferença correta; nenhuma tabela ou payload aceita campo de senha.

## 7. Paleta e tokens

Base pedida: `#0131CD`, `#FFFFFF`, `#B9FC01`. Contrastes calculados (WCAG): `#0131CD` sobre branco ≈ 9:1; `#B9FC01` sobre branco ≈ 1,2:1 (não serve como texto); `#B9FC01` sobre `#111` ≈ 15:1; `#B9FC01` sobre `#0131CD` ≈ 7:1; `#0131CD` sobre fundo quase preto ≈ 2,2:1 (não serve como texto no escuro).

Regras:
- Claro: fundo de página branco/cinza muito claro, painéis brancos; `--primary` = `#0131CD` com texto branco; links e ícones ativos em `#0131CD`; anel de foco azul.
- Escuro: painéis pretos/quase pretos; botão primário continua `#0131CD` com texto branco; texto e ícones azuis usam um tom claro derivado (ex.: `#5C7CFF`, ≥ 4,5:1 sobre o fundo escuro); anel de foco verde-limão.
- Verde-limão `#B9FC01` só como preenchimento com texto escuro (`#111`) ou sobre o azul: destaque de item ativo na navegação, indicador "novo", seleção, chips de destaque.
- Caixas menores em variações de azul (fundos com o azul em baixa opacidade, calculados por tema).
- Status com cor própria e contraste AA nos dois temas: Disponível (verde), Em uso (azul), Pendente (âmbar), Pendente Devolução (roxo), Em Manutenção (laranja), Baixado (cinza), Erro/Recusado (vermelho). Texto sempre em negrito.
- O guard test de tokens (`tokens.test.ts`) é atualizado com a nova lista.

## 8. Segurança e LGPD

- Base de RH só em leitura, sessão com `default_transaction_read_only`; recomendação ao Enzo de criar um usuário só-leitura para o Patrimônio (script do RevSign) e trocar a senha do `postgres`, que foi exposta no chat.
- Não guardar data de nascimento; CPF mascarado em logs e na auditoria; QR sem dados pessoais.
- Tokens de redefinição e códigos de recuperação só como hash; segredo TOTP cifrado.
- Webhook do RevSign validado por segredo; payload tratado como dado não confiável.
- Nenhuma senha de sistema de terceiros no módulo de Acessos.
- Segredos só no `.env` do servidor; `.env.example` com valores fictícios.

## 9. Testes

- Backend: pytest no MySQL de teste (`docker-compose.test.yml`, porta 3307, derrubado com `down -v` no fim). Fonte de colaboradores simulada num database temporário criado no `dev-postgres` (localhost:5433) com uma tabela de colunas iguais à real, apagado (`DROP DATABASE`) no fim. SMTP com backend em memória; RevSign com o cliente falso. Testes de migração (do zero e sobre o schema atual, ida e volta).
- Frontend: Vitest + Testing Library + MSW para cada tela e componente novo; testes atuais ajustados quando o visual mudar de propósito.
- CI atual (`.github/workflows/ci.yml`) continua obrigatório; o job do backend passa a rodar as migrações.
- Teste de fumaça final: subir backend e frontend localmente, percorrer login com 2FA, empréstimo com termo (cliente falso), manutenção, inventário e painel.

## 10. Execução

- Branch única `feat/patrimonio-melhorias`.
- **F0 primeiro**, sozinha: é a base de todas.
- Depois, **F1, F2, F3, F4, F5, F6 e F8 em paralelo**, cada uma em worktree própria a partir da F0, com um agente implementador por tarefa e revisão de conformidade ao spec e de qualidade a cada tarefa. Pontos de conflito previstos e como evitar: rotas (`App.tsx`), menu (`navigation.ts`) e registro de routers (`main.py`) recebem só linhas novas; abas do detalhe ficam em arquivos próprios; migrações em faixas reservadas.
- **F7 por último**, depois de F3, F4 e F5 integradas.
- Integração: merge de cada fase na branch, suíte completa, build e teste de fumaça; revisão final do conjunto.
- A parte do RevSign é feita pela sessão do RevSign, no repositório dele. A F3 não depende dela para ser concluída (integração desligada por padrão); quando o contrato chegar, a implementação HTTP do `RevSignClient` é completada e testada em homologação.

## 11. Riscos e pendências

| Risco/pendência | Tratamento |
|---|---|
| Contrato do RevSign ainda não existe | `RevSignClient` + `REVSIGN_ENABLED=false`; fluxo manual continua |
| Nomes reais das colunas de `"colaboradores-revalle"` | Descoberta por nome + `COLABORADORES_COLUMNS`; validar no primeiro deploy pela tela de sincronização |
| Unidades da base de RH com nomes diferentes das do Patrimônio | `unidades.alias_rh` com sugestão automática |
| Câmera exige HTTPS | Produção já atrás de proxy com TLS; confirmar antes de usar o inventário no celular |
| Modelo exato da folha de etiquetas | Layout configurável; padrão 3×8 (70 × 37 mm) a confirmar com a folha comprada |
| Volume de mudanças no frontend | Fases com abas e páginas em arquivos próprios; revisão a cada tarefa |
