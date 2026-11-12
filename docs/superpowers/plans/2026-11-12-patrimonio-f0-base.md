# Fase 0 — Base — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Criar a base que todas as outras fases usam: migrações versionadas, configuração nova, e-mail, auditoria, colaboradores sincronizados do Postgres de RH, agendador, locais, exportações e os componentes de frontend (paleta, StatusPill, Pagination, Button com loading, página de detalhe com abas).

**Architecture:** Módulos novos isolados (router + manager + schemas próprios) sobre o SQL cru existente (`get_cursor`). Migrações em arquivos `.sql` numerados aplicados no startup com `GET_LOCK`. No frontend, componentes compartilhados em `components/ui` e um registro de abas por arquivo (`import.meta.glob`) para o detalhe do equipamento.

**Tech Stack:** FastAPI 0.115, PyMySQL + DBUtils, Pydantic 2.9, psycopg 3.2 (binary), openpyxl 3.1, fpdf2 2.8; React 18 + TS 5.6 + Tailwind 3.4 + TanStack Query 5 + Vitest/MSW.

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md` (seções 4, 5 "F0", 6 "F0", 7, 8, 9).

## Global Constraints

- Mudanças de banco só aditivas e sempre por migração em `backend/migrations/NNNN_*.sql` + `NNNN_*.down.sql`; faixa da F0: `0001–0099`.
- Base de RH: somente leitura (`default_transaction_read_only=on`); nunca ler nem guardar data de nascimento.
- CPF mascarado (`***.456.789-**`) em auditoria e logs.
- Segredos só via `.env`; `.env.example` só com valores fictícios.
- Textos de interface e mensagens de erro em português.
- Nenhum fluxo atual pode quebrar: a suíte atual (235 testes backend, 174 frontend) continua verde.
- Paleta base: `#0131CD`, `#FFFFFF`, `#B9FC01` (regras na seção 7 do spec).
- Testes do backend: `C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh <worktree> <sufixo> [args]` (container python:3.12, MySQL de teste na 3307, Postgres `dev-postgres` na 5433). Frontend: `npx vitest run` e `npx tsc --noEmit` em `frontend/`.
- Commits pequenos, mensagem em português `tipo(escopo): descrição`, terminando com `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus

- Base de RH fora do ar ou URL vazia → sincronização registra erro em `colaboradores_sync` e a API responde 503 "Base de colaboradores indisponível no momento", sem derrubar o app (teste na Task 5).
- Colaborador editado à mão (e-mail/telefone) e sincronizado de novo → o valor manual permanece (teste na Task 5).
- Dois workers subindo ao mesmo tempo → migrações aplicadas uma vez só (GET_LOCK; teste chamando `aplicar_pendentes()` duas vezes na Task 1).
- Falha ao gravar auditoria sem cursor compartilhado → a operação do usuário não falha (teste na Task 3).
- Usuário digita página fora do intervalo no `Pagination` (0, 999, texto) → vai para a página válida mais próxima (teste na Task 10).

---

### Task 1: Migrações versionadas, configuração nova e infraestrutura de teste

**Files:**
- Create: `backend/app/db/migrations.py`, `backend/scripts/migrate.py`, `backend/migrations/README.md`, `backend/tests/test_migrations.py`
- Modify: `backend/app/core/config.py` (Settings), `backend/.env.example`, `backend/app/main.py` (startup), `backend/tests/conftest.py`, `backend/requirements.txt`

**Interfaces:**
- Produces:
  - `app.db.migrations.MIGRATIONS_DIR: Path` (= `backend/migrations`)
  - `@dataclass Migracao(versao: int, nome: str, up_path: Path, down_path: Path | None)`
  - `listar_migracoes() -> list[Migracao]` (ordenadas; nome do arquivo `NNNN_descricao.sql`; `.down.sql` é o par)
  - `garantir_baseline() -> None` (instancia `InventoryDBManager`, `UnidadeDBManager`, `TokenDBManager` para as tabelas legadas existirem antes das migrações)
  - `aplicar_pendentes() -> list[int]` (versões aplicadas nesta chamada; dentro de `SELECT GET_LOCK('patrimonio_migrations', 30)` / `RELEASE_LOCK`; cria `schema_migrations(versao INT PRIMARY KEY, nome VARCHAR(200), aplicada_em DATETIME)` se não existir)
  - `reverter_ultima() -> int | None`
  - `executar_sql(cur, texto: str) -> None` (divide em comandos por `;` no fim da linha, ignorando linhas `--`)
  - Settings novas (todas com padrão seguro): `APP_PUBLIC_URL="http://localhost:5173"`, `MAIL_BACKEND="smtp"` (`smtp`|`memory`), `SMTP_HOST=""`, `SMTP_PORT=465`, `SMTP_USER=""`, `SMTP_PASS=""`, `SMTP_FROM=""`, `COLABORADORES_DATABASE_URL=""`, `COLABORADORES_TABLE="colaboradores-revalle"`, `COLABORADORES_SCHEMA="public"`, `COLABORADORES_COLUMNS=""`, `COLABORADORES_SYNC_HORA="05:00"`, `SCHEDULER_ENABLED=True`, `TOTP_ENCRYPTION_KEY=""`, `REVSIGN_ENABLED=False`, `REVSIGN_BASE_URL=""`, `REVSIGN_API_TOKEN=""`, `REVSIGN_WEBHOOK_SECRET=""`, `ALERTA_GARANTIA_DIAS=30`, `ALERTA_TERMO_PENDENTE_DIAS=3`, `ALERTA_DEVOLUCAO_PENDENTE_DIAS=7`, `ALERTA_MANUTENCAO_DIAS=15`. Propriedade `smtp_from_effective` (= `SMTP_FROM or SMTP_USER`).
  - conftest: `limpar_banco` passa a truncar **todas** as tabelas do schema exceto `schema_migrations` e `scheduler_execucoes` (lista vinda de `information_schema.tables`), e `_test_schema` chama `aplicar_pendentes()` depois dos managers. `pytest_configure` define `MAIL_BACKEND=memory` e `SCHEDULER_ENABLED=false`.

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_migrations.py` (marcador `integration`):
  - `test_listar_migracoes_ordena_e_pareia_down` usa `tmp_path` com arquivos `0002_b.sql`, `0001_a.sql`, `0001_a.down.sql` (monkeypatch de `MIGRATIONS_DIR`) e afirma `[m.versao for m in ...] == [1, 2]`, `down_path` de 1 preenchido e de 2 `None`.
  - `test_aplicar_pendentes_e_idempotente`: chama `aplicar_pendentes()` duas vezes; a segunda retorna `[]`; `schema_migrations` contém todas as versões de `listar_migracoes()`.
  - `test_reverter_ultima_e_reaplicar`: `v = reverter_ultima()` remove `v` de `schema_migrations`; `aplicar_pendentes() == [v]`.
  - `test_executar_sql_divide_comandos`: texto com dois `CREATE TABLE` e um comentário cria as duas tabelas temporárias.
  - Teste unitário (`unit`) `test_settings_padroes_novos` afirma `settings.ALERTA_GARANTIA_DIAS == 30` e `settings.smtp_from_effective` usa `SMTP_USER` quando `SMTP_FROM` vazio (instanciar `Settings(_env_file=None, DB_HOST=..., ...)`).
- [ ] **Step 2: Rodar e ver falhar:** `run-backend-tests.sh <wt> f0 tests/test_migrations.py -q` → FAIL (módulo inexistente).
- [ ] **Step 3: Implementar** `app/db/migrations.py`, as Settings novas, o `scripts/migrate.py` (`python -m scripts.migrate [aplicar|pendentes|reverter]`), o README das migrações (convenção de nomes, faixas por fase, par `.down.sql` obrigatório) e o `.env.example` (cada variável nova com comentário; `SMTP_HOST=mail.revalle.com.br`, `SMTP_PORT=465`, `SMTP_USER=naoresponda@revalle.com.br`, `SMTP_PASS=<defina no servidor>`, `COLABORADORES_DATABASE_URL=postgres://<usuario_leitura>:<senha>@agente_revalle_postgres:5432/agente_revalle`). Em `main.py`, um handler `@app.on_event("startup")` chama `garantir_baseline()` e `aplicar_pendentes()` dentro de `try/except` com `logger.exception` (banco fora do ar não impede o boot, igual ao comportamento atual). Adicionar a `requirements.txt`: `psycopg[binary]==3.2.3`, `pyotp==2.9.0`, `openpyxl==3.1.5`, `fpdf2==2.8.1` (a imagem de teste já os tem).
- [ ] **Step 4: Rodar a suíte inteira:** `run-backend-tests.sh <wt> f0 -q -m "not mudanca_esperada"` → todos passam (235 antigos + novos).
- [ ] **Step 5: Commit** `feat(migracoes): executor de migrações versionadas e configuração nova`.

### Task 2: Serviço de e-mail

**Files:**
- Create: `backend/app/services/__init__.py`, `backend/app/services/mail.py`, `backend/app/templates/email/base.html`, `backend/tests/test_mail.py`

**Interfaces:**
- Produces:
  - `class MailError(Exception)`
  - `@dataclass MailMessage(to: list[str], subject: str, html: str, text: str)`
  - `outbox: list[MailMessage]` (só usado com `MAIL_BACKEND=memory`) e `limpar_outbox() -> None`
  - `send_mail(to: str | list[str], subject: str, html: str, text: str | None = None) -> None` — `smtplib.SMTP_SSL` quando `SMTP_PORT == 465`, senão `SMTP` + `starttls()`; timeout 15 s; `From` = `settings.smtp_from_effective`; texto alternativo gerado removendo tags quando `text` é `None`; qualquer `smtplib.SMTPException`/`OSError` vira `MailError("Não foi possível enviar o e-mail.")`; nunca loga corpo nem senha. Sem `SMTP_HOST` configurado → `MailError("Envio de e-mail não configurado.")`.
  - `render_email(titulo: str, corpo_html: str, acao_url: str | None = None, acao_rotulo: str | None = None) -> str` — envolve o corpo no `base.html` (cabeçalho "Revalle · Patrimônio" em `#0131CD`, botão de ação, rodapé "Mensagem automática, não responda.").
- [ ] **Step 1: Testes** (`unit`): `test_memory_backend_guarda_mensagem` (com `MAIL_BACKEND=memory`, `send_mail("a@b.com", "Assunto", "<p>Oi</p>")` → `outbox[-1].to == ["a@b.com"]` e `text == "Oi"`); `test_smtp_ssl_na_porta_465` (monkeypatch de `smtplib.SMTP_SSL` com um falso que registra `login(SMTP_USER, SMTP_PASS)` e `send_message`); `test_starttls_em_outra_porta`; `test_falha_smtp_vira_mail_error`; `test_sem_host_configurado_recusa`; `test_render_email_inclui_botao` (contém `acao_url` e `acao_rotulo`, escapa `<` no título).
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `tests/test_mail.py` → PASS.
- [ ] **Step 5: Commit** `feat(email): serviço de envio SMTP com backend em memória para testes`.

### Task 3: Auditoria (infraestrutura e consulta)

**Files:**
- Create: `backend/migrations/0005_audit_log.sql` (+ `.down.sql`), `backend/app/services/audit.py`, `backend/app/db/audit_db.py`, `backend/app/schemas/audit.py`, `backend/app/routers/audit.py`, `backend/tests/test_audit.py`
- Modify: `backend/app/dependencies.py` (dependência de contexto), `backend/app/main.py` (router)

**Interfaces:**
- Produces:
  - Tabela `audit_log (id BIGINT PK AI, ocorrido_em DATETIME(3) NOT NULL /*UTC*/, usuario_id INT NULL, username VARCHAR(100) NULL, papel VARCHAR(50) NULL, ip VARCHAR(45) NULL, entidade VARCHAR(50) NOT NULL, entidade_id VARCHAR(64) NULL, acao VARCHAR(50) NOT NULL, antes JSON NULL, depois JSON NULL)`, índices `(entidade, entidade_id)` e `(ocorrido_em)`.
  - `@dataclass AuditContext(usuario_id: int | None, username: str | None, papel: str | None, ip: str | None)`; `AuditContext.sistema() -> AuditContext` (username `"sistema"`, para jobs).
  - `app.dependencies.get_audit_context(request: Request, user: CurrentUser = Depends(get_current_user)) -> AuditContext` — IP: `request.client.host`; usa o primeiro IP de `X-Forwarded-For` só quando `request.client.host` é loopback ou rede privada (proxy do compose).
  - `audit.mascarar(valor: dict | None) -> dict | None` (chave `cpf` vira `***.XXX.XXX-**` preservando os dígitos 4–9; recursivo em dicts aninhados)
  - `audit.diff(antes: dict | None, depois: dict | None) -> tuple[dict | None, dict | None]` (só as chaves que mudaram; `datetime`/`date`/`Decimal` serializados como texto)
  - `audit.record(ctx: AuditContext, entidade: str, entidade_id: int | str | None, acao: str, antes: dict | None = None, depois: dict | None = None, cur=None) -> None` — com `cur` grava na mesma transação e propaga erro; sem `cur` abre `get_cursor()` próprio e, em erro, faz `logger.exception` e **não** propaga.
  - `AuditDBManager.listar(*, entidade: str | None, entidade_id: str | None, username: str | None, de: date | None, ate: date | None, limit: int, offset: int) -> tuple[list[dict], int]`
  - `GET /api/audit` (`gestor_only`) → `Paginated[AuditEntry]` com os filtros acima (`de`/`ate` em `AAAA-MM-DD`).
  - Valores de `entidade` usados no sistema: `item`, `emprestimo`, `devolucao`, `periferico`, `usuario`, `unidade`, `colaborador`, `local`, `historico` (as fases seguintes acrescentam as suas). Valores de `acao`: `criar`, `editar`, `excluir`, `reativar`, `desativar`, `confirmar`, `estornar`, `vincular`, `desvincular`, `substituir`, `sincronizar`.
- [ ] **Step 1: Testes:** `test_mascarar_cpf` (`{"cpf": "12345678909"}` → `{"cpf": "***.456.789-**"}`); `test_diff_so_mudancas`; `test_record_grava_linha` (lê `audit_log` via `executar_sql`); `test_record_sem_cursor_nao_propaga` (monkeypatch de `get_cursor` que levanta → nenhuma exceção); `test_get_audit_exige_gestor` (403 para Técnico); `test_get_audit_filtra_por_entidade`; `test_ip_x_forwarded_for_so_de_proxy_confiavel` (cliente `testclient` não é privado → ignora o header; simular `request.client.host="172.18.0.5"` → usa o header).
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `tests/test_audit.py` → PASS.
- [ ] **Step 5: Commit** `feat(auditoria): registro de alterações com consulta para o Gestor`.

### Task 4: Auditoria nos fluxos existentes e filtro do histórico por item

**Files:**
- Modify: `backend/app/routers/items.py`, `loans.py`, `peripherals.py`, `users.py`, `unidades.py`, `history.py`; `backend/app/db/inventory_manager_db.py` (`list_history` ganha `item_id`)
- Test: `backend/tests/test_audit_fluxos.py`

**Interfaces:**
- Consumes: `get_audit_context`, `audit.record`, `audit.diff` (Task 3).
- Produces: `GET /api/history?item_id=<id>` (filtro opcional, combinável com `search`); `InventoryDBManager.list_history(*, search=None, limit, offset, item_id: int | None = None)`.
- Cada endpoint de escrita grava auditoria **depois** do sucesso, com `antes`/`depois` vindos do registro lido antes e depois da operação: itens (`criar`, `editar` com diff, `excluir` com motivo), empréstimo (`criar` na emissão, `confirmar`), devolução (`criar`, `confirmar`), estorno (`estornar`, entidade `historico`), periféricos (`criar`, `editar`, `vincular`, `desvincular`, `substituir`), usuários (`criar`, `excluir`, `editar` só com `{"senha": "alterada"}` — nunca o hash), unidades (`criar`, `editar`, `desativar`, `reativar`).
- [ ] **Step 1: Testes:** um teste por grupo afirmando a linha em `audit_log` (`entidade`, `acao`, `username` do token) — `test_editar_item_grava_diff` (só os campos alterados em `antes`/`depois`), `test_emprestimo_grava_auditoria_com_cpf_mascarado`, `test_trocar_senha_nao_grava_hash`, `test_unidade_desativar_grava`, `test_history_filtra_por_item_id`.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar** (parâmetro `ctx: AuditContext = Depends(get_audit_context)` nos endpoints; nada muda nas respostas).
- [ ] **Step 4: Rodar a suíte inteira** → verde.
- [ ] **Step 5: Commit** `feat(auditoria): registra alterações em itens, empréstimos, periféricos, usuários e unidades`.

### Task 5: Fonte de colaboradores (Postgres de RH)

**Files:**
- Create: `backend/app/services/colaboradores_source.py`, `backend/tests/test_colaboradores_source.py`, `backend/tests/fixtures_colaboradores.py` (fixture da base fictícia)
- Modify: `backend/tests/conftest.py` (registrar o plugin de fixtures `pytest_plugins = ["tests.fixtures_colaboradores"]`)

**Interfaces:**
- Produces:
  - `class ColaboradoresSourceError(Exception)`; `class ColaboradoresIndisponivel(ColaboradoresSourceError)` (mensagem "Base de colaboradores indisponível no momento.")
  - `CAMPOS = ("cpf", "nome", "cargo", "setor", "unidade", "matricula", "email", "telefone", "situacao", "gestor", "demissao")`
  - `SINONIMOS: dict[str, tuple[str, ...]]` — cpf: `cpf, cpf_completo, nr_cpf, num_cpf, documento`; nome: `nome, nome_completo, colaborador, nome_colaborador, funcionario`; cargo: `cargo, cargo_nome, funcao, nome_cargo`; setor: `setor, setor_nome, departamento, area`; unidade: `unidade, unidade_nome, filial, empresa, revenda, loja`; matricula: `matricula, matricula_rh, registro, codigo, id_colaborador`; email: `email, e_mail, email_corp, email_corporativo, email_pessoal`; telefone: `telefone, celular, fone, whatsapp, telefone_celular`; situacao: `situacao, status, status_rh, situacao_rh, ativo`; gestor: `gestor, superior, gestor_imediato, lider, nome_gestor, superior_imediato`; demissao: `demissao, data_demissao, dt_demissao, desligamento, data_desligamento`.
  - `normalizar_nome_coluna(nome: str) -> str` (minúsculas, sem acento, não alfanumérico → `_`)
  - `descobrir_colunas(colunas_reais: list[str], override: str = "") -> dict[str, str]` — para cada campo, primeira coluna real cujo nome normalizado está nos sinônimos (na ordem); `override` `"campo=coluna,campo=coluna"` tem prioridade; levanta `ColaboradoresSourceError("Base de colaboradores: coluna não encontrada para cpf")` se faltar `cpf` ou `nome`.
  - `normalizar_situacao(valor: object, demissao: object = None) -> Literal["ativo", "afastado", "desligado"]` — `demissao` preenchida → `desligado`; texto contendo `deslig`, `demit`, `inativ`, `rescis` ou valores `false`/`0`/`n`/`nao` → `desligado`; contendo `afast`, `licen` → `afastado`; senão `ativo`.
  - `@dataclass ColaboradorRH(cpf: str, nome: str, cargo: str | None, setor: str | None, unidade: str | None, matricula: str | None, email: str | None, telefone: str | None, situacao: str, gestor_nome: str | None)`
  - `ler_todos() -> list[ColaboradorRH]` — `psycopg.connect(settings.COLABORADORES_DATABASE_URL, connect_timeout=10, options="-c default_transaction_read_only=on")`; colunas via `information_schema.columns` (`table_schema`/`table_name` dos settings); SELECT montado com `psycopg.sql.Identifier`; CPF só dígitos com `zfill(11)`; linhas com CPF de tamanho ≠ 11 ou nome vazio são ignoradas; URL vazia ou erro de conexão → `ColaboradoresIndisponivel`.
  - `buscar_por_cpf(cpf: str) -> ColaboradorRH | None` (mesma conexão, `WHERE` comparando só dígitos).
  - Fixture `base_rh` (em `fixtures_colaboradores.py`): cria o database `colab_test_<COLAB_TEST_DB_SUFFIX>` no Postgres de `COLAB_TEST_PG_ADMIN_URL`, cria a tabela `"colaboradores-revalle"` com colunas `"CPF" text, "Nome Completo" text, "Data Nascimento" date, "Cargo" text, "Setor" text, "Unidade" text, "Matrícula" text, "E-mail" text, "Celular" text, "Situação" text, "Gestor" text`, insere 4 pessoas (uma desligada, uma afastada, CPF com máscara em uma), aponta `settings.COLABORADORES_DATABASE_URL` para ele (monkeypatch) e devolve um helper `executar(sql)`; no teardown `DROP DATABASE ... WITH (FORCE)`. Sem `COLAB_TEST_PG_ADMIN_URL` → `pytest.skip`.
- [ ] **Step 1: Testes:** `test_descobrir_colunas_por_sinonimo` (`["CPF", "Nome Completo", "E-mail", "Situação"]` → `{"cpf": "CPF", "nome": "Nome Completo", "email": "E-mail", "situacao": "Situação"}`); `test_override_tem_prioridade`; `test_sem_cpf_levanta`; `test_normalizar_situacao` (tabela de casos); `test_ler_todos_da_base_ficticia` (4 registros, CPF mascarado normalizado, situação certa, data de nascimento não aparece no dataclass); `test_sessao_somente_leitura` (com a conexão do módulo, `INSERT` levanta erro de read-only); `test_url_vazia_indisponivel`.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** → PASS.
- [ ] **Step 5: Commit** `feat(colaboradores): leitura somente leitura da base de RH com descoberta de colunas`.

### Task 6: Cadastro de colaboradores (banco, sincronização e API)

**Files:**
- Create: `backend/migrations/0001_colaboradores.sql` (+ down), `backend/migrations/0002_unidades_alias_rh.sql` (+ down), `backend/app/db/colaborador_db.py`, `backend/app/schemas/colaboradores.py`, `backend/app/routers/colaboradores.py`, `backend/tests/test_colaboradores.py`
- Modify: `backend/app/schemas/validators.py` (novos validadores), `backend/app/schemas/unidades.py` (`alias_rh`), `backend/app/dependencies.py` (`get_colaborador_db`), `backend/app/main.py` (router, rótulos `email`/`telefone`/`superior_id`)

**Interfaces:**
- Consumes: `ColaboradorRH`, `ler_todos`, `buscar_por_cpf`, `ColaboradoresIndisponivel` (Task 5); `audit.record` (Task 3).
- Produces:
  - Tabelas: `colaboradores (id INT PK AI, cpf CHAR(11) NOT NULL UNIQUE, nome VARCHAR(200) NOT NULL, matricula VARCHAR(50), cargo VARCHAR(150), setor VARCHAR(150), unidade_rh VARCHAR(150), unidade_id INT NULL FK unidades(id), situacao ENUM('ativo','afastado','desligado') NOT NULL DEFAULT 'ativo', email VARCHAR(200), telefone VARCHAR(20), email_manual TINYINT(1) DEFAULT 0, telefone_manual TINYINT(1) DEFAULT 0, superior_id INT NULL FK colaboradores(id) ON DELETE SET NULL, gestor_nome_rh VARCHAR(200), origem ENUM('rh','manual') NOT NULL DEFAULT 'rh', sincronizado_em DATETIME NULL, criado_em DATETIME NOT NULL, atualizado_em DATETIME NOT NULL)` com índices em `nome`, `situacao`, `unidade_id`; `colaboradores_sync (id INT PK AI, iniciado_em DATETIME, concluido_em DATETIME NULL, status ENUM('ok','erro','em_andamento'), total_base INT, criados INT, atualizados INT, desligados INT, erro VARCHAR(500))`; `unidades.alias_rh VARCHAR(150) NULL`.
  - Validadores em `validators.py`: `validar_email(email: str | None) -> str | None` (minúsculas, regex `^[^@\s]+@[^@\s]+\.[^@\s]+$`, erro "E-mail inválido."); `normalizar_telefone(telefone: str | None) -> str | None` (só dígitos; remove `55` inicial quando sobram 12–13; aceita 10 ou 11 dígitos com DDD 11–99; erro "Telefone inválido: informe DDD + número.").
  - `@dataclass ResumoSync(total_base: int, criados: int, atualizados: int, desligados: int)`
  - `ColaboradorDBManager`: `sincronizar(registros: list[ColaboradorRH]) -> ResumoSync` (upsert por CPF; atualiza nome, matrícula, cargo, setor, unidade_rh, situação, gestor_nome_rh e `sincronizado_em`; e-mail/telefone só quando `*_manual = 0`; `unidade_id` resolvido por `unidades.alias_rh` igual (sem acento/caixa) a `unidade_rh` ou, na falta, por `unidades.nome` contendo `unidade_rh`; colaboradores `origem='rh'` ausentes da base viram `desligado`; segunda passada: `superior_id` vazio + `gestor_nome_rh` igual ao nome de exatamente um colaborador → preenche); `registrar_sync_inicio() -> int`, `registrar_sync_fim(sync_id: int, resumo: ResumoSync | None, erro: str | None) -> None`, `ultima_sync() -> dict | None`; `listar(*, busca: str | None, situacao: str | None, unidade_id: int | None, limit: int, offset: int) -> tuple[list[dict], int]` (busca por nome, CPF só dígitos ou matrícula; inclui `unidade_nome` e `superior_nome` via JOIN); `obter(id: int) -> dict | None`; `obter_por_cpf(cpf: str) -> dict | None`; `criar_manual(dados: dict) -> int`; `atualizar(id: int, dados: dict) -> None` (e-mail/telefone marcam `*_manual = 1`; `nome`, `cargo`, `setor`, `unidade_id`, `situacao` só para `origem='manual'`, senão `ValueError("Campo vem da base de RH e não pode ser editado aqui.")`; `superior_id` igual ao próprio id → `ValueError("O colaborador não pode ser superior de si mesmo.")`); `cargos() -> list[str]` (distintos, não vazios, ordenados).
  - `executar_sincronizacao(db: ColaboradorDBManager, ctx: AuditContext) -> ResumoSync` (função de serviço em `colaborador_db.py`: registra início, `ler_todos()`, `sincronizar`, registra fim e auditoria `colaborador/sincronizar`; erro → registra fim com erro e relança `ColaboradoresIndisponivel`).
  - Schemas: `ColaboradorResponse` (todos os campos da tabela exceto `gestor_nome_rh`, mais `unidade_nome: str | None`, `superior_nome: str | None`), `ColaboradorCreate(cpf, nome, cargo?, setor?, unidade_id?, email?, telefone?, superior_id?, situacao: Literal[...] = "ativo")`, `ColaboradorUpdate` (todos opcionais), `SyncStatus(ultima: dict | None, configurado: bool)`.
  - Router `/api/colaboradores`: `GET ""` (autenticado; `busca`, `situacao`, `unidade_id`, `limit` padrão 25, `offset`) → `Paginated[ColaboradorResponse]`; `GET /cargos` → `list[str]`; `GET /sincronizacao` → `SyncStatus`; `POST /sincronizar` (`gestor_only`) → `ResumoSync` ou 503; `GET /cpf/{cpf}` (autenticado) → local, senão `buscar_por_cpf` na base (sem gravar), com 422 "CPF inválido.", 404 "CPF não encontrado na base de colaboradores.", 503 "Base de colaboradores indisponível no momento."; `GET /{id}`; `POST ""` (`gestor_or_tecnico`, cria `origem='manual'`, 409 se CPF já existe); `PUT /{id}` (`gestor_or_tecnico`). Escritas auditadas (`colaborador`).
  - Unidades: `alias_rh` opcional em create/update/response.
- [ ] **Step 1: Testes** (usam `base_rh` da Task 5): `test_sincronizar_cria_e_atualiza`; `test_sync_preserva_email_manual`; `test_ausente_da_base_vira_desligado`; `test_unidade_mapeada_por_alias`; `test_superior_preenchido_pelo_gestor_da_base`; `test_sincronizar_endpoint_exige_gestor`; `test_base_fora_do_ar_responde_503_e_registra_erro` (monkeypatch da URL para `postgresql://x:y@127.0.0.1:1/z`); `test_cpf_lookup_mensagens` (inválido 422, inexistente 404); `test_editar_campo_da_base_recusado`; `test_superior_de_si_mesmo_recusado`; `test_validar_email_e_telefone` (unit, tabela de casos incluindo `+55 (74) 99999-0000` → `74999990000`); `test_listar_busca_por_cpf_com_mascara`.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar a suíte inteira** → verde.
- [ ] **Step 5: Commit** `feat(colaboradores): cadastro sincronizado da base de RH com complemento manual`.

### Task 7: Agendador de tarefas

**Files:**
- Create: `backend/migrations/0006_scheduler_execucoes.sql` (+ down), `backend/app/services/scheduler.py`, `backend/app/jobs.py`, `backend/tests/test_scheduler.py`
- Modify: `backend/app/main.py` (startup/shutdown)

**Interfaces:**
- Consumes: `executar_sincronizacao` (Task 6), `AuditContext.sistema()` (Task 3).
- Produces:
  - Tabela `scheduler_execucoes (nome VARCHAR(100) PK, ultima_execucao DATETIME NULL, ultimo_status ENUM('ok','erro') NULL, ultimo_erro VARCHAR(500) NULL)`.
  - `registrar_job(nome: str, func: Callable[[], None], *, a_cada_segundos: int | None = None, diariamente_as: str | None = None) -> None` (exatamente um dos dois; `diariamente_as` `"HH:MM"` no fuso `America/Sao_Paulo`)
  - `jobs_registrados() -> dict[str, Job]`; `limpar_jobs() -> None` (testes)
  - `job_vencido(job: Job, ultima: datetime | None, agora: datetime) -> bool` (pura: intervalo → `ultima is None or agora - ultima >= intervalo`; diário → hoje já passou do horário e `ultima` é anterior ao horário de hoje)
  - `executar_pendentes(agora: datetime | None = None) -> list[str]` — para cada job vencido: conexão própria, `SELECT GET_LOCK('patrimonio_job_<nome>', 0)`; se pegou, relê `ultima_execucao`, roda `func()`, grava status/erro, `RELEASE_LOCK`; exceção do job é registrada e não interrompe os demais; retorna os nomes executados.
  - `iniciar() -> None` / `parar() -> None` — `asyncio` task com tique de 60 s chamando `await asyncio.to_thread(executar_pendentes)`; só inicia com `settings.SCHEDULER_ENABLED`.
  - `app.jobs.registrar_jobs() -> None` — registra `sincronizar_colaboradores` diário em `COLABORADORES_SYNC_HORA` quando `COLABORADORES_DATABASE_URL` não é vazio. **As fases seguintes acrescentam seus jobs aqui.**
- [ ] **Step 1: Testes:** `test_job_vencido_intervalo`, `test_job_vencido_diario` (antes/depois do horário, já executado hoje), `test_executar_pendentes_roda_uma_vez_e_grava`, `test_erro_no_job_registra_e_continua`, `test_lock_impede_execucao_simultanea` (segurar `GET_LOCK('patrimonio_job_x', 0)` numa conexão crua e afirmar que `executar_pendentes` não roda `x`).
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** → PASS.
- [ ] **Step 5: Commit** `feat(agendador): tarefas periódicas com trava no banco para vários workers`.

### Task 8: Locais por unidade, vínculos no item e exportações

**Files:**
- Create: `backend/migrations/0003_locais.sql` (+ down), `backend/migrations/0004_items_colaborador_local.sql` (+ down), `backend/app/db/local_db.py`, `backend/app/schemas/locais.py`, `backend/app/routers/locais.py`, `backend/app/services/exports.py`, `backend/tests/test_locais.py`, `backend/tests/test_exports.py`
- Modify: `backend/app/schemas/items.py` (`ItemResponse` ganha `colaborador_id: Optional[int] = None`, `local_id: Optional[int] = None`, `local_nome: Optional[str] = None`, `colaborador_nome: Optional[str] = None`, `colaborador_setor: Optional[str] = None`, `colaborador_cargo: Optional[str] = None`, `superior_nome: Optional[str] = None`), `backend/app/db/inventory_manager_db.py` (`list_items`/`get_item` com LEFT JOIN em `colaboradores`, superior e `locais` para preencher esses campos), `backend/app/dependencies.py`, `backend/app/main.py`

**Interfaces:**
- Produces:
  - Tabelas `locais (id INT PK AI, unidade_id INT NOT NULL FK unidades(id), nome VARCHAR(150) NOT NULL, descricao VARCHAR(255) NULL, ativo TINYINT(1) DEFAULT 1, UNIQUE(unidade_id, nome))`; `items.colaborador_id INT NULL FK colaboradores(id) ON DELETE SET NULL`, `items.local_id INT NULL FK locais(id) ON DELETE SET NULL`.
  - `LocalDBManager`: `listar(unidade_id: int | None, include_inactive: bool = False) -> list[dict]` (com `unidade_nome`), `obter(id) -> dict | None`, `criar(unidade_id: int, nome: str, descricao: str | None) -> int` (duplicado → `ValueError("Já existe um local com esse nome nesta unidade.")`), `atualizar(id: int, dados: dict) -> None`.
  - `/api/locais`: `GET` (autenticado; `unidade_id`, `include_inactive`), `POST` e `PUT /{id}` (`gestor_or_tecnico`, auditados como `local`).
  - `exports.py`: `@dataclass Coluna(chave: str, titulo: str, largura: int = 18, tipo: Literal["texto", "numero", "moeda", "data"] = "texto", maiusculas: bool = False)`; `gerar_excel(titulo: str, colunas: list[Coluna], linhas: list[dict], filtros: dict[str, str] | None = None) -> bytes` (linha 1 título, linha 2 "Gerado em DD/MM/AAAA HH:MM" + filtros, cabeçalho em negrito com fundo `#0131CD` e texto branco, congela o cabeçalho, moeda no formato `R$ #,##0.00`); `gerar_pdf(titulo: str, colunas: list[Coluna], linhas: list[dict], filtros: dict[str, str] | None = None, paisagem: bool = True) -> bytes` (fpdf2, fonte Helvetica, texto convertido para latin-1 com substituição, quebra de página repetindo o cabeçalho, rodapé "Página N"); `resposta_arquivo(conteudo: bytes, nome_base: str, formato: Literal["xlsx", "pdf"]) -> Response` (content-type correto e `Content-Disposition: attachment; filename="<nome_base>_<AAAAMMDD>.<formato>"`).
- [ ] **Step 1: Testes:** `test_criar_e_listar_locais_por_unidade`, `test_local_duplicado_na_mesma_unidade_400`, `test_local_exige_gestor_ou_tecnico`, `test_item_traz_colaborador_setor_cargo_superior_e_local` (cria colaborador manual com superior, vincula via SQL, `GET /api/items/{id}` traz os campos), `test_gerar_excel_tem_cabecalho_e_linhas` (abrir com `openpyxl.load_workbook(BytesIO(...))`), `test_gerar_pdf_comeca_com_assinatura_pdf` (`bytes[:4] == b"%PDF"`, texto com acentos não quebra), `test_maiusculas_aplicadas`.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar a suíte inteira** → verde.
- [ ] **Step 5: Commit** `feat(base): locais por unidade, vínculo de colaborador/local no item e exportação Excel/PDF`.

### Task 9: Paleta e tokens do frontend

**Files:**
- Modify: `frontend/src/index.css`, `frontend/tailwind.config.js`, `frontend/src/lib/tokens.test.ts`
- Create: `frontend/src/lib/contrast.ts`, `frontend/src/lib/palette.test.ts`

**Interfaces:**
- Produces tokens CSS (claro / escuro):
  - `--primary` `#0131CD` / `#0131CD`; `--primary-foreground` `#FFFFFF` / `#FFFFFF`; `--ring` `#0131CD` / `#B9FC01`
  - `--brand` `#0131CD` / `#0131CD`; `--brand-text` `#0131CD` / `#7D97FF`; `--brand-soft` `#E8EEFD` / `#0D1B4D`; `--brand-soft-foreground` `#0131CD` / `#C9D5FF`
  - `--lime` `#B9FC01` / `#B9FC01`; `--lime-foreground` `#111111` / `#111111`
  - `--canvas` `#F5F7FB` / `#050505`; `--surface` `#FFFFFF` / `#0C0C0C` (painéis grandes brancos/pretos)
  - Status (texto / fundo): available `#1F6F4A`/`#E9F1EC` · `#6FBF93`/`#12261C`; loaned `#0131CD`/`#E8EEFD` · `#9DB1FF`/`#0D1B4D`; pending `#7A5E17`/`#F4EEDD` · `#D4B468`/`#2A2212`; return `#53397E`/`#EDE9F4` · `#AE96D9`/`#1E1830`; maintenance `#9A4A00`/`#FDEBD9` · `#F5A15C`/`#2E1A08`; retired `#5F5F5A`/`#EEEEEA` · `#A4A4A0`/`#1A1A1A`; danger `#A32B2B`/`#F6E8E8` · `#E4A0A0`/`#2B1414`.
  - Tailwind: cores `brand` (`DEFAULT`, `text`, `soft`, `soft-foreground`), `lime` (`DEFAULT`, `foreground`), `status-*` não precisam entrar no Tailwind (uso por `var()`).
  - `contrastRatio(hexA: string, hexB: string) -> number` (fórmula WCAG 2.x de luminância relativa).
- [ ] **Step 1: Teste** `palette.test.ts`: lê `index.css` (como `tokens.test.ts` faz), extrai os blocos `:root` e `.dark` e afirma `contrastRatio >= 4.5` para: `--primary-foreground`×`--primary`, `--brand-text`×`--surface`, `--brand-soft-foreground`×`--brand-soft`, `--lime-foreground`×`--lime`, e cada `--status-X`×`--status-X-bg` (7 pares), nos dois temas; e `contrastRatio('#0131CD', '#FFFFFF')` entre 9 e 9,2.
- [ ] **Step 2: Rodar** `npx vitest run src/lib/palette.test.ts` → FAIL.
- [ ] **Step 3: Implementar** tokens, `contrast.ts`, atualizar `tokens.test.ts` (lista de tokens registrados) e o comentário de topo do `index.css` (nova identidade: base neutra com azul `#0131CD` e destaque verde-limão).
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(visual): paleta azul e verde-limão com contraste AA nos dois temas`.

### Task 10: Componentes compartilhados (StatusPill, Pagination, Button loading, formatação)

**Files:**
- Create: `frontend/src/lib/status.ts`, `frontend/src/components/ui/StatusPill.tsx`, `frontend/src/components/ui/Pagination.tsx`, `frontend/src/hooks/usePageSize.ts`, `frontend/src/lib/format.ts`, testes `*.test.tsx/ts` ao lado de cada um
- Modify: `frontend/src/components/ui/button.tsx`, `frontend/src/components/ui/badge.tsx` (`StatusBadge` passa a renderizar `StatusPill`)

**Interfaces:**
- Produces:
  - `type StatusTone = 'available' | 'loaned' | 'pending' | 'return' | 'maintenance' | 'retired' | 'danger' | 'neutral'`
  - `statusTone(status: string | null | undefined): StatusTone` — `Disponível`, `Ativo`, `Ativa`, `Assinado`, `OK` → available; `Indisponível`, `Em uso`, `Em Uso`, `Enviado` → loaned; `Pendente`, `Agendado`, `Aberta`, `Aberto` → pending; `Pendente Devolução` → return; `Em Manutenção`, `Enviada`, `Em reparo` → maintenance; `Baixado`, `Descartado`, `Substituido`, `Desativado`, `Cancelada`, `Cancelado`, `Desligado`, `Encerrado` → retired; `Erro`, `Recusado`, `Expirado`, `Sem conserto` → danger; resto → neutral (comparação sem acento e sem caixa).
  - `StatusPill({ status, tone, className }: { status?: string | null; tone?: StatusTone; className?: string })` — `rounded-full px-2.5 py-0.5 text-xs font-bold`, cor `var(--status-<tone>)` sobre `var(--status-<tone>-bg)`; `status` vazio → "—" neutro.
  - `Pagination({ page, pageSize, total, onPageChange, onPageSizeChange, pageSizeOptions = [25, 50, 100], isFetching }: { page: number; pageSize: number; total: number; onPageChange: (page: number) => void; onPageSizeChange?: (size: number) => void; pageSizeOptions?: number[]; isFetching?: boolean })` — `page` começa em 1; texto "Mostrando X–Y de Z"; botões "Anterior"/"Próxima" (`aria-label`); campo numérico "Ir para a página" que confirma com Enter ou botão "Ir" e corrige para `[1, últimaPágina]` (texto inválido não muda nada); seletor "Itens por página" só quando `onPageSizeChange` existe.
  - `usePageSize(tela: string, padrao: 25 | 50 | 100 = 25): [number, (n: number) => void]` — `localStorage` chave `revalle_page_size:<tela>`, só aceita 25/50/100, `try/catch` em leitura/escrita.
  - `Button` ganha `loading?: boolean`: mostra `Loader2` com `animate-spin motion-reduce:animate-none`, `disabled`, `aria-busy="true"`; não dispara `onClick` enquanto `loading`.
  - `format.ts`: `upper(v: string | null | undefined): string` (`toLocaleUpperCase('pt-BR')`, vazio → `''`); `formatTelefone(v)` (`(74) 99999-0000` / `(74) 3611-0000`); `formatMoeda(v: number | string | null | undefined)` (`R$ 1.234,56`); `formatDataBR(iso)` (reexporta `formatDate` de `utils.ts`).
- [ ] **Step 1: Testes:** `status.test.ts` (tabela de casos); `StatusPill.test.tsx` (texto em negrito e classe/estilo do tom); `Pagination.test.tsx` (`'Mostrando 26–50 de 120'` em `page=2,pageSize=25,total=120`; "Próxima" desabilitado na última; "Ir para" com `999` chama `onPageChange(5)`, com `0` chama `onPageChange(1)`, com `abc` não chama; troca de tamanho chama `onPageSizeChange(50)`); `usePageSize.test.ts` (valor inválido no storage → padrão; persiste); `button.test.tsx` (`loading` → `aria-busy`, desabilitado, clique não chama `onClick`); `format.test.ts`; `badge.test.tsx` existente continua passando (ajustar só o que for visual de propósito).
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(ui): StatusPill, Pagination, botão com carregamento e helpers de formatação`.

### Task 11: Página de detalhe do equipamento com abas plugáveis

**Files:**
- Create: `frontend/src/features/item-detail/types.ts`, `frontend/src/features/item-detail/registry.ts`, `frontend/src/features/item-detail/tabs/dados.tsx`, `frontend/src/features/item-detail/tabs/historico.tsx`, `frontend/src/pages/ItemDetailPage.tsx`, `frontend/src/pages/ItemDetailPage.test.tsx`
- Modify: `frontend/src/App.tsx` (rota `items/:id`, todos os papéis), `frontend/src/api/items.ts` (campos novos de `Item`), `frontend/src/api/history.ts` (`item_id` em `ListHistoryParams`), `frontend/src/components/equipment/ItemDetailsModal.tsx` (link "Ver página completa")

**Interfaces:**
- Consumes: `GET /api/items/{id}` com os campos novos e `GET /api/history?item_id=` (Tasks 4 e 8); `StatusPill`, `upper` (Task 10).
- Produces:
  - `Item` ganha `colaborador_id?: number | null; local_id?: number | null; local_nome?: string | null; colaborador_nome?: string | null; colaborador_setor?: string | null; colaborador_cargo?: string | null; superior_nome?: string | null`.
  - `interface ItemDetailTab { id: string; label: string; order: number; roles?: string[]; Component: ComponentType<{ item: Item }> }`
  - `registry.ts`: `getItemDetailTabs(role: string | undefined): ItemDetailTab[]` — monta a lista com `import.meta.glob<{ default: ItemDetailTab }>('./tabs/*.tsx', { eager: true })`, filtra por papel e ordena por `order`. **Cada fase acrescenta uma aba criando um arquivo em `features/item-detail/tabs/` (default export `ItemDetailTab`), sem editar outros arquivos.** Ordens reservadas: dados 10, linha do tempo 20 (F4), termos 30 (F3), manutenções 40 (F4), linha telefônica 50 (F5), acessos 60 (F8), histórico 90.
  - `ItemDetailPage` em `/items/:id`: cabeçalho com tipo, marca/modelo, identificador, `StatusPill`, unidade em caixa alta; abas acessíveis (`role="tablist"`, aba ativa em `?aba=<id>`); 404 → estado vazio "Equipamento não encontrado".
  - Aba `dados`: todos os campos preenchidos do item, mais colaborador (nome, setor, cargo), **superior imediato** e local.
  - Aba `historico`: lista paginada (`Pagination`) de `GET /api/history?item_id=`.
- [ ] **Step 1: Testes** (`ItemDetailPage.test.tsx`, MSW): renderiza abas `Dados` e `Histórico`; mostra "Superior imediato" com o nome; `?aba=historico` abre a aba de histórico e chama a API com `item_id`; item inexistente mostra "Equipamento não encontrado"; `registry` filtra aba com `roles: ['Gestor']` para Técnico.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(ui): página de detalhe do equipamento com abas plugáveis`.

### Task 12: Tela de colaboradores e seletor de colaborador

**Files:**
- Create: `frontend/src/api/colaboradores.ts`, `frontend/src/pages/ColaboradoresPage.tsx`, `frontend/src/pages/ColaboradoresPage.test.tsx`, `frontend/src/components/colaboradores/ColaboradorSelect.tsx`, `frontend/src/components/colaboradores/ColaboradorSelect.test.tsx`, `frontend/src/components/colaboradores/ColaboradorFormDialog.tsx`
- Modify: `frontend/src/App.tsx` (rota `colaboradores`, Gestor e Técnico), `frontend/src/components/layout/navigation.ts` (grupo "Gestão": "Colaboradores", ícone `Contact` do lucide)

**Interfaces:**
- Consumes: API da Task 6; `Pagination`, `usePageSize`, `StatusPill`, `Button loading`, `upper`, `formatTelefone` (Task 10).
- Produces:
  - `api/colaboradores.ts`: `interface Colaborador { id: number; cpf: string; nome: string; matricula?: string | null; cargo?: string | null; setor?: string | null; unidade_rh?: string | null; unidade_id?: number | null; unidade_nome?: string | null; situacao: 'ativo' | 'afastado' | 'desligado'; email?: string | null; telefone?: string | null; email_manual: boolean; telefone_manual: boolean; superior_id?: number | null; superior_nome?: string | null; origem: 'rh' | 'manual'; sincronizado_em?: string | null }`; `listColaboradores(params: { busca?: string; situacao?: string; unidade_id?: number; limit?: number; offset?: number }): Promise<Paginated<Colaborador>>`; `getColaborador(id)`; `getColaboradorPorCpf(cpf)`; `createColaborador(data)`; `updateColaborador(id, data)`; `sincronizarColaboradores(): Promise<{ total_base: number; criados: number; atualizados: number; desligados: number }>`; `getSyncStatus()`; `listCargos(): Promise<string[]>`.
  - `ColaboradorSelect({ value, onChange, somenteAtivos = true, placeholder, disabled }: { value: number | null; onChange: (c: Colaborador | null) => void; somenteAtivos?: boolean; placeholder?: string; disabled?: boolean })` — busca no servidor (debounce 300 ms, `listColaboradores({ busca, situacao: somenteAtivos ? 'ativo' : undefined, limit: 20 })`) usando `SearchableSelect` com `search`/`onSearchChange`; rótulo `NOME`, subtítulo `CPF formatado · cargo · UNIDADE`.
  - `ColaboradoresPage`: busca, filtros de situação e unidade, tabela (nome, CPF, cargo, setor, UNIDADE, situação em `StatusPill`, e-mail, telefone, superior), `Pagination` com `usePageSize('colaboradores')`; Gestor vê "Sincronizar agora" (com `loading`) e o texto da última sincronização (data, status, criados/atualizados/desligados ou o erro); "Novo colaborador" (cadastro manual) e "Editar" em `ColaboradorFormDialog` (e-mail, telefone, superior via `ColaboradorSelect`; campos da base só leitura para `origem='rh'`, com aviso "Vem da base de RH").
- [ ] **Step 1: Testes** (MSW): lista e pagina; "Sincronizar agora" só para Gestor e mostra o resumo em toast; editar e-mail inválido mostra a mensagem da API; `ColaboradorSelect` chama a API com `busca` depois de digitar e devolve o objeto selecionado.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(colaboradores): tela de colaboradores e seletor reutilizável`.

### Task 13: Unidades (nome na base de RH e locais) e tela de auditoria

**Files:**
- Create: `frontend/src/api/locais.ts`, `frontend/src/api/audit.ts`, `frontend/src/components/locais/LocaisDialog.tsx`, `frontend/src/components/locais/LocalSelect.tsx`, `frontend/src/pages/AuditPage.tsx`, testes `LocaisDialog.test.tsx`, `AuditPage.test.tsx`
- Modify: `frontend/src/api/unidades.ts` (`alias_rh`), `frontend/src/pages/UnidadesPage.tsx` (campo "Nome na base de RH" e ação "Locais"), `frontend/src/App.tsx` (rota `auditoria`, só Gestor), `frontend/src/components/layout/navigation.ts` ("Auditoria", grupo "Gestão", Gestor, ícone `ScrollText`)

**Interfaces:**
- Consumes: `/api/locais`, `/api/audit`, `alias_rh` (Tasks 3, 6, 8).
- Produces:
  - `api/locais.ts`: `interface Local { id: number; unidade_id: number; unidade_nome?: string; nome: string; descricao?: string | null; ativo: boolean }`; `listLocais(params: { unidade_id?: number; include_inactive?: boolean }): Promise<Local[]>`; `createLocal(data: { unidade_id: number; nome: string; descricao?: string })`; `updateLocal(id, data)`.
  - `LocalSelect({ unidadeId, value, onChange, disabled }: { unidadeId: number | null; value: number | null; onChange: (id: number | null) => void; disabled?: boolean })` — desabilitado sem unidade; opção "Sem local".
  - `LocaisDialog({ unidade, open, onOpenChange })` — lista, cria, renomeia e ativa/desativa locais da unidade.
  - `api/audit.ts`: `interface AuditEntry { id: number; ocorrido_em: string; username?: string | null; papel?: string | null; ip?: string | null; entidade: string; entidade_id?: string | null; acao: string; antes?: Record<string, unknown> | null; depois?: Record<string, unknown> | null }`; `listAudit(params)`.
  - `AuditPage`: filtros (entidade, usuário, período), tabela com data/hora local, usuário, entidade #id, ação e um resumo "campo: antes → depois", `Pagination`.
- [ ] **Step 1: Testes** (MSW): `LocaisDialog` cria local e mostra o erro de duplicado; `UnidadesPage` envia `alias_rh`; `AuditPage` mostra "status: Disponível → Indisponível" a partir de `antes`/`depois`.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run`, `npx tsc --noEmit` e `npm run build` → verdes.
- [ ] **Step 5: Commit** `feat(base): locais e nome na base de RH nas unidades, e tela de auditoria`.
