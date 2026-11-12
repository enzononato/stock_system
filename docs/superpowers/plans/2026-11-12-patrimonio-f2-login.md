# Fase 2 — Login (esqueci minha senha e 2FA) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Recuperação de senha por e-mail e verificação em duas etapas (TOTP) no login — obrigatória para Gestor, opcional para os demais —, com e-mail dos usuários do sistema, página de perfil e zeragem auditada do 2FA.

**Architecture:** No backend, `routers/auth.py` (dono: F2) passa a fazer o login em duas etapas e o cadastro do autenticador durante o login; routers novos `auth_senha.py` (forgot/reset) e `perfil.py`; managers novos `password_reset_db.py` e `mfa_db.py`; serviços novos `mfa.py` (cifra, TOTP e códigos — funções puras), `segundo_fator.py` (orquestra banco + auditoria) e `sessao.py` (versão de sessão por usuário: todo token leva o claim `sv`; `get_current_user` e `/refresh` comparam com `usuarios.sessao_versao`; redefinir a senha ou zerar o 2FA incrementa a versão e derruba todas as sessões). No frontend, `AuthContext`/`api/auth.ts` com as etapas, componentes de QR e de códigos em `components/auth/`, páginas públicas `/esqueci-senha` e `/redefinir-senha`, `/perfil` para todos os papéis e a `UsersPage` refeita como tabela paginada.

**Tech Stack:** FastAPI 0.115 + PyMySQL (SQL cru) + slowapi + python-jose; `pyotp==2.9.0`; `cryptography` (Fernet). React 18 + TS 5.6 + TanStack Query 5 + Vitest/MSW; `qrcode@1.5.4` + `@types/qrcode@1.5.5` (a mesma lib que a F6 usa nas etiquetas).

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md` — seção 6 "F2 — Login", matriz (seção 3) linhas 1.1 e 1.2, seção 5 "F2", seção 8. Contexto comum das fases: `C:/Users/enzo.jz/dev/patrimonio/_tools/contexto-planejamento.md`. Interfaces da F0: `docs/superpowers/plans/2026-11-12-patrimonio-f0-base.md`.

**Ambiente:** worktree `C:/Users/enzo.jz/dev/patrimonio/f2` (branch `fase/f2-login`, criada de `feat/patrimonio-melhorias` com a F0 integrada). Backend: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh C:/Users/enzo.jz/dev/patrimonio/f2 f2 <args do pytest>` — abaixo abreviado `RBT <args>`. Frontend: em `frontend/`, `npx vitest run <arquivo>`, `npx tsc --noEmit`, `npm run build`.

## Global Constraints

- Esqueci minha senha: `POST /auth/forgot` com resposta **sempre genérica** e rate limit; e-mail com link `APP_PUBLIC_URL/redefinir-senha?token=...` **válido por 30 min e de uso único**, token guardado **como hash**; `POST /auth/reset` aplica as **mesmas regras de força atuais** (hoje: senha não pode estar em branco — `"A nova senha não pode estar em branco."`), **revoga as sessões abertas** do usuário e registra na auditoria.
- 2FA (TOTP): cadastro com QR code e confirmação de um código; **10 códigos de recuperação exibidos uma vez**; segredo **cifrado com `TOTP_ENCRYPTION_KEY`** (Fernet). Login em duas etapas: senha → token temporário de **5 min** com `mfa_pendente` → `POST /auth/mfa/verificar` (código TOTP ou de recuperação, com rate limit) → tokens normais.
- Gestor sem 2FA é levado ao cadastro no primeiro login e **não usa o sistema antes de concluir** — garantido no servidor (sem access/refresh token normal antes do cadastro). Demais papéis ativam no perfil, se quiserem.
- Gestor pode zerar o 2FA de outro usuário (com auditoria). `python -m scripts.reset_2fa <usuario>`, rodado no servidor, zera o 2FA (e registra na auditoria).
- Modelo (seção 5, F2): `usuarios.email VARCHAR NULL UNIQUE`, `usuarios.totp_secret_cifrado VARCHAR NULL`, `usuarios.totp_ativo TINYINT DEFAULT 0`; `password_reset_tokens (id, usuario_id FK, token_hash, expira_em, usado_em, criado_ip)`; `mfa_recovery_codes (id, usuario_id FK, code_hash, usado_em)`.
- Migrações da F2: `backend/migrations/02NN_*.sql` + `.down.sql` (faixa `0200–0299`); comandos terminam com `;` no fim da linha, sem `; -- comentário` e sem `;` dentro de literais. A primeira que altera `usuarios` começa com `CREATE TABLE IF NOT EXISTS usuarios (id INT AUTO_INCREMENT PRIMARY KEY, username VARCHAR(100) UNIQUE NOT NULL, password VARCHAR(255) NOT NULL, role ENUM('Gestor','Técnico','Jovem Aprendiz') NOT NULL)`.
- Segurança (seção 8): tokens de redefinição e códigos de recuperação só como hash; segredo TOTP cifrado; segredos só no `.env`; nada de token, senha, segredo ou código em log ou na auditoria.
- Toda escrita nova grava `audit.record(ctx, "usuario", <id>, <acao>, antes, depois)`; ações desta fase: `editar` (e-mail), `solicitar_redefinicao_senha`, `redefinir_senha`, `ativar_2fa`, `desativar_2fa`, `regenerar_codigos_2fa`, `usar_codigo_recuperacao`, `zerar_2fa`.
- Mensagens 403 padrão: `Acesso restrito ao Gestor.`; textos de interface e mensagens em português.
- Arquivos que só a F2 edita: `routers/auth.py`, `routers/users.py`, `db/user_manager_db.py`, `core/security.py`, `LoginPage`, `contexts/AuthContext.tsx`, `api/auth.ts`, `UsersPage`, `api/users.ts`. Compartilhados (`app/main.py`, `app/dependencies.py`, `.env.example`, `tests/conftest.py`, `App.tsx`, `navigation.ts`): só **acrescentar** linhas, sem reordenar nem reformatar. Não alterar `StatusPill`, `Pagination`, `Button`, `api/client.ts`.
- Dependências novas: só `qrcode@1.5.4` e `@types/qrcode@1.5.5` (frontend). `pyotp` e `cryptography` já estão no `requirements.txt`. Nenhuma variável de ambiente nova (`TOTP_ENCRYPTION_KEY`, `APP_PUBLIC_URL`, `LOGIN_RATE_LIMIT` já existem).
- Commits em português `tipo(escopo): descrição`, terminando com linha em branco + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Decisões desta fase (já tomadas — o executor não decide de novo)

1. **Revogação de sessões = contador por usuário**, não `iat`: coluna `usuarios.sessao_versao INT NOT NULL DEFAULT 0`; todo access/refresh/mfa token leva `sv`; token sem `sv` (emitido antes do deploy) vale como `0`. Contador não tem a corrida de "token emitido no mesmo segundo da redefinição" que `iat` (resolução de 1 s) teria, e não depende de relógio. `get_current_user` faz 1 SELECT por requisição (`role, sessao_versao, totp_ativo` pela PK). `CurrentUser` continua vindo dos claims do token (o `test_nao_remove_o_ultimo_gestor` depende disso).
2. **Chave TOTP ausente ou inválida** (`TOTP_ENCRYPTION_KEY` vazia ou que o Fernet recusa): (a) ninguém é obrigado a cadastrar 2FA — o Gestor sem 2FA entra só com a senha; (b) os endpoints que precisam da chave (iniciar/confirmar cadastro, verificar código TOTP) respondem **503** `"A verificação em duas etapas não está disponível: o servidor está sem a chave TOTP_ENCRYPTION_KEY. Avise o administrador."`; (c) usuário que **já tem** 2FA ativo continua passando pela 2ª etapa (sem bypass): o código TOTP responde 503, mas **código de recuperação funciona** (só hash, não precisa da chave), e `scripts.reset_2fa` é a saída final; (d) chave trocada (segredo não decifra) → 503 `"Não foi possível validar o código do autenticador com a chave atual do servidor. Use um código de recuperação ou peça a um Gestor para zerar o seu 2FA."`; (e) o boot registra `logger.warning` e o perfil mostra o aviso. Consequência para os testes: o runner não define a chave, então os fixtures `client_gestor` de todas as fases continuam funcionando; os testes da F2 ligam a chave com o fixture `chave_totp`.
3. **Token de cadastro do Gestor no login** usa o mesmo tipo `mfa_pendente` com `proposito="cadastrar"` e dura **15 min** (instalar o app autenticador leva mais que 5 min); o de verificação (`proposito="verificar"`) dura os **5 min** do spec. Os dois são de uso único (o `jti` é revogado no sucesso) e não servem como access token (`type != "access"`).
4. **Erros de negócio das etapas de login, senha e perfil nunca são 401** (código errado/etapa expirada/link inválido/senha incorreta → 400; já ativo → 409; sem chave → 503; limite → 429). O interceptor de `api/client.ts` trata 401 como sessão expirada; assim ele não precisa mudar. 401 fica para o access token (sessão encerrada, cadastro de 2FA obrigatório).
5. **Rate limit**: o `limiter` existente com `settings.LOGIN_RATE_LIMIT` (por IP, por rota) em `forgot`, `reset`, `mfa/verificar`, `mfa/cadastro/confirmar`, `perfil/2fa/desativar` e `perfil/2fa/codigos` — sem variável nova, e os testes já sobem esse limite para `10000/minute`. Além disso, no máximo **3 e-mails de redefinição por usuário por hora** (contra bombardeio de e-mail vindo de vários IPs); acima disso a resposta continua a genérica.
6. "Esqueci minha senha" aceita **usuário ou e-mail** (`identificador`; com `@` busca por e-mail, sem diferenciar caixa). O e-mail sai numa `BackgroundTask` (o tempo de resposta não depende do SMTP, o que evitaria enumeração por tempo); falha de envio só vira `logger.error` sem dados.
7. **Só o Gestor edita o e-mail** de um usuário (tela de Usuários); o perfil mostra o próprio e-mail **só leitura**. O e-mail é o canal de recuperação: se uma sessão sequestrada pudesse trocá-lo, tomaria a conta.
8. Códigos de recuperação: 16 caracteres do alfabeto `abcdefghijklmnopqrstuvwxyz234567`, exibidos como `xxxx-xxxx-xxxx-xxxx` (80 bits); token de redefinição: `secrets.token_urlsafe(32)`. Ambos guardados como SHA-256 hex (entropia alta dispensa bcrypt). Código de recuperação é aceito com ou sem hífen/espaços, em qualquer caixa.
9. Anti-replay do TOTP: coluna `usuarios.totp_ultimo_passo BIGINT NULL`; aceita a janela ±1 passo de 30 s, mas só passo maior que o último aceito.
10. Redefinir a senha **não** mexe no 2FA. Zerar o 2FA (Gestor ou script) também encerra as sessões do alvo.
11. `GET /api/users` continua devolvendo lista (sem quebrar contrato); a `UsersPage` pagina no cliente (lista pequena, já vem inteira).
12. Colunas além do spec: `usuarios.totp_ultimo_passo`, `usuarios.sessao_versao`, `password_reset_tokens.criado_em`. Datas novas da F2 são gravadas em UTC (`datetime.now(timezone.utc).replace(tzinfo=None)`), sempre passadas pelo Python (nunca `NOW()` do MySQL).
13. Premissa de produção: `usuarios` é InnoDB (as FKs `ON DELETE CASCADE` das tabelas novas dependem disso); a validação "migrações sobre uma cópia do schema atual" da integração confirma.

## Review Focus

- Usuários antigos de produção (sem e-mail, `totp_ativo` padrão) → listar, editar, entrar e "esqueci a senha" funcionam; sem e-mail, o forgot dá a resposta genérica e não envia nada (testes `test_usuarios_legados_sem_email_e_sem_2fa`, Task 1, e `test_forgot_usuario_sem_email_nao_envia`, Task 3).
- Servidor sem `TOTP_ENCRYPTION_KEY` (ou com chave inválida/trocada) → Gestor entra só com a senha, endpoints de 2FA dão 503 com a mensagem da Decisão 2, quem já tem 2FA entra com código de recuperação (testes da Task 6).
- Código digitado com espaço (`123 456`), código de recuperação em maiúsculas/sem hífen e o mesmo código TOTP usado duas vezes → os dois primeiros aceitos, o repetido recusado (testes da Task 4 e `test_verificar_codigo_reutilizado_recusado`, Task 5).
- Outra aba/dispositivo com sessão aberta quando a senha é redefinida ou o 2FA é zerado → access e refresh antigos recusados com `"Sessão encerrada. Faça login novamente."` (testes `test_reset_revoga_sessoes`, Task 3, e `test_zerar_2fa_encerra_sessoes_do_alvo`, Task 8).
- SMTP fora do ar ou não configurado no "esqueci minha senha" → mesma resposta genérica, sem 500 e sem token/e-mail no log (teste `test_forgot_falha_de_envio_responde_igual`, Task 3).

---

### Task 1: Colunas novas de `usuarios` e e-mail no cadastro de usuários

**Files:**
- Create: `backend/migrations/0200_usuarios_email_2fa.sql`, `backend/migrations/0200_usuarios_email_2fa.down.sql`, `backend/tests/fixtures_auth.py`, `backend/tests/test_usuarios_email.py`
- Modify: `backend/app/db/user_manager_db.py`, `backend/app/schemas/users.py`, `backend/app/routers/users.py`, `backend/app/main.py` (rótulo `"email": "E-mail"` só se a F0 ainda não o tiver posto)

**Interfaces:**
- Consumes: `validar_email(email: str | None) -> str | None` (`app/schemas/validators.py`, F0 Task 6; levanta `ValueError("E-mail inválido.")`); `audit.record`, `get_audit_context` (F0 Task 3); auditoria que a F0 Task 4 já pôs em `users.py` (`criar`, `excluir`, `editar` com `{"senha": "alterada"}`) — mantida.
- Produces:
  - Colunas: `email VARCHAR(200) NULL` + `UNIQUE INDEX uq_usuarios_email (email)`, `totp_secret_cifrado VARCHAR(255) NULL`, `totp_ativo TINYINT(1) NOT NULL DEFAULT 0`, `totp_ultimo_passo BIGINT NULL`, `sessao_versao INT NOT NULL DEFAULT 0`.
  - `UserDBManager`: `get_all_users() -> list[dict]` (`id, username, role, email, totp_ativo`); `get_user_by_username(username) -> dict | None` (`id, username, password, role, email, totp_ativo, sessao_versao`); `get_user_by_id(user_id) -> dict | None` (`id, username, role, email, totp_ativo, sessao_versao`); `get_user_by_email(email: str) -> dict | None` (mesmos campos de `get_user_by_id`, `LOWER(email) = LOWER(%s)`); `add_user(username, password, role, email: str | None = None) -> tuple[bool, str]`; `update_email(user_id: int, email: str | None) -> tuple[bool, str]` (`"E-mail atualizado com sucesso."`). Erro 1062 cuja mensagem cita `uq_usuarios_email` → `(False, "Este e-mail já está em uso por outro usuário.")` nos dois métodos (o 1062 de username continua com a mensagem atual). Não usar `rowcount` para existência em `update_email` (o MySQL devolve 0 quando o valor não muda).
  - Schemas: `UserCreate.email: Optional[str] = None`; `UserResponse.email: Optional[str] = None`, `UserResponse.totp_ativo: bool = False`; `UserUpdate(email: Optional[str])` (chave obrigatória, aceita `null`). Validador comum: texto vazio/só espaços → `None`; senão `validar_email(v.strip())`.
  - `PUT /api/users/{user_id}` (`gestor_only`) body `UserUpdate` → `UserResponse`; 404 `"Usuário não encontrado."`; 400 e-mail duplicado; auditoria `editar` com `audit.diff({"email": antes}, {"email": depois})`. `POST /api/users` grava `email` e inclui `"email"` no `depois` da auditoria `criar`.
  - `tests/fixtures_auth.py`: `linhas_auditoria(entidade: str, acao: str | None = None) -> list[dict]` (lê `audit_log` com `get_cursor`, `antes`/`depois` já com `json.loads`, ordem por `id`); `ida_e_volta_migracao(versao: int) -> None` (acha `backend/migrations/{versao:04d}_*.sql`, roda o `.down.sql` e depois o `.sql` com `app.db.migrations.executar_sql`; o `.sql` roda num `finally`). Testes da F2 importam o que usam: `from tests.fixtures_auth import ...  # noqa: F401` (fixtures da F2 não entram no `conftest.py`, que só recebe a linha da Task 6).

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_usuarios_email.py` (`pytestmark = pytest.mark.integration`):
  - `test_migracao_0200_cria_colunas_e_indice` — `information_schema.columns` de `usuarios` tem `email` (`varchar`, 200, `IS_NULLABLE='YES'`), `totp_secret_cifrado`, `totp_ativo` (`COLUMN_DEFAULT='0'`), `totp_ultimo_passo`, `sessao_versao` (`COLUMN_DEFAULT='0'`); `information_schema.statistics` tem `uq_usuarios_email` com `NON_UNIQUE=0`.
  - `test_migracao_0200_idempotente_e_reversivel` — executa o texto do `.sql` duas vezes com `executar_sql` sem erro; `ida_e_volta_migracao(200)`; as 5 colunas existem no fim.
  - `test_usuarios_legados_sem_email_e_sem_2fa` — `executar_sql("INSERT INTO usuarios (username, password, role) VALUES ('legado', %s, 'Técnico')", (hash_password("Senha#1"),))`; `GET /api/users` traz `legado` com `email is None` e `totp_ativo is False`; `POST /api/auth/login` de `legado` → 200 com `access_token`.
  - `test_criar_usuario_com_email_normaliza` — `POST /api/users` com `"email": "Joao.Silva@Revalle.com.br"` → listagem mostra `"joao.silva@revalle.com.br"`; `linhas_auditoria("usuario", "criar")[-1]["depois"]["email"] == "joao.silva@revalle.com.br"` e `"$2b$"` não aparece em `json.dumps` da linha.
  - `test_criar_usuario_email_vazio_vira_null` (`"email": "   "` → `email is None`).
  - `test_criar_usuario_email_invalido_422` (`"joao@"` → 422 e `"E-mail inválido."` em `detail`).
  - `test_criar_usuario_email_duplicado_400` → 400 `"Este e-mail já está em uso por outro usuário."`.
  - `test_editar_email` — `PUT /api/users/{tecnico}` `{"email": "tec@revalle.com.br"}` → 200 com `email == "tec@revalle.com.br"`; auditoria `editar` com `antes == {"email": None}` e `depois == {"email": "tec@revalle.com.br"}`.
  - `test_editar_email_null_limpa`; `test_editar_email_duplicado_400`; `test_editar_sem_chave_email_422` (body `{}`); `test_editar_usuario_inexistente_404` (`"Usuário não encontrado."`); `test_editar_email_exige_gestor` (Técnico → 403 `"Acesso restrito ao Gestor."`, anônimo → 401).
  - `test_listar_traz_email_e_totp_ativo` (chaves `email` e `totp_ativo` em cada item; `password` ausente).
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_usuarios_email.py -q` → FAIL (colunas e endpoint inexistentes).
- [ ] **Step 3: Implementar.** A migração começa exatamente com o `CREATE TABLE IF NOT EXISTS usuarios (...)` dos Global Constraints e, para cada coluna (e para o índice, consultando `information_schema.statistics`), usa este bloco idempotente (MySQL 8 não tem `ADD COLUMN IF NOT EXISTS`):
  ```sql
  SET @faltando := (SELECT COUNT(*) = 0 FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'usuarios' AND column_name = 'email');
  SET @sql := IF(@faltando, 'ALTER TABLE usuarios ADD COLUMN email VARCHAR(200) NULL', 'DO 0');
  PREPARE passo FROM @sql;
  EXECUTE passo;
  DEALLOCATE PREPARE passo;
  ```
  O `.down.sql` usa o mesmo bloco invertido (`DROP INDEX uq_usuarios_email`, depois `DROP COLUMN` de cada coluna, só se existir) e nunca apaga a tabela `usuarios`.
- [ ] **Step 4: Rodar:** `RBT tests/test_usuarios_email.py tests/test_users.py tests/test_audit_fluxos.py tests/test_migrations.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(usuarios): e-mail dos usuários do sistema e colunas do 2FA`.

### Task 2: Versão da sessão (encerrar todas as sessões de um usuário)

**Files:**
- Create: `backend/app/services/sessao.py`, `backend/tests/test_sessoes.py`
- Modify: `backend/app/db/user_manager_db.py`, `backend/app/dependencies.py` (`get_current_user`: +1 import, +1 linha), `backend/app/routers/auth.py` (`login`, `refresh`)

**Interfaces:**
- Consumes: colunas e `get_user_by_username` da Task 1; `decode_token`, `create_access_token`, `create_refresh_token` (inalteradas).
- Produces:
  - `UserDBManager.get_estado_sessao(user_id: int) -> dict | None` (`role, sessao_versao, totp_ativo`); `UserDBManager.revogar_sessoes(user_id: int) -> None` (`SET sessao_versao = sessao_versao + 1`).
  - `sessao.MSG_SESSAO_ENCERRADA = "Sessão encerrada. Faça login novamente."`
  - `sessao.dados_do_token(usuario: dict) -> dict` → `{"sub": str(id), "username": ..., "role": ..., "sv": int(sessao_versao or 0)}`.
  - `sessao.validar_sessao_ativa(payload: dict) -> None` — usuário inexistente ou `payload.get("sv", 0) != sessao_versao` → `HTTPException(401, MSG_SESSAO_ENCERRADA, headers={"WWW-Authenticate": "Bearer"})`. Instancia `UserDBManager()` direto (não importa `app.dependencies`, que importa este módulo).
  - `get_current_user`: chama `validar_sessao_ativa(payload)` logo depois de `decode_token`; o resto não muda.
  - `/api/auth/login` emite os tokens com `dados_do_token(user)`; `/api/auth/refresh` chama `validar_sessao_ativa(payload)` depois da checagem de `jti` revogado e emite com `{"sub", "username", "role"}` do payload + `"sv": payload.get("sv", 0)`.

- [ ] **Step 1: Testes** (`integration`) em `backend/tests/test_sessoes.py`:
  - `test_tokens_levam_sv` — login → `jose_jwt.get_unverified_claims(access)["sv"] == 0` e o mesmo no cookie `refresh_token`.
  - `test_revogar_sessoes_derruba_access_e_refresh` — login (guarda access e cookie); `UserDBManager().revogar_sessoes(id)`; `GET /api/auth/me` com o access antigo → 401 `"Sessão encerrada. Faça login novamente."`; `POST /api/auth/refresh` com o cookie antigo → 401 com a mesma mensagem; novo login → `sv == 1` e `/me` 200.
  - `test_token_legado_sem_sv_continua_valido` — `create_access_token({"sub": str(id), "username": "tecnico.teste", "role": "Técnico"})` → `/me` 200.
  - `test_token_de_usuario_removido_e_recusado` — login do Técnico; Gestor remove o Técnico; `/me` com o token do Técnico → 401 `MSG_SESSAO_ENCERRADA`.
  - `test_refresh_mantem_sv` — após `revogar_sessoes` + novo login (`sv` 1), `/refresh` devolve access com `sv == 1` que autentica.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_sessoes.py -q`.
- [ ] **Step 3: Implementar** conforme as Interfaces.
- [ ] **Step 4: Rodar:** `RBT tests/test_sessoes.py tests/test_auth.py tests/test_rbac.py tests/test_users.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(auth): versão da sessão por usuário para encerrar todas as sessões`.

### Task 3: Esqueci minha senha

**Files:**
- Create: `backend/migrations/0201_password_reset_tokens.sql` (+ `.down.sql`: `DROP TABLE IF EXISTS password_reset_tokens;`), `backend/app/db/password_reset_db.py`, `backend/app/routers/auth_senha.py`, `backend/tests/test_redefinir_senha.py`
- Modify: `backend/app/core/security.py`, `backend/app/schemas/auth.py`, `backend/app/dependencies.py` (acrescentar `get_password_reset_db`, `contexto_auditoria_de`), `backend/app/main.py` (acrescentar `from app.routers import auth_senha`, `app.include_router(auth_senha.router)` e os rótulos `"identificador": "Usuário ou e-mail"`, `"nova_senha": "Nova senha"`, `"token": "Link"`), `backend/tests/fixtures_auth.py`

**Interfaces:**
- Consumes: `send_mail`, `render_email`, `MailError`, `outbox`, `limpar_outbox` (F0 Task 2); `get_user_by_username`, `get_user_by_email`, `get_user_by_id`, `update_password`, `revogar_sessoes` (Tasks 1–2); `limiter`, `ip_real_do_cliente`.
- Produces:
  - Tabela:
    ```sql
    CREATE TABLE IF NOT EXISTS password_reset_tokens (
        id INT AUTO_INCREMENT PRIMARY KEY,
        usuario_id INT NOT NULL,
        token_hash CHAR(64) NOT NULL,
        criado_em DATETIME NOT NULL,
        expira_em DATETIME NOT NULL,
        usado_em DATETIME NULL,
        criado_ip VARCHAR(45) NULL,
        UNIQUE KEY uq_password_reset_token_hash (token_hash),
        INDEX idx_password_reset_usuario (usuario_id, criado_em),
        CONSTRAINT fk_password_reset_usuario FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
    ) CHARACTER SET utf8mb4;
    ```
  - `security.sha256_hex(texto: str) -> str`.
  - `PasswordResetDBManager`: `criar(usuario_id: int, token_hash: str, criado_em: datetime, expira_em: datetime, criado_ip: str | None) -> None`; `contar_criados_desde(usuario_id: int, desde: datetime) -> int`; `consumir(token_hash: str, agora: datetime) -> int | None` (um `UPDATE ... SET usado_em=%s WHERE token_hash=%s AND usado_em IS NULL AND expira_em > %s`; `rowcount == 1` → devolve `usuario_id`, senão `None` — atômico contra uso duplo); `invalidar_abertos(usuario_id: int, agora: datetime) -> None`.
  - `dependencies.get_password_reset_db()` (`@lru_cache`); `dependencies.contexto_auditoria_de(usuario: dict | None, request: Request) -> AuditContext` (`None` → `AuditContext(None, None, None, ip)`).
  - Schemas: `ForgotRequest(identificador: str, max_length=200)`, `ResetRequest(token: str, nova_senha: str)`, `MensagemResponse(detail: str)`.
  - `auth_senha.py` (`APIRouter(prefix="/api/auth", tags=["auth"])`): `RESET_VALIDADE_MINUTOS = 30`, `RESET_MAX_POR_HORA = 3`, `MSG_FORGOT = "Se o usuário existir e tiver e-mail cadastrado, enviaremos um link para redefinir a senha."`, `MSG_LINK_INVALIDO = "Link inválido ou expirado. Peça um novo."`, `MSG_SENHA_REDEFINIDA = "Senha redefinida. Entre com a nova senha."`.
    - `POST /api/auth/forgot` (`@limiter.limit(settings.LOGIN_RATE_LIMIT)`, recebe `request: Request` e `BackgroundTasks`) → sempre 200 `{"detail": MSG_FORGOT}`. Só quando o usuário existe, tem e-mail e `contar_criados_desde(id, agora - 1 h) < RESET_MAX_POR_HORA`: gera `secrets.token_urlsafe(32)`, grava `sha256_hex(token)` com `expira_em = agora + 30 min` e o IP, audita `solicitar_redefinicao_senha` (contexto anônimo com IP, `depois={"canal": "email"}`) e agenda `enviar_email_redefinicao`.
    - `enviar_email_redefinicao(email: str, username: str, token: str) -> None` — assunto `"Redefinição de senha — Patrimônio Revalle"`; link `f"{settings.APP_PUBLIC_URL.rstrip('/')}/redefinir-senha?token={token}"`; `render_email("Redefinir senha", corpo, acao_url=link, acao_rotulo="Redefinir senha")`, corpo: `<p>Recebemos um pedido para redefinir a senha do usuário <strong>{html.escape(username)}</strong>.</p><p>O link vale por 30 minutos e só pode ser usado uma vez. Se não foi você, ignore este e-mail: sua senha continua a mesma.</p>`; `MailError` → `logger.error("Falha ao enviar o e-mail de redefinição de senha.")`, sem argumentos.
    - `POST /api/auth/reset` (mesmo rate limit) → senha em branco → 400 `"A nova senha não pode estar em branco."` **antes** de consumir o token; `consumir` → `None` → 400 `MSG_LINK_INVALIDO`; senão `update_password`, `invalidar_abertos`, `revogar_sessoes`, auditoria `redefinir_senha` com o contexto do próprio usuário (`depois={"senha": "redefinida", "sessoes": "encerradas"}`) → 200 `{"detail": MSG_SENHA_REDEFINIDA}`.
  - `fixtures_auth.py` (acrescentar): `cliente_com_rate_limit(valor: str)` (`@contextmanager` que copia o `_cliente_com_login_rate_limit` de `tests/test_auth.py`, recarregando `app.core.config`, `app.dependencies`, `app.routers.auth`, `app.routers.auth_senha` e `app.main`, e desfazendo no fim); fixture `usuario_com_email(criar_usuario, senha_padrao_teste, executar_sql) -> dict` (Técnico `tecnico.email` com e-mail `tecnico@revalle.com.br`).

- [ ] **Step 1: Testes** em `backend/tests/test_redefinir_senha.py` (`integration`; `limpar_outbox()` no início de cada teste):
  - `test_forgot_envia_email_com_link` — `monkeypatch.setattr(app.routers.auth_senha.settings, "APP_PUBLIC_URL", "https://patrimonio.exemplo/")`; `{"identificador": "tecnico.email"}` → 200 `MSG_FORGOT`; `outbox[-1].to == ["tecnico@revalle.com.br"]`; o texto contém `https://patrimonio.exemplo/redefinir-senha?token=` (barra final removida); token extraído com `re.search(r"token=([A-Za-z0-9_\-]+)", outbox[-1].text)`; no banco `token_hash == sha256_hex(token)` e o token puro não aparece em nenhuma coluna; `expira_em - criado_em == timedelta(minutes=30)`; auditoria `solicitar_redefinicao_senha` sem o token.
  - `test_forgot_por_email_sem_diferenciar_caixa` (`"TECNICO@Revalle.com.br"` → envia).
  - `test_forgot_usuario_inexistente_resposta_igual` — status e corpo idênticos aos do usuário existente; `outbox` vazio; `password_reset_tokens` vazia.
  - `test_forgot_usuario_sem_email_nao_envia` (`usuario_tecnico`, sem e-mail → mesma resposta, nada enviado).
  - `test_forgot_limite_de_3_por_hora` — 4 pedidos → 4 respostas iguais, `len(outbox) == 3`.
  - `test_forgot_falha_de_envio_responde_igual` — `monkeypatch` de `app.routers.auth_senha.send_mail` levantando `MailError("Não foi possível enviar o e-mail.")`; resposta 200 `MSG_FORGOT`; `caplog` sem o e-mail nem o token.
  - `test_reset_troca_a_senha` — com o token do e-mail, `{"token": t, "nova_senha": "NovaSenha#9"}` → 200 `MSG_SENHA_REDEFINIDA`; login com a senha antiga 401, com a nova 200.
  - `test_reset_token_reutilizado_recusado` (2º uso → 400 `MSG_LINK_INVALIDO`).
  - `test_reset_token_expirado_recusado` (`executar_sql("UPDATE password_reset_tokens SET expira_em = %s", (agora_utc - 1 min,))` → 400 `MSG_LINK_INVALIDO`).
  - `test_reset_token_desconhecido_recusado`.
  - `test_reset_senha_em_branco_nao_gasta_o_token` (`""` → 400 `"A nova senha não pode estar em branco."`; depois o mesmo token com senha válida → 200).
  - `test_reset_invalida_outros_links_abertos` (dois pedidos; usa o 2º; o 1º → 400).
  - `test_reset_revoga_sessoes` — login antes (guarda access e cookie); reset; `/me` com o access antigo → 401 `"Sessão encerrada. Faça login novamente."`; `/refresh` com o cookie antigo → 401.
  - `test_reset_grava_auditoria_sem_segredos` — linha `redefinir_senha` com `entidade_id == str(id)`; `json.dumps` da linha não contém o token, `"NovaSenha#9"` nem `"$2b$"`.
  - `test_forgot_tem_rate_limit` — `with cliente_com_rate_limit("3/minute") as c:` 3 × 200 e a 4ª → 429.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_redefinir_senha.py -q`.
- [ ] **Step 3: Implementar** conforme as Interfaces.
- [ ] **Step 4: Rodar:** `RBT tests/test_redefinir_senha.py tests/test_auth.py tests/test_migrations.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(auth): esqueci minha senha com link por e-mail de uso único`.

### Task 4: Serviço de 2FA (cifra, TOTP e códigos de recuperação)

**Files:**
- Create: `backend/app/services/mfa.py`, `backend/tests/test_mfa_servico.py`

**Interfaces:**
- Consumes: `security.sha256_hex` (Task 3).
- Produces (`app/services/mfa.py`; lê a chave sempre por `from app.core import config` → `config.settings.TOTP_ENCRYPTION_KEY` na hora da chamada, para os testes poderem trocá-la):
  - `ISSUER = "Revalle Patrimônio"`; `QTD_CODIGOS_RECUPERACAO = 10`.
  - Mensagens: `MSG_MFA_INDISPONIVEL`, `MSG_SEGREDO_ILEGIVEL` (textos da Decisão 2), `MSG_CODIGO_INVALIDO = "Código inválido."`, `MSG_MFA_JA_ATIVO = "A verificação em duas etapas já está ativa."`, `MSG_CADASTRO_NAO_INICIADO = "Inicie o cadastro do autenticador antes de confirmar."`.
  - Exceções com `status_code`: `MfaErro(Exception)` (400), `MfaIndisponivel(MfaErro)` (503, padrão `MSG_MFA_INDISPONIVEL`), `CodigoInvalido(MfaErro)` (400), `MfaJaAtivo(MfaErro)` (409), `CadastroNaoIniciado(MfaErro)` (400). Cada uma usa sua `MSG_*` como mensagem padrão.
  - `mfa_configurado() -> bool` (chave não vazia e `Fernet(chave)` aceita).
  - `cifrar_segredo(segredo: str) -> str`; `decifrar_segredo(cifrado: str) -> str` (sem chave → `MfaIndisponivel()`; `InvalidToken` → `MfaIndisponivel(MSG_SEGREDO_ILEGIVEL)`).
  - `gerar_segredo() -> str` (`pyotp.random_base32()`, 32 caracteres); `uri_provisionamento(segredo: str, username: str) -> str` (`pyotp.TOTP(segredo).provisioning_uri(name=username, issuer_name=ISSUER)`).
  - `parece_codigo_totp(codigo: str) -> bool` (sem espaços, exatamente 6 dígitos).
  - `verificar_codigo_totp(segredo: str, codigo: str, ultimo_passo: int | None, agora: datetime | None = None) -> int | None` — devolve o passo aceito:
    ```python
    codigo = re.sub(r"\s", "", codigo or "")
    if not re.fullmatch(r"\d{6}", codigo):
        return None
    totp = pyotp.TOTP(segredo)
    atual = totp.timecode(agora or datetime.now(timezone.utc))
    for passo in (atual - 1, atual, atual + 1):
        if ultimo_passo is not None and passo <= ultimo_passo:
            continue
        if hmac.compare_digest(totp.generate_otp(passo), codigo):
            return passo
    return None
    ```
  - `gerar_codigos_recuperacao(quantidade: int = QTD_CODIGOS_RECUPERACAO) -> list[str]` (`secrets.choice` no alfabeto da Decisão 8, `xxxx-xxxx-xxxx-xxxx`, sem repetição); `normalizar_codigo_recuperacao(codigo: str) -> str` (minúsculas, sem `-` e sem espaços); `hash_codigo_recuperacao(codigo: str) -> str` (`sha256_hex(normalizar_codigo_recuperacao(codigo))`).

- [ ] **Step 1: Testes** (`pytestmark = pytest.mark.unit`; a chave trocada com `monkeypatch.setattr(config.settings, "TOTP_ENCRYPTION_KEY", ...)`):
  - `test_mfa_configurado` (`""` → False; `"nao-e-fernet"` → False; `Fernet.generate_key().decode()` → True).
  - `test_cifrar_e_decifrar` (cifrado ≠ segredo; ida e volta igual).
  - `test_sem_chave_levanta_indisponivel` (`cifrar_segredo` → `MfaIndisponivel`, `status_code == 503`, `str(e) == MSG_MFA_INDISPONIVEL`).
  - `test_chave_trocada_levanta_segredo_ilegivel` (cifra com a chave A, decifra com a B → `MfaIndisponivel` com `MSG_SEGREDO_ILEGIVEL`).
  - `test_gerar_segredo_base32` (`re.fullmatch(r"[A-Z2-7]{32}", s)`).
  - `test_uri_provisionamento` (começa com `otpauth://totp/`, contém `f"secret={s}"` e `issuer=Revalle`).
  - `test_verificar_codigo_totp_janela` — `agora = datetime(2026, 11, 12, 12, 0, 0, tzinfo=timezone.utc)`, `p = TOTP(s).timecode(agora)`: código de `agora` → `p`; de `agora - 30 s` → `p - 1`; de `agora + 30 s` → `p + 1`; de `agora + 90 s` → `None`.
  - `test_verificar_codigo_totp_anti_replay` (`ultimo_passo=p` → `None`; `ultimo_passo=p - 1` → `p`).
  - `test_verificar_codigo_totp_aceita_espacos` (`f"{c[:3]} {c[3:]}"` → `p`; `"12345"` e `"abcdef"` → `None`).
  - `test_codigos_recuperacao` (10, sem repetição, `re.fullmatch(r"[a-z2-7]{4}(-[a-z2-7]{4}){3}", c)`).
  - `test_hash_codigo_normalizado` (`"ABCD-EFGH-IJKL-MNOP"`, `"abcdefghijklmnop"` e `" abcd efgh ijkl mnop "` têm o mesmo hash, igual a `sha256_hex("abcdefghijklmnop")`, com 64 caracteres).
  - `test_parece_codigo_totp` (`"123456"` e `"123 456"` → True; `"abcd-efgh-ijkl-mnop"` → False).
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_mfa_servico.py -q`.
- [ ] **Step 3: Implementar** `app/services/mfa.py`.
- [ ] **Step 4: Rodar** → PASS.
- [ ] **Step 5: Commit** `feat(2fa): serviço de TOTP com segredo cifrado e códigos de recuperação`.

### Task 5: Login em duas etapas

**Files:**
- Create: `backend/migrations/0202_mfa_recovery_codes.sql` (+ `.down.sql`: `DROP TABLE IF EXISTS mfa_recovery_codes;`), `backend/app/db/mfa_db.py`, `backend/app/services/segundo_fator.py`, `backend/tests/test_mfa_login.py`
- Modify: `backend/app/core/security.py`, `backend/app/schemas/auth.py`, `backend/app/routers/auth.py`, `backend/app/dependencies.py` (acrescentar `get_mfa_db`), `backend/app/main.py` (rótulos `"codigo": "Código"`, `"mfa_token": "Etapa de verificação"`), `backend/tests/fixtures_auth.py`

**Interfaces:**
- Consumes: Task 4 inteira; `dados_do_token`, `validar_sessao_ativa`, `MSG_SESSAO_ENCERRADA` (Task 2); `TokenDBManager.revoke/is_revoked`; `contexto_auditoria_de` (Task 3).
- Produces:
  - Tabela:
    ```sql
    CREATE TABLE IF NOT EXISTS mfa_recovery_codes (
        id INT AUTO_INCREMENT PRIMARY KEY,
        usuario_id INT NOT NULL,
        code_hash CHAR(64) NOT NULL,
        usado_em DATETIME NULL,
        UNIQUE KEY uq_mfa_recovery (usuario_id, code_hash),
        CONSTRAINT fk_mfa_recovery_usuario FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
    ) CHARACTER SET utf8mb4;
    ```
  - `MfaDBManager` (`app/db/mfa_db.py`): `estado(usuario_id: int) -> dict | None` (`totp_secret_cifrado, totp_ativo, totp_ultimo_passo`); `salvar_segredo_pendente(usuario_id: int, cifrado: str) -> None` (grava o segredo com `totp_ativo=0`, `totp_ultimo_passo=NULL`); `ativar(usuario_id: int, hashes: list[str], passo: int) -> None` (numa transação: `totp_ativo=1`, `totp_ultimo_passo=passo`, apaga e insere os códigos); `registrar_passo(usuario_id: int, passo: int) -> bool` (`UPDATE ... WHERE id=%s AND (totp_ultimo_passo IS NULL OR totp_ultimo_passo < %s)`, `rowcount == 1`); `consumir_codigo(usuario_id: int, code_hash: str) -> bool` (`UPDATE ... SET usado_em=%s WHERE usuario_id=%s AND code_hash=%s AND usado_em IS NULL`, `rowcount == 1`); `contar_codigos_restantes(usuario_id: int) -> int`.
  - `security.py`: `MFA_VERIFICAR_MINUTOS = 5`, `MFA_CADASTRO_MINUTOS = 15`, `create_mfa_token(data: dict, proposito: Literal["verificar", "cadastrar"]) -> str` (claims de `data` + `type="mfa_pendente"`, `proposito`, `jti=uuid4().hex`, `exp`).
  - `segundo_fator.verificar_segundo_fator(db: MfaDBManager, usuario: dict, codigo: str, ctx: AuditContext) -> int | None` — `parece_codigo_totp` → decifra (pode levantar `MfaIndisponivel`), `verificar_codigo_totp(..., estado["totp_ultimo_passo"])` e `registrar_passo`; falha → `CodigoInvalido`; devolve `None`. Senão, código de recuperação: `consumir_codigo(id, hash_codigo_recuperacao(codigo))` ou `CodigoInvalido`; audita `usar_codigo_recuperacao` com `depois={"codigos_restantes": n}`; devolve `n`.
  - Schemas: `LoginResponse(etapa: Literal["ok", "mfa", "cadastro_2fa"], access_token: Optional[str] = None, token_type: str = "bearer", mfa_token: Optional[str] = None, expira_em_segundos: Optional[int] = None, codigos_recuperacao_restantes: Optional[int] = None, codigos_recuperacao: Optional[list[str]] = None)`; `MfaVerificarRequest(mfa_token: str, codigo: str)`.
  - `routers/auth.py`: `MSG_ETAPA_EXPIRADA = "Etapa de verificação expirada. Entre novamente."`; `_emitir_sessao(response: Response, usuario: dict, **extras) -> LoginResponse` (access + cookie de refresh com `dados_do_token`, `etapa="ok"`); `_decodificar_mfa(mfa_token: str, proposito: str, token_db: TokenDBManager, user_db: UserDBManager) -> tuple[dict, dict]` (payload, usuario) — token inválido/expirado, `type`/`proposito` errado, `jti` revogado, usuário inexistente ou `sv` diferente → `HTTPException(400, MSG_ETAPA_EXPIRADA)`; `_revogar_mfa(payload: dict, token_db: TokenDBManager) -> None`; `_erro_mfa(exc: MfaErro) -> HTTPException` (`exc.status_code`, `str(exc)`).
  - `POST /api/auth/login` → `LoginResponse`: credencial errada → 401 (como hoje); `totp_ativo` → `etapa="mfa"`, `mfa_token=create_mfa_token(dados_do_token(user), "verificar")`, `expira_em_segundos=300`, **sem** cookie; senão `_emitir_sessao`.
  - `POST /api/auth/mfa/verificar` (`@limiter.limit(settings.LOGIN_RATE_LIMIT)`) → `_decodificar_mfa(..., "verificar")` → `verificar_segundo_fator` (`MfaErro` → `_erro_mfa`) → `_revogar_mfa` → `_emitir_sessao(..., codigos_recuperacao_restantes=n)`.
  - `fixtures_auth.py` (acrescentar): fixture `chave_totp(monkeypatch) -> str` (`Fernet.generate_key().decode()` em `config.settings.TOTP_ENCRYPTION_KEY`); `@dataclass Dados2fa(segredo: str, codigos: list[str])`; fixture `ativar_2fa_direto(chave_totp) -> Callable[[int], Dados2fa]` (`salvar_segredo_pendente` + `ativar(id, hashes, passo=0)`); `codigo_totp(segredo: str, passos: int = 0) -> str` (`pyotp.TOTP(segredo).at(datetime.now(timezone.utc) + timedelta(seconds=30 * passos))` — use `passos=1` num segundo login do mesmo usuário no mesmo teste, por causa do anti-replay); `login_com_2fa(client, username: str, senha: str, codigo: str) -> str` (login → afirma `etapa == "mfa"` → verificar → afirma 200 → devolve o access token); fixture `gestor_2fa(_fastapi_app, client, usuario_gestor, senha_padrao_teste, ativar_2fa_direto) -> dict` (`{"usuario", "segredo", "codigos", "client"}`, com um `TestClient` novo autenticado pelo access de `login_com_2fa`). **Regra para os testes da F2:** com `chave_totp` ligada, não use `client_gestor`/`token_gestor` do `conftest.py` (a partir da Task 6 o login do Gestor sem 2FA não devolve access token) — use `gestor_2fa["client"]`; `client_tecnico`/`client_aprendiz` continuam valendo.

- [ ] **Step 1: Testes** em `backend/tests/test_mfa_login.py` (`integration`, usam `chave_totp`):
  - `test_migracao_0202_ida_e_volta` (`ida_e_volta_migracao(202)`, tabela existe no fim).
  - `test_login_com_2fa_pede_codigo` — usuário com `ativar_2fa_direto` → 200 `etapa == "mfa"`, `access_token is None`, `"refresh_token" not in resp.cookies`, `expira_em_segundos == 300`; claims do `mfa_token`: `type == "mfa_pendente"`, `proposito == "verificar"`, `exp` entre agora+295 s e agora+305 s.
  - `test_mfa_token_nao_serve_como_access` (`/api/auth/me` com Bearer `mfa_token` → 401).
  - `test_verificar_com_totp_emite_sessao` (`etapa == "ok"`, access autentica em `/me`, cookie `refresh_token` presente, `codigos_recuperacao_restantes is None`).
  - `test_verificar_codigo_errado` (código ≠ do atual → 400 `"Código inválido."`).
  - `test_verificar_codigo_reutilizado_recusado` — login + verificar com `c = codigo_totp(s)`; novo login + verificar com o mesmo `c` → 400 `"Código inválido."`.
  - `test_verificar_com_codigo_de_recuperacao` — com `codigos[0].upper().replace("-", "")` → 200, `codigos_recuperacao_restantes == 9`; auditoria `usar_codigo_recuperacao` com `depois == {"codigos_restantes": 9}` e sem o código; novo login com `codigos[0]` → 400.
  - `test_mfa_token_usado_nao_vale_de_novo` (mesmo `mfa_token` depois do sucesso, com `codigo_totp(s, passos=1)` → 400 `"Etapa de verificação expirada. Entre novamente."`).
  - `test_mfa_token_expirado` (`jose_jwt.encode` com `type="mfa_pendente"`, `proposito="verificar"`, `exp` no passado → 400 `MSG_ETAPA_EXPIRADA`).
  - `test_access_token_nao_serve_como_mfa_token` (access normal no campo `mfa_token` → 400 `MSG_ETAPA_EXPIRADA`).
  - `test_sessoes_revogadas_entre_as_etapas_invalidam_o_mfa_token` (login → `revogar_sessoes(id)` → verificar → 400 `MSG_ETAPA_EXPIRADA`).
  - `test_reset_de_senha_nao_desativa_2fa` (token criado direto com `PasswordResetDBManager().criar(id, sha256_hex("tok-teste"), ...)`; `POST /api/auth/reset` `{"token": "tok-teste", "nova_senha": "NovaSenha#9"}` → 200; login com a nova senha → `etapa == "mfa"`).
  - `test_verificar_tem_rate_limit` — `with cliente_com_rate_limit("3/minute") as c:` 3 × `{"mfa_token": "x", "codigo": "000000"}` → 400 e a 4ª → 429.
  - `test_usuario_sem_2fa_continua_com_login_direto` (Técnico sem 2FA → `etapa == "ok"` com `access_token`).
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_mfa_login.py -q`.
- [ ] **Step 3: Implementar** conforme as Interfaces.
- [ ] **Step 4: Rodar:** `RBT tests/test_mfa_login.py tests/test_auth.py tests/test_sessoes.py tests/test_redefinir_senha.py -q` → PASS (o `test_login_com_credenciais_validas_retorna_access_token` atual continua passando: sem chave, `etapa="ok"` traz `access_token` e `token_type`).
- [ ] **Step 5: Commit** `feat(auth): login em duas etapas com código TOTP ou de recuperação`.

### Task 6: Cadastro obrigatório do Gestor e regra da chave ausente

**Files:**
- Modify: `backend/app/services/sessao.py`, `backend/app/services/segundo_fator.py`, `backend/app/schemas/auth.py`, `backend/app/routers/auth.py`, `backend/app/main.py` (acrescentar no fim um handler de startup), `backend/.env.example` (acrescentar 1 linha de comentário logo abaixo de `TOTP_ENCRYPTION_KEY=`), `backend/tests/conftest.py` (em `pytest_configure`, acrescentar `os.environ["TOTP_ENCRYPTION_KEY"] = ""` logo depois de `SCHEDULER_ENABLED` — mesma proteção do `MAIL_BACKEND`: um `backend/.env` com chave não pode derrubar os fixtures `client_gestor` de todas as fases)
- Test: `backend/tests/test_mfa_cadastro_gestor.py`

**Interfaces:**
- Consumes: Tasks 2, 4 e 5.
- Produces:
  - `sessao.MSG_CADASTRO_OBRIGATORIO = "Cadastre a verificação em duas etapas para continuar. Entre novamente."`; `sessao.exige_cadastro_2fa(usuario: dict) -> bool` (`role == "Gestor"` e não `totp_ativo` e `mfa_configurado()`); `validar_sessao_ativa` passa a levantar 401 `MSG_CADASTRO_OBRIGATORIO` quando `exige_cadastro_2fa(estado)` — vale para `get_current_user` e para `/refresh` (sessões anteriores ao cadastro, inclusive tokens sem `sv`, ficam bloqueadas).
  - `segundo_fator.iniciar_cadastro(db: MfaDBManager, usuario: dict) -> tuple[str, str]` (segredo, uri) — `totp_ativo` → `MfaJaAtivo`; sem chave → `MfaIndisponivel`; grava `cifrar_segredo(gerar_segredo())` como pendente (sobrescreve um pendente anterior).
  - `segundo_fator.confirmar_cadastro(db: MfaDBManager, usuario: dict, codigo: str, ctx: AuditContext) -> list[str]` — `totp_ativo` → `MfaJaAtivo`; sem segredo pendente → `CadastroNaoIniciado`; `verificar_codigo_totp(segredo, codigo, None)` → `None` → `CodigoInvalido`; gera os 10 códigos, `ativar(id, hashes, passo)`, audita `ativar_2fa` (`antes={"totp_ativo": False}`, `depois={"totp_ativo": True}`), devolve os códigos em texto (única vez).
  - Schemas: `MfaTokenRequest(mfa_token: str)`, `MfaConfirmarRequest(mfa_token: str, codigo: str)`, `CadastroTotpResponse(segredo: str, otpauth_uri: str)`.
  - Decisão do `POST /api/auth/login`, nesta ordem: (1) credencial errada → 401; (2) `totp_ativo` → `etapa="mfa"` (mesmo sem chave — ver Decisão 2); (3) `exige_cadastro_2fa(user)` → `etapa="cadastro_2fa"`, `mfa_token=create_mfa_token(..., "cadastrar")`, `expira_em_segundos=900`, sem cookie; (4) `etapa="ok"`.
  - `POST /api/auth/mfa/cadastro/iniciar` body `MfaTokenRequest` → `CadastroTotpResponse` (`_decodificar_mfa(..., "cadastrar")`; não revoga o token).
  - `POST /api/auth/mfa/cadastro/confirmar` (`@limiter.limit(settings.LOGIN_RATE_LIMIT)`) body `MfaConfirmarRequest` → `confirmar_cadastro` → `_revogar_mfa` → `_emitir_sessao(..., codigos_recuperacao=codigos)`.
  - `app.main.avisar_2fa_sem_chave() -> None` (`@app.on_event("startup")`, `logger.warning("TOTP_ENCRYPTION_KEY ausente ou inválida: verificação em duas etapas desligada; o Gestor entra só com a senha.")` quando `not mfa_configurado()`).
  - `.env.example`: `# Vazia = 2FA desligado (o Gestor entra só com a senha e os endpoints de 2FA respondem 503). Trocar a chave invalida os autenticadores já cadastrados.`

- [ ] **Step 1: Testes** em `backend/tests/test_mfa_cadastro_gestor.py` (`integration`):
  - `test_gestor_sem_2fa_recebe_etapa_cadastro` (com `chave_totp`: `etapa == "cadastro_2fa"`, sem `access_token`, sem cookie, `expira_em_segundos == 900`, claim `proposito == "cadastrar"`).
  - `test_token_de_cadastro_nao_acessa_a_api` (Bearer `mfa_token` em `/api/items` → 401).
  - `test_cadastro_completo_emite_sessao_e_codigos` — iniciar → `otpauth_uri` começa com `otpauth://totp/` e contém `f"secret={segredo}"`; no banco `totp_secret_cifrado != segredo` e `decifrar_segredo(...) == segredo`; confirmar com `codigo_totp(segredo)` → `etapa == "ok"`, access autentica `/api/items`, cookie presente, `len(codigos_recuperacao) == 10` e sem repetição; `mfa_recovery_codes` tem 10 hashes e nenhum igual a um código; auditoria `ativar_2fa` sem o segredo e sem códigos no `json.dumps`.
  - `test_confirmar_com_codigo_errado` (400 `"Código inválido."`; `totp_ativo` continua 0).
  - `test_confirmar_sem_iniciar` (400 `"Inicie o cadastro do autenticador antes de confirmar."`).
  - `test_token_de_cadastro_reutilizado_apos_confirmar` (400 `"Etapa de verificação expirada. Entre novamente."`).
  - `test_token_de_verificacao_nao_serve_para_cadastro` (`mfa_token` de `proposito="verificar"` no iniciar → 400 `MSG_ETAPA_EXPIRADA`).
  - `test_gestor_sem_2fa_com_sessao_antiga_e_bloqueado` — com `chave_totp`: `create_access_token({"sub": str(id), "username": ..., "role": "Gestor"})` em `/api/items` → 401 `MSG_CADASTRO_OBRIGATORIO`; cookie `create_refresh_token(...)` em `/api/auth/refresh` → 401 com a mesma mensagem.
  - `test_tecnico_sem_2fa_nao_e_obrigado` (com chave → `etapa == "ok"`).
  - `test_chave_vazia_gestor_entra_sem_2fa` (sem `chave_totp` → `etapa == "ok"` e `/api/items` 200).
  - `test_chave_vazia_endpoints_de_cadastro_respondem_503` (token de cadastro gerado com a chave ligada; `monkeypatch` da chave para `""`; iniciar → 503 com `MSG_MFA_INDISPONIVEL`).
  - `test_chave_vazia_usuario_com_2fa_entra_com_codigo_de_recuperacao` — ativa com chave; chave → `""`; login → `etapa == "mfa"`; código TOTP → 503 `MSG_MFA_INDISPONIVEL`; código de recuperação → 200 `etapa == "ok"`.
  - `test_chave_trocada_responde_segredo_ilegivel` (ativa com a chave A; troca para uma chave B válida; TOTP → 503 `MSG_SEGREDO_ILEGIVEL`).
  - `test_aviso_de_chave_ausente_no_boot` (`caplog` + `app.main.avisar_2fa_sem_chave()` → `"TOTP_ENCRYPTION_KEY ausente ou inválida"` no log; com chave, nada).
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_mfa_cadastro_gestor.py -q`.
- [ ] **Step 3: Implementar** conforme as Interfaces.
- [ ] **Step 4: Rodar a suíte inteira:** `RBT -q -m "not mudanca_esperada"` → verde (prova que os fixtures `client_gestor` das outras suítes seguem funcionando sem chave).
- [ ] **Step 5: Commit** `feat(2fa): cadastro obrigatório do autenticador para Gestor`.

### Task 7: Perfil e 2FA opcional

**Files:**
- Create: `backend/app/routers/perfil.py`, `backend/app/schemas/perfil.py`, `backend/tests/test_perfil.py`
- Modify: `backend/app/db/mfa_db.py`, `backend/app/services/segundo_fator.py`, `backend/app/main.py` (acrescentar `from app.routers import perfil`, `app.include_router(perfil.router)`, rótulo `"senha": "Senha"`)

**Interfaces:**
- Consumes: Tasks 4–6; `get_current_user`, `get_audit_context`, `verify_password`.
- Produces:
  - `MfaDBManager.substituir_codigos(usuario_id: int, hashes: list[str]) -> None`; `MfaDBManager.desativar(usuario_id: int) -> None` (segredo e último passo `NULL`, `totp_ativo=0`, apaga os códigos; não mexe em `sessao_versao`).
  - `segundo_fator.regenerar_codigos(db: MfaDBManager, usuario: dict, ctx: AuditContext) -> list[str]` (sem 2FA ativo → `MfaErro("A verificação em duas etapas não está ativa.")`; audita `regenerar_codigos_2fa` com `depois={"codigos": 10}`); `segundo_fator.desativar_2fa(db: MfaDBManager, usuario: dict, ctx: AuditContext) -> None` (Gestor → `MfaErro("Para o Gestor a verificação em duas etapas é obrigatória.")`; sem 2FA → `MfaErro("A verificação em duas etapas não está ativa.")`; audita `desativar_2fa`).
  - Schemas: `PerfilResponse(id: int, username: str, role: str, email: Optional[str], totp_ativo: bool, mfa_obrigatorio: bool, mfa_disponivel: bool, codigos_recuperacao_restantes: int)`; `CodigoRequest(codigo: str)`; `SenhaRequest(senha: str)`; `CodigosRecuperacaoResponse(codigos_recuperacao: list[str])`.
  - Router `/api/perfil` (todos os papéis autenticados): `GET ""` → `PerfilResponse` (`mfa_obrigatorio = role == "Gestor"`, `mfa_disponivel = mfa_configurado()`); `POST /2fa/iniciar` → `CadastroTotpResponse`; `POST /2fa/confirmar` body `CodigoRequest` → `CodigosRecuperacaoResponse`; `POST /2fa/desativar` body `SenhaRequest` → `MensagemResponse("Verificação em duas etapas desativada.")`; `POST /2fa/codigos` body `SenhaRequest` → `CodigosRecuperacaoResponse`. Os dois com senha têm `@limiter.limit(settings.LOGIN_RATE_LIMIT)`; senha errada → 400 `"Senha incorreta."`. `MfaErro` → `HTTPException(exc.status_code, str(exc))`.

- [ ] **Step 1: Testes** em `backend/tests/test_perfil.py` (`integration`):
  - `test_perfil_exige_login` (401).
  - `test_get_perfil` (Técnico com e-mail, chave ligada → `email == "tecnico@revalle.com.br"`, `totp_ativo False`, `mfa_obrigatorio False`, `mfa_disponivel True`, `codigos_recuperacao_restantes == 0`).
  - `test_perfil_sem_chave_indica_indisponivel` (`mfa_disponivel False`; `POST /2fa/iniciar` → 503 `MSG_MFA_INDISPONIVEL`).
  - `test_tecnico_ativa_2fa_no_perfil` (iniciar → confirmar com `codigo_totp` → 10 códigos; `GET` mostra `totp_ativo True` e `codigos_recuperacao_restantes == 10`; próximo login → `etapa == "mfa"`; auditoria `ativar_2fa`).
  - `test_iniciar_com_2fa_ativo_409` (`"A verificação em duas etapas já está ativa."`).
  - `test_confirmar_sem_iniciar_400`; `test_confirmar_codigo_errado_400` (`"Código inválido."`).
  - `test_desativar_exige_senha` (errada → 400 `"Senha incorreta."`; certa → 200, `totp_ativo` 0, `mfa_recovery_codes` sem linhas do usuário, auditoria `desativar_2fa`, próximo login `etapa == "ok"`).
  - `test_gestor_nao_desativa` (`gestor_2fa["client"]` → 400 `"Para o Gestor a verificação em duas etapas é obrigatória."`).
  - `test_regenerar_codigos` (senha certa → 10 novos; o código antigo não entra mais; um novo entra; auditoria `regenerar_codigos_2fa` sem códigos no `json.dumps`).
  - `test_regenerar_sem_2fa_400` (`"A verificação em duas etapas não está ativa."`).
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_perfil.py -q`.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar:** `RBT tests/test_perfil.py tests/test_mfa_login.py tests/test_mfa_cadastro_gestor.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(perfil): página de perfil com 2FA opcional para os demais papéis`.

### Task 8: Gestor zera o 2FA de outro usuário e comando `reset_2fa`

**Files:**
- Create: `backend/scripts/reset_2fa.py`, `backend/tests/test_zerar_2fa.py`
- Modify: `backend/app/db/mfa_db.py`, `backend/app/services/segundo_fator.py`, `backend/app/routers/users.py`

**Interfaces:**
- Consumes: Tasks 2, 5–7; `AuditContext.sistema()`.
- Produces:
  - `MfaDBManager.zerar(usuario_id: int) -> None` (o mesmo que `desativar` + `sessao_versao = sessao_versao + 1`, numa transação).
  - `segundo_fator.zerar_2fa(db: MfaDBManager, usuario: dict, ctx: AuditContext, origem: str) -> None` (audita `zerar_2fa` com `antes={"totp_ativo": <bool>}`, `depois={"totp_ativo": False, "origem": origem}`).
  - `DELETE /api/users/{user_id}/2fa` (`gestor_only`, `ctx` da F0) → 200 `{"detail": "Verificação em duas etapas zerada."}`; próprio usuário → 400 `"Você não pode zerar a sua própria verificação em duas etapas. Peça a outro Gestor ou use o comando no servidor."`; inexistente → 404 `"Usuário não encontrado."`; sem 2FA ativo → 400 `"Este usuário não tem a verificação em duas etapas ativa."`; `origem="tela de usuários"`. Não precisa da chave TOTP.
  - `scripts/reset_2fa.py`: `main(argv: list[str] | None = None) -> int` (`argparse`, argumento `usuario`; uso `python -m scripts.reset_2fa <usuario>` a partir de `backend/`; docstring com o uso e o aviso "rode só no servidor"); inexistente → imprime `"Usuário '<u>' não encontrado."` e devolve 1; sem 2FA → `"O usuário '<u>' não tem 2FA ativo. Nada a fazer."` e 0; senão `zerar_2fa(..., AuditContext.sistema(), origem="script reset_2fa")`, imprime `"2FA de '<u>' zerado e sessões encerradas. No próximo login, um Gestor cadastra um novo autenticador."` e devolve 0.

- [ ] **Step 1: Testes** em `backend/tests/test_zerar_2fa.py` (`integration`; quem age é `gestor_2fa["client"]`, porque a chave está ligada):
  - `test_gestor_zera_2fa_de_outro` (alvo Técnico com `ativar_2fa_direto` → 200; banco: `totp_ativo` 0, segredo `NULL`, 0 códigos; próximo login do alvo `etapa == "ok"`; auditoria `zerar_2fa` com `entidade_id == str(alvo)`, `depois["origem"] == "tela de usuários"` e sem segredo).
  - `test_zerar_2fa_encerra_sessoes_do_alvo` (access do alvo obtido antes → `/me` 401 `"Sessão encerrada. Faça login novamente."`).
  - `test_gestor_zerado_volta_ao_cadastro` (alvo Gestor com 2FA → zerado → login → `etapa == "cadastro_2fa"`).
  - `test_zerar_2fa_exige_gestor` (Técnico → 403 `"Acesso restrito ao Gestor."`).
  - `test_zerar_proprio_2fa_recusado`; `test_zerar_2fa_usuario_sem_2fa_400`; `test_zerar_2fa_usuario_inexistente_404`.
  - `test_listagem_mostra_2fa_zerado` (`GET /api/users` → alvo com `totp_ativo False`).
  - `test_script_reset_2fa_zera_e_audita` (`main(["gestor.teste"]) == 0` para um Gestor com 2FA; banco zerado; `sessao_versao` +1; auditoria com `username == "sistema"` e `depois["origem"] == "script reset_2fa"`; `capsys` contém `"zerado"`). O script não lê a chave: o teste esvazia `config.settings.TOTP_ENCRYPTION_KEY` antes de chamar `main` e ele funciona igual.
  - `test_script_reset_2fa_usuario_inexistente` (`main(["nao.existe"]) == 1`; saída `"Usuário 'nao.existe' não encontrado."`).
  - `test_script_reset_2fa_sem_2fa` (devolve 0 e `"não tem 2FA ativo"`).
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_zerar_2fa.py -q`.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar a suíte inteira:** `RBT -q -m "not mudanca_esperada"` → verde.
- [ ] **Step 5: Commit** `feat(2fa): Gestor zera o 2FA de outro usuário e comando reset_2fa no servidor`.

### Task 9: Contratos de autenticação no frontend (API e AuthContext)

**Files:**
- Create: `frontend/src/api/perfil.ts`, `frontend/src/lib/auth-errors.ts`, `frontend/src/lib/auth-errors.test.ts`, `frontend/src/contexts/AuthContext.test.tsx`
- Modify: `frontend/src/api/auth.ts`, `frontend/src/contexts/AuthContext.tsx`, `frontend/src/pages/LoginPage.test.tsx` (o mock de login do teste "credencial válida" passa a responder `{ etapa: 'ok', access_token: 'tok-abc', token_type: 'bearer' }`)

**Interfaces:**
- Consumes: contratos HTTP das Tasks 3, 5–7 (MSW).
- Produces:
  - `api/auth.ts`: `type EtapaLogin = 'ok' | 'mfa' | 'cadastro_2fa'`; `interface LoginResponse { etapa: EtapaLogin; access_token?: string | null; token_type?: string; mfa_token?: string | null; expira_em_segundos?: number | null; codigos_recuperacao_restantes?: number | null; codigos_recuperacao?: string[] | null }`; `interface CadastroTotp { segredo: string; otpauth_uri: string }`; `login(username, password): Promise<LoginResponse>`; `verificarMfa(mfaToken: string, codigo: string): Promise<LoginResponse>` (`POST /auth/mfa/verificar` `{ mfa_token, codigo }`); `iniciarCadastroMfa(mfaToken: string): Promise<CadastroTotp>` (`/auth/mfa/cadastro/iniciar` `{ mfa_token }`); `confirmarCadastroMfa(mfaToken: string, codigo: string): Promise<LoginResponse>`; `esqueciSenha(identificador: string): Promise<{ detail: string }>` (`/auth/forgot`); `redefinirSenha(token: string, novaSenha: string): Promise<{ detail: string }>` (`/auth/reset` `{ token, nova_senha }`); `logout`, `getMe` inalterados.
  - `api/perfil.ts`: `interface Perfil { id: number; username: string; role: string; email: string | null; totp_ativo: boolean; mfa_obrigatorio: boolean; mfa_disponivel: boolean; codigos_recuperacao_restantes: number }`; `getPerfil(): Promise<Perfil>`; `iniciarCadastro2fa(): Promise<CadastroTotp>`; `confirmarCadastro2fa(codigo: string): Promise<{ codigos_recuperacao: string[] }>`; `desativar2fa(senha: string): Promise<{ detail: string }>`; `regenerarCodigos(senha: string): Promise<{ codigos_recuperacao: string[] }>`.
  - `lib/auth-errors.ts`: `MSG_MUITAS_TENTATIVAS = 'Muitas tentativas. Aguarde um minuto e tente de novo.'`; `mensagemErroAuth(error: unknown, fallback: string): string` (status 429 → `MSG_MUITAS_TENTATIVAS`, pois o corpo do slowapi não tem `detail`; senão `getErrorMessage(error, fallback)`).
  - `AuthContext`: `login(username, password): Promise<LoginResponse>` — com `etapa === 'ok'` chama `concluirLogin(access_token)`; nas outras etapas não guarda token nem usuário. Novo `concluirLogin(accessToken: string): Promise<void>` (`setAccessToken` + `getMe` + `setUser`). O resto (refresh silencioso, sessão expirada, `logout`, `hasRole`) não muda.

- [ ] **Step 1: Testes:**
  - `AuthContext.test.tsx` (`renderHook` com wrapper `<MemoryRouter><AuthProvider>`; `POST /api/auth/refresh` → 401 em todos): `login com etapa ok guarda o usuário` (`/auth/me` → `{ id: 1, username: 'gestor', role: 'Gestor' }`; `result.current.user?.username === 'gestor'`); `login com etapa mfa não chama /auth/me` (contador do handler de `/auth/me` em 0; `user` null; retorno com `mfa_token: 'mfa-1'`); `concluirLogin usa o token recebido` (o handler de `/auth/me` vê `Authorization: Bearer tok-2`).
  - `auth-errors.test.ts`: `AxiosError` com `response.status = 429` e `data = { error: 'Rate limit exceeded: 5 per 1 minute' }` → `'Muitas tentativas. Aguarde um minuto e tente de novo.'`; status 400 com `detail: 'Código inválido.'` → `'Código inválido.'`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/contexts/AuthContext.test.tsx src/lib/auth-errors.test.ts`.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run src/contexts src/lib src/pages/LoginPage.test.tsx` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(auth): contratos de login em etapas, redefinição de senha e perfil no frontend`.

### Task 10: Componentes do autenticador (QR code, cadastro e códigos de recuperação)

**Files:**
- Create: `frontend/src/components/auth/QrCodeOtpauth.tsx`, `frontend/src/components/auth/CadastroAutenticador.tsx`, `frontend/src/components/auth/CodigosRecuperacao.tsx` e os testes `*.test.tsx` ao lado
- Modify: `frontend/package.json`, `frontend/package-lock.json` (`npm install qrcode@1.5.4` e `npm install -D @types/qrcode@1.5.5`)

**Interfaces:**
- Consumes: `CadastroTotp` (Task 9), `mensagemErroAuth`, `Button` (`loading`), `Input`, `Label`, `Checkbox` (`components/ui/checkbox.tsx`).
- Produces:
  - `QrCodeOtpauth({ uri, size = 192 }: { uri: string; size?: number })` — `await import('qrcode')` (fora do bundle principal) e `toString(uri, { type: 'svg', margin: 1 })`; renderiza `<img alt="QR code para o aplicativo autenticador" src={'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg)} width={size} height={size}>` sobre fundo branco (contraste nos dois temas); antes de pronto, um placeholder com `aria-busy`. Sem `dangerouslySetInnerHTML`. Justificativa da dependência (vai no relatório): gerar QR à mão é inviável e a F6 já usa `qrcode` para as etiquetas — as duas fases precisam pinar a mesma versão.
  - `CadastroAutenticador({ iniciar, confirmar, onConcluido }: { iniciar: () => Promise<CadastroTotp>; confirmar: (codigo: string) => Promise<string[]>; onConcluido: (codigos: string[]) => void })` — chama `iniciar` ao montar (uma vez); mostra "Gerando QR code...", depois o QR, o texto "Não consegue ler o QR code? Digite esta chave no aplicativo:" com a chave em grupos de 4 separados por espaço, o campo "Código de 6 dígitos" (`inputMode="numeric"`, `autoComplete="one-time-code"`) e o botão "Confirmar" (`loading`). Erro de `iniciar` ou `confirmar` → `mensagemErroAuth(err, 'Não foi possível concluir o cadastro.')` num `role="alert"`.
  - `CodigosRecuperacao({ codigos, onContinuar, rotuloContinuar = 'Continuar' }: { codigos: string[]; onContinuar: () => void; rotuloContinuar?: string })` — texto "Guarde estes 10 códigos em local seguro. Cada um serve uma única vez para entrar se você perder o celular. Eles não serão mostrados de novo."; lista em fonte mono; botões "Copiar códigos" (`navigator.clipboard.writeText(codigos.join('\n'))`) e "Baixar .txt" (`codigos-recuperacao-patrimonio.txt`); checkbox "Guardei os códigos em local seguro"; botão `rotuloContinuar` desabilitado até marcar.

- [ ] **Step 1: Testes:**
  - `QrCodeOtpauth.test.tsx`: `await screen.findByAltText('QR code para o aplicativo autenticador')` com `src` começando por `data:image/svg+xml`.
  - `CadastroAutenticador.test.tsx`: `iniciar` resolve `{ segredo: 'JBSWY3DPEHPK3PXPJBSWY3DPEHPK3PXP', otpauth_uri: 'otpauth://totp/Revalle:x?secret=JBSWY3DPEHPK3PXPJBSWY3DPEHPK3PXP' }` → mostra o QR e `'JBSW Y3DP EHPK 3PXP JBSW Y3DP EHPK 3PXP'`; digitar `123456` e "Confirmar" chama `confirmar('123456')` e depois `onConcluido` com os códigos; `confirmar` rejeitando com `AxiosError` 400 `detail: 'Código inválido.'` mostra `'Código inválido.'` e não chama `onConcluido`; `iniciar` rejeitando com 503 mostra o `detail` e nenhum QR.
  - `CodigosRecuperacao.test.tsx`: 10 códigos visíveis; "Continuar" desabilitado até marcar o checkbox; "Copiar códigos" → `await navigator.clipboard.readText()` (clipboard falso do `userEvent.setup()`) igual a `codigos.join('\n')`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/components/auth`.
- [ ] **Step 3: Instalar as dependências e implementar.**
- [ ] **Step 4: Rodar** `npx vitest run src/components/auth`, `npx tsc --noEmit` e `npm run build` (o `qrcode` vai para um chunk separado) → verdes.
- [ ] **Step 5: Commit** `feat(2fa): componentes de QR code, cadastro do autenticador e códigos de recuperação`.

### Task 11: Tela de login em etapas

**Files:**
- Create: `frontend/src/components/auth/AuthShell.tsx`
- Modify: `frontend/src/pages/LoginPage.tsx`, `frontend/src/pages/LoginPage.test.tsx`

**Interfaces:**
- Consumes: `useAuth().login/concluirLogin`, `verificarMfa`, `iniciarCadastroMfa`, `confirmarCadastroMfa` (Task 9); `CadastroAutenticador`, `CodigosRecuperacao` (Task 10); `mensagemErroAuth`; `toast`.
- Produces:
  - `AuthShell({ titulo, descricao, children }: { titulo: string; descricao?: string; children: ReactNode })` — o layout atual do `LoginPage` (painel institucional com o 3D sob demanda, card com "Acesso Restrito", rodapé "Autenticação Segura JWT"), extraído sem mudança visual; o `LoginPage` e as páginas da Task 12 o usam.
  - `LoginPage` com estado `type Etapa = { tipo: 'senha' } | { tipo: 'mfa'; mfaToken: string } | { tipo: 'cadastro'; mfaToken: string } | { tipo: 'codigos'; codigos: string[]; accessToken: string }`:
    - `senha`: como hoje ("Entrar no Sistema", campos "Usuário" e "Senha", botão "Entrar" com `loading`), link "Esqueci minha senha" → `/esqueci-senha`; erro → `mensagemErroAuth(err, 'Usuário ou senha inválidos.')`. Resposta `ok` → `navigate('/', { replace: true })`; `mfa` → etapa `mfa`; `cadastro_2fa` → etapa `cadastro`.
    - `mfa`: título "Verificação em duas etapas", texto "Digite o código de 6 dígitos do aplicativo autenticador ou um código de recuperação.", campo "Código", botão "Verificar" (`loading`), botão "Voltar" (volta para `senha` limpando a senha). Sucesso → `concluirLogin(access_token)`; se `codigos_recuperacao_restantes` não for nulo e for `<= 3`, `toast('Poucos códigos de recuperação restantes (N). Gere novos em Meu perfil.')`; navega para `/`.
    - `cadastro`: título "Cadastre o autenticador", texto "Gestores precisam da verificação em duas etapas. Instale um aplicativo autenticador (Google Authenticator, Microsoft Authenticator ou similar), leia o QR code e digite o código gerado."; `CadastroAutenticador` com `iniciar={() => iniciarCadastroMfa(mfaToken)}` e `confirmar` que chama `confirmarCadastroMfa`, guarda o `access_token` e devolve `codigos_recuperacao`.
    - `codigos`: título "Códigos de recuperação"; `CodigosRecuperacao` com `rotuloContinuar="Entrar no sistema"` → `concluirLogin(accessToken)` e navega para `/` (o `/auth/me` só é chamado aqui).

- [ ] **Step 1: Testes** (`LoginPage.test.tsx`, MSW; renderizar com `<MemoryRouter initialEntries={['/login']}><AuthProvider><Routes><Route path="/login" element={<LoginPage />} /><Route path="/" element={<p>Início</p>} /><Route path="/esqueci-senha" element={<p>Tela esqueci</p>} /></Routes></AuthProvider></MemoryRouter>` e `<ToastContainer />`):
  - os dois testes atuais continuam passando.
  - `etapa mfa pede o código e entra` (login → `{ etapa: 'mfa', mfa_token: 'mfa-1', expira_em_segundos: 300 }`; `/auth/mfa/verificar` recebe `{ mfa_token: 'mfa-1', codigo: '123456' }` e responde `ok`; aparece "Início").
  - `código errado mostra a mensagem do servidor` (400 `{ detail: 'Código inválido.' }` → texto visível, continua na etapa).
  - `voltar retorna ao formulário de senha`.
  - `poucos códigos de recuperação avisa` (`codigos_recuperacao_restantes: 2` → toast com `'Poucos códigos de recuperação restantes (2)'`).
  - `cadastro obrigatório mostra QR, confirma e exibe os códigos antes de entrar` (login → `cadastro_2fa`; iniciar → `CadastroTotp`; confirmar → `{ etapa: 'ok', access_token: 'tok-9', codigos_recuperacao: [10 códigos] }`; os códigos aparecem; o handler de `/auth/me` ainda não foi chamado; marcar o checkbox e "Entrar no sistema" → aparece "Início").
  - `muitas tentativas` (login → 429 → `'Muitas tentativas. Aguarde um minuto e tente de novo.'`).
  - `link esqueci minha senha` (clicar → aparece "Tela esqueci").
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/LoginPage.test.tsx`.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run src/pages/LoginPage.test.tsx src/components/auth` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(login): login em duas etapas e cadastro do autenticador para Gestor`.

### Task 12: Telas "Esqueci minha senha" e "Redefinir senha"

**Files:**
- Create: `frontend/src/pages/EsqueciSenhaPage.tsx`, `frontend/src/pages/RedefinirSenhaPage.tsx`, `frontend/src/pages/EsqueciSenhaPage.test.tsx`, `frontend/src/pages/RedefinirSenhaPage.test.tsx`
- Modify: `frontend/src/App.tsx` (acrescentar 2 imports e, logo depois da rota `/login`, `<Route path="/esqueci-senha" element={<EsqueciSenhaPage />} />` e `<Route path="/redefinir-senha" element={<RedefinirSenhaPage />} />` — fora do `AppLayout`, sem login)

**Interfaces:**
- Consumes: `esqueciSenha`, `redefinirSenha` (Task 9); `AuthShell` (Task 11); `mensagemErroAuth`; `Button loading`.
- Produces:
  - `EsqueciSenhaPage`: título "Esqueci minha senha", texto "Informe seu usuário ou e-mail. Se houver um e-mail cadastrado, enviaremos um link para criar uma nova senha.", campo "Usuário ou e-mail", botão "Enviar link"; sucesso → esconde o formulário e mostra o `detail` do servidor em `role="status"`; link "Voltar para o login".
  - `RedefinirSenhaPage`: lê `token` de `?token=`; sem token → `'Link inválido ou expirado. Peça um novo.'` e link "Pedir novo link" (`/esqueci-senha`); campos "Nova senha" e "Confirme a nova senha", botão "Salvar nova senha"; senhas diferentes → `'As senhas não conferem.'` sem chamar a API; nova senha vazia → botão desabilitado; sucesso → `detail` do servidor e link "Ir para o login"; 400 → o `detail` e o link "Pedir novo link".

- [ ] **Step 1: Testes** (MSW):
  - `EsqueciSenhaPage.test.tsx`: envia `{ identificador: 'joao.silva' }` e mostra `'Se o usuário existir e tiver e-mail cadastrado, enviaremos um link para redefinir a senha.'`; 429 → mensagem de muitas tentativas; `rota pública` — `renderWithClient(<MemoryRouter initialEntries={['/esqueci-senha']}><App /></MemoryRouter>)` mostra o título "Esqueci minha senha" sem chamar `/api/auth/refresh` (sem `AuthProvider`, sem login).
  - `RedefinirSenhaPage.test.tsx`: `?token=abc` + senhas iguais → `POST /api/auth/reset` recebe `{ token: 'abc', nova_senha: 'Nova#123' }` → `'Senha redefinida. Entre com a nova senha.'` e link "Ir para o login"; senhas diferentes → `'As senhas não conferem.'` e nenhuma chamada (contador do handler em 0); sem token → mensagem de link inválido e link para `/esqueci-senha`; 400 `{ detail: 'Link inválido ou expirado. Peça um novo.' }` → mensagem e link "Pedir novo link".
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/EsqueciSenhaPage.test.tsx src/pages/RedefinirSenhaPage.test.tsx`.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** os dois arquivos e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(login): telas de esqueci minha senha e redefinição por link`.

### Task 13: Página "Meu perfil"

**Files:**
- Create: `frontend/src/pages/PerfilPage.tsx`, `frontend/src/pages/PerfilPage.test.tsx`
- Modify: `frontend/src/App.tsx` (acrescentar import e, dentro do `AppLayout`, `<Route path="perfil" element={<PerfilPage />} />` sem `RequireRole`), `frontend/src/components/layout/navigation.ts` (acrescentar a linha `import { UserCog } from 'lucide-react'` separada e, no fim de `navGroups`, o grupo `{ title: 'Conta', items: [{ to: '/perfil', label: 'Meu perfil', icon: UserCog }] }`)

**Interfaces:**
- Consumes: `api/perfil.ts` (Task 9); `CadastroAutenticador`, `CodigosRecuperacao` (Task 10); `StatusPill`, `Button loading`, `Dialog`, `PageHeader`; `getVisibleNavItems`.
- Produces: `PerfilPage` (query `['perfil']`):
  - Bloco "Dados": "Usuário", "Função", "E-mail" (valor ou "Não cadastrado") e a nota "Para alterar o e-mail, fale com um Gestor."
  - Bloco "Verificação em duas etapas": `StatusPill` "Ativo" (ativo) ou `status="Inativo" tone="neutral"`; com `mfa_obrigatorio`, o texto "Obrigatória para Gestor."; "Códigos de recuperação restantes: N" quando ativo; sem `mfa_disponivel`, o aviso "A verificação em duas etapas está indisponível neste servidor. Avise o administrador." e nenhum botão de ativar/gerar.
  - Ações: "Ativar verificação em duas etapas" (inativo) → `CadastroAutenticador` (`iniciarCadastro2fa`, `confirmar` → `confirmarCadastro2fa(c).codigos_recuperacao`) → `CodigosRecuperacao` com `rotuloContinuar="Concluir"` → invalida `['perfil']` e `toast('Verificação em duas etapas ativada.')`; "Gerar novos códigos de recuperação" (ativo) e "Desativar" (ativo e não obrigatório) abrem o diálogo "Confirme sua senha" (campo "Senha", botões "Cancelar"/"Confirmar" com `loading`); erro → `mensagemErroAuth`. Desativar com sucesso → `toast('Verificação em duas etapas desativada.')` e invalida `['perfil']`.

- [ ] **Step 1: Testes** (`PerfilPage.test.tsx`, MSW, `renderWithClient` + `MemoryRouter`):
  - mostra usuário, função e e-mail; e-mail `null` → "Não cadastrado"; a nota de alteração aparece.
  - Técnico inativo ativa: "Ativar verificação em duas etapas" → QR (iniciar) → código → códigos → "Concluir" → `GET /api/perfil` de novo devolve `totp_ativo: true` e aparece "Ativo".
  - Técnico ativo desativa: diálogo envia `{ senha: 'minha-senha' }` a `/api/perfil/2fa/desativar`; 400 `{ detail: 'Senha incorreta.' }` mostra a mensagem.
  - Gestor ativo: sem botão "Desativar" e com "Obrigatória para Gestor.".
  - `mfa_disponivel: false` → aviso de indisponível e sem "Ativar verificação em duas etapas".
  - "Gerar novos códigos de recuperação" com senha → mostra os 10 códigos devolvidos.
  - `getVisibleNavItems('Jovem Aprendiz')` contém `to: '/perfil'`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/PerfilPage.test.tsx`.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run src/pages/PerfilPage.test.tsx src/components/layout` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(perfil): página Meu perfil com e-mail e verificação em duas etapas`.

### Task 14: Tela de Usuários (e-mail, 2FA, paginação)

**Files:**
- Create: `frontend/src/pages/UsersPage.test.tsx`
- Modify: `frontend/src/api/users.ts`, `frontend/src/pages/UsersPage.tsx`

**Interfaces:**
- Consumes: `GET/POST /api/users`, `PUT /api/users/{id}`, `DELETE /api/users/{id}/2fa`, `PUT /api/users/{id}/password`, `DELETE /api/users/{id}` (Tasks 1, 8); `Pagination`, `usePageSize`, `StatusPill`, `Button loading`, `Dialog`, `AlertDialog`, `Select`, `useAuth().user`.
- Produces:
  - `api/users.ts`: `interface User { id: number; username: string; role: string; email: string | null; totp_ativo: boolean }`; `createUser(data: { username: string; password: string; role: string; email?: string | null })`; `updateUserEmail(id: number, email: string | null): Promise<User>` (`PUT /users/{id}` `{ email }`); `zerar2fa(id: number): Promise<{ detail: string }>` (`DELETE /users/{id}/2fa`); `listUsers`, `removeUser`, `updatePassword` inalterados.
  - `UsersPage`: o formulário de cadastro ganha "E-mail" (opcional, `type="email"`, placeholder `nome@revalle.com.br`; vazio → `null`) e o botão "Criar Usuário" usa `loading`. Os três cards por função dão lugar a uma tabela (dentro de `overflow-x-auto`) com colunas "Usuário", "E-mail" ("—" quando vazio), "Função", "2FA" e "Ações"; 2FA: `totp_ativo` → `<StatusPill status="Ativo" />`, Gestor sem 2FA → `<StatusPill status="Pendente" />`, demais → `<StatusPill status="Inativo" tone="neutral" />`. Filtros: busca "Buscar por usuário ou e-mail" e seletor "Todas as funções"/papéis; paginação no cliente com `usePageSize('usuarios')` e `Pagination` (volta à página 1 ao mudar filtro/tamanho; corrige a página quando passa da última). Ações por linha (botões com `aria-label`): "Editar e-mail" (diálogo `Editar e-mail — <usuário>`, campo "E-mail", "Salvar" com `loading`, vazio envia `null`, sucesso `toast('E-mail atualizado.')`), "Alterar senha" (o banner atual, com `loading`), "Zerar 2FA" (só quando `totp_ativo` e não é o usuário logado; `AlertDialog` `Zerar a verificação em duas etapas de "<usuário>"?` com a descrição "As sessões abertas do usuário serão encerradas. Um Gestor terá de cadastrar um novo autenticador no próximo login; os demais voltam a entrar só com a senha." e ação "Zerar 2FA"; sucesso → `toast(detail)`), "Remover usuário" (como hoje). Erros → `getErrorMessage`.

- [ ] **Step 1: Testes** (`UsersPage.test.tsx`, MSW, `vi.mock('@/contexts/AuthContext')` como em `ColaboradoresPage.test.tsx` com `user: { id: 1, username: 'gestor', role: 'Gestor' }`; seletor Radix como em `RemovePage.test.tsx`):
  - tabela com e-mail e as pílulas "Ativo", "Pendente" (Gestor sem 2FA) e "Inativo".
  - 30 usuários → `'Mostrando 1–25 de 30'`; "Próxima" mostra o 26º.
  - busca por e-mail filtra e volta à página 1.
  - criar envia `{ username, password, role: 'Técnico', email: 'novo@revalle.com.br' }`; sem e-mail envia `email: null`.
  - editar e-mail envia `PUT /api/users/2` `{ email: 'tec@revalle.com.br' }`; 400 `{ detail: 'Este e-mail já está em uso por outro usuário.' }` mostra a mensagem; campo vazio envia `{ email: null }`.
  - "Zerar 2FA" aparece só para outro usuário com `totp_ativo` (não para o id 1 nem para quem está inativo); confirmar chama `DELETE /api/users/3/2fa`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/UsersPage.test.tsx`.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` (suíte inteira) e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(usuarios): tela de usuários com e-mail, 2FA e paginação`.

### Task 15: Verificação final da fase

**Files:** nenhum novo (só correções que a verificação apontar).

- [ ] **Step 1: Backend completo:** `RBT -q -m "not mudanca_esperada"` → verde (inclui `test_migrations.py`, que aplica as 0200–0202 do zero no banco de teste).
- [ ] **Step 2: Frontend completo:** em `frontend/`, `npx vitest run`, `npx tsc --noEmit` e `npm run build` → verdes.
- [ ] **Step 3: Segredos fora de log e auditoria:** `grep -rnE "logger\.(info|warning|error|exception)\(.*(token|segredo|codigo|senha|email)" backend/app/routers/auth.py backend/app/routers/auth_senha.py backend/app/routers/perfil.py backend/app/services/segundo_fator.py backend/app/services/mfa.py` → nada que registre o valor (só mensagens fixas).
- [ ] **Step 4: Conferir o spec da F2 item a item** (marcar cada um com o teste que prova):
  - [ ] 1.1 / F2.2 — `POST /auth/forgot` genérico (`test_forgot_usuario_inexistente_resposta_igual`) com rate limit (`test_forgot_tem_rate_limit`); link `APP_PUBLIC_URL/redefinir-senha?token=` de 30 min (`test_forgot_envia_email_com_link`), uso único (`test_reset_token_reutilizado_recusado`), expirado (`test_reset_token_expirado_recusado`), hash (`test_forgot_envia_email_com_link`); mesmas regras de força (`test_reset_senha_em_branco_nao_gasta_o_token`); revoga sessões (`test_reset_revoga_sessoes`); auditoria (`test_reset_grava_auditoria_sem_segredos`); telas (Task 12).
  - [ ] F2.1 — e-mail na tela de Usuários (Tasks 1 e 14) e no perfil (Tasks 7 e 13).
  - [ ] 1.2 / F2.3 — QR + confirmação (`test_cadastro_completo_emite_sessao_e_codigos`, Task 10); 10 códigos exibidos uma vez e guardados como hash (idem + `CodigosRecuperacao`); segredo cifrado (idem); login em duas etapas com token de 5 min `mfa_pendente` (`test_login_com_2fa_pede_codigo`); verificar com TOTP, recuperação e rate limit (`test_verificar_com_totp_emite_sessao`, `test_verificar_com_codigo_de_recuperacao`, `test_verificar_tem_rate_limit`); código TOTP errado (`test_verificar_codigo_errado`); Gestor obrigado a cadastrar, garantido no servidor (`test_gestor_sem_2fa_recebe_etapa_cadastro`, `test_gestor_sem_2fa_com_sessao_antiga_e_bloqueado`); Gestor zera o 2FA de outro com auditoria (`test_gestor_zera_2fa_de_outro`); `python -m scripts.reset_2fa <usuario>` (`test_script_reset_2fa_zera_e_audita`); demais papéis no perfil (`test_tecnico_ativa_2fa_no_perfil`).
  - [ ] Seção 8 — tokens e códigos só como hash, segredo TOTP cifrado, nada no `.env.example` além de comentário.
  - [ ] Review Focus — os 5 testes citados existem e passam.
- [ ] **Step 5: Relatório da fase** (texto para o coordenador, não arquivo): dependência nova `qrcode@1.5.4`/`@types/qrcode@1.5.5` e a justificativa; as Decisões 1–13; arquivos compartilhados tocados (`main.py`, `dependencies.py` — incluindo a linha nova em `get_current_user` —, `.env.example`, `tests/conftest.py` (1 linha em `pytest_configure`), `App.tsx`, `navigation.ts`); sugestão para a F1 (dona do `TopBar`): o nome do usuário pode virar link para `/perfil`.
- [ ] **Step 6: Commit** (só se houve correção) `fix(login): ajustes da verificação final da Fase 2`.
