# Login por e-mail: plano de implementação

> **Para quem executar:** use a skill superpowers:subagent-driven-development (recomendada) ou superpowers:executing-plans para executar este plano tarefa por tarefa. Os passos usam caixas de seleção (`- [ ]`) para acompanhar o progresso.

**Objetivo:** todas as contas passam a entrar com e-mail `@revalle.com.br` e senha. As contas antigas cadastram e confirmam o e-mail no próximo login, sem perder o `id` nem o histórico.

**Arquitetura:** o login em etapas já existente ganha a etapa `cadastro_email`, que vem antes do 2FA. Ela usa um token de etapa de propósito `"email"` e um código de 6 dígitos enviado por e-mail, guardado só como hash numa tabela nova. O `username` continua existindo por dentro (o `history.usuario` usa ele), mas não aparece mais no login.

**Tecnologias:** FastAPI, PyMySQL com SQL cru (`get_cursor`), migrações SQL aditivas, React 18, TanStack Query, Vitest e MSW.

**Spec:** `docs/superpowers/specs/2026-11-17-login-por-email-design.md`.

## Restrições globais
- Domínio aceito: configuração `LOGIN_EMAIL_DOMINIO: str = "revalle.com.br"`. A comparação é exata com a parte depois do `@`, sem maiúsculas e sem espaços nas pontas.
- Código: 6 dígitos com zeros à esquerda, válido por 15 minutos. Reenvio só depois de 60 segundos. Com 5 códigos errados, o código e o token da etapa deixam de valer.
- Mensagens fixas:
  - 401 `"Esta conta agora entra com o e-mail. Use o seu e-mail @revalle.com.br."`
  - 409 `"Este e-mail já está em uso por outra conta."`
  - 422 `"Use o seu e-mail @revalle.com.br."`
  - 400 `"Código inválido ou expirado."`
  - 400 `"Muitas tentativas. Entre de novo para receber outro código."`
  - 429 `"Aguarde 1 minuto para pedir outro código."`
- Login com senha errada ou conta inexistente: mantém a mensagem genérica atual e o custo do bcrypt com `DUMMY_PASSWORD_HASH`.
- Erros das etapas são 400, nunca 401: o interceptor do frontend trata 401 como sessão expirada (ver `_decodificar_mfa`).
- Migração nova `backend/migrations/0900_usuarios_email_login.sql` com `.down.sql`, aditiva e sem `CHARACTER SET`/`COLLATE` (Ruling 39). Acrescentar a faixa 0900–0999 ao `backend/migrations/README.md`.
- Auditoria com `audit.record`. O e-mail aparece mascarado (`j***@revalle.com.br`), nunca o código.
- Português em toda mensagem visível. Componentes da F0 (`Button loading`, `StatusPill`).
- Testes do backend: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh C:/Users/enzo.jz/dev/patrimonio/login-email email <args>`. Frontend: em `frontend/`, `npx vitest run <arquivo>`, `npx tsc --noEmit` e `npm run build`.
- Commits no formato `tipo(escopo): descrição`, linha em branco, `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Pontos de atenção na revisão
1. E-mail digitado com maiúsculas ou espaços (`" Joao.Silva@Revalle.COM.br "`) → é normalizado e aceito. O login pelo e-mail também não diferencia maiúsculas (Task 1 e Task 2).
2. Domínios parecidos (`a@revalle.com.br.evil.com`, `a@evilrevalle.com.br`, `a@sub.revalle.com.br`) → recusados com 422 (Task 1).
3. Conta com e-mail confirmado tentando entrar pelo usuário antigo: com a senha certa, recebe o 401 específico; com a senha errada, a mensagem genérica (Task 2).
4. Token da etapa de e-mail usado nas rotas do 2FA, e o contrário → 400 de etapa expirada. Gestor sem 2FA que confirma o e-mail cai em `cadastro_2fa` e não recebe sessão (Task 2).
5. O mesmo código confirmado duas vezes, código vencido e código pedido de novo antes de 60 segundos → recusados. Depois de um novo pedido, o código anterior deixa de valer (Task 1).

---

### Task 1: dados, validação e serviço de confirmação

**Arquivos:**
- Criar: `backend/migrations/0900_usuarios_email_login.sql` e `.down.sql`.
  - Coluna `usuarios.email_confirmado TINYINT(1) NOT NULL DEFAULT 0`, com o mesmo padrão idempotente `SET @faltando … PREPARE` da `0200_usuarios_email_2fa.sql`.
  - Tabela `email_confirmacoes`: `id` INT auto, `usuario_id` INT NOT NULL com FK para `usuarios(id)` ON DELETE CASCADE, `email` VARCHAR(200) NOT NULL, `codigo_hash` CHAR(64) NOT NULL, `expira_em` DATETIME NOT NULL, `tentativas` INT NOT NULL DEFAULT 0, `usado` TINYINT(1) NOT NULL DEFAULT 0, `criado_em` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, `INDEX (usuario_id)`.
- Modificar: `backend/app/core/config.py`: `LOGIN_EMAIL_DOMINIO: str = "revalle.com.br"`. Também em `backend/.env.example` e na tabela de variáveis do README.
- Modificar: `backend/app/schemas/validators.py`: `validar_email_corporativo(email: str) -> str`.
- Modificar: `backend/app/db/user_manager_db.py`.
- Criar: `backend/app/db/email_confirmacao_db.py`.
- Criar: `backend/app/services/confirmacao_email.py`.
- Testes: `backend/tests/test_confirmacao_email.py`, além de casos novos em `backend/tests/test_migrations.py` se ele conferir a lista de migrações.

**Interfaces:**
- Produz `validar_email_corporativo(email) -> str`. Devolve o e-mail sem espaços e em minúsculas. Lança `ValueError("Use o seu e-mail @revalle.com.br.")` com formato ruim ou domínio diferente de `settings.LOGIN_EMAIL_DOMINIO`.
- Produz em `UserDBManager`:
  - `get_user_para_login(identificador: str) -> dict | None`. Com `@` no texto, busca por `LOWER(email)`; senão, por `username`. Devolve `id, username, password, role, email, email_confirmado, totp_ativo, sessao_versao`.
  - `confirmar_email(user_id: int, email: str) -> None`. Grava o `email` e `email_confirmado=1`. Lança `EmailEmUso` (exceção nova no mesmo módulo) quando o e-mail é duplicado.
  - `email_em_uso_por_outro(email: str, user_id: int) -> bool`.
  - `gerar_username(email: str) -> str`. Usa a parte antes do `@`; se já existir, `base2`, `base3` e assim por diante.
  - `update_email` passa a gravar `email_confirmado=0` junto com o e-mail.
  - `_CAMPOS_PUBLICOS` e `get_all_users` incluem `email_confirmado`.
- Produz `EmailConfirmacaoDB` em `app/db/email_confirmacao_db.py`, com:
  - `criar(usuario_id, email, codigo_hash, expira_em) -> int`, que antes marca como usados os pendentes do usuário;
  - `pendente(usuario_id) -> dict | None`, o mais recente com `usado=0`;
  - `registrar_falha(id) -> int`, que devolve o total de tentativas;
  - `marcar_usado(id) -> None`;
  - `segundos_desde_ultimo(usuario_id) -> int | None`.
- Produz em `app/services/confirmacao_email.py`:
  - exceções `ConfirmacaoErro(Exception)` com `status_code`, e `CodigoInvalido` (400), `CodigoEsgotado` (400), `AguardeReenvio` (429), `EmailEmUsoErro` (409), `DominioInvalido` (422), cada uma com a mensagem fixa;
  - `iniciar(usuario: dict, email: str, ctx: AuditContext) -> str`: valida, confere o cooldown e se o e-mail está em uso, gera o código, grava o hash e envia o e-mail. Devolve o e-mail normalizado.
  - `confirmar(usuario: dict, codigo: str, ctx: AuditContext) -> str`: devolve o e-mail confirmado.
  - O hash é `hmac.new(settings.JWT_SECRET.encode(), f"{usuario_id}:{email}:{codigo}".encode(), sha256).hexdigest()`, comparado com `hmac.compare_digest`. O código vem de `secrets.randbelow(10**6)` formatado com `:06d`.
  - O e-mail é montado com `render_email("Confirme seu e-mail", …)` e enviado com `send_mail`. O assunto é `"Código de confirmação — Patrimônio Revalle"` e o corpo mostra o código grande e "Vale por 15 minutos."

- [ ] **Passo 1: escrever os testes que falham**, em `test_confirmacao_email.py`:
  - `test_valida_dominio_normaliza`: `validar_email_corporativo(" Joao.Silva@Revalle.COM.br ") == "joao.silva@revalle.com.br"`.
  - `test_recusa_dominios_parecidos`: cada um de `a@revalle.com.br.evil.com`, `a@evilrevalle.com.br`, `a@sub.revalle.com.br`, `a@gmail.com` e `semarroba` lança `ValueError`.
  - `test_iniciar_envia_codigo_e_guarda_so_hash`: `MAIL_BACKEND=memory`. Sai 1 mensagem para o e-mail, com um código de 6 dígitos no corpo. O `codigo_hash` gravado é diferente do código e tem 64 caracteres.
  - `test_reenvio_antes_de_60s_recusado`: um segundo `iniciar` imediato lança `AguardeReenvio`. Com o `criado_em` envelhecido em 61 segundos por SQL, passa, e o código antigo deixa de valer.
  - `test_email_de_outra_conta_recusado`: lança `EmailEmUsoErro` e nenhum e-mail é enviado.
  - `test_confirmar_grava_email_confirmado`: depois de confirmar, `get_user_by_id` traz `email` e `email_confirmado == 1`, e existe uma linha em `audit_log` com acao `confirmar_email` e e-mail mascarado.
  - `test_codigo_errado_5_vezes_esgota`: as 4 primeiras falhas lançam `CodigoInvalido` e a 5ª lança `CodigoEsgotado`. Depois disso, nem o código certo funciona.
  - `test_codigo_vencido_e_reuso_recusados`: com `expira_em` no passado, recusado; depois de um sucesso, o mesmo código é recusado.
  - `test_update_email_zera_confirmacao`.
  - `test_gerar_username_unico`: com `joao.silva` já existente, devolve `joao.silva2`.
- [ ] **Passo 2:** rodar `… run-backend-tests.sh … email tests/test_confirmacao_email.py -q` e confirmar que FALHA (import ou atributo inexistente).
- [ ] **Passo 3:** implementar a migração, a configuração, o validador, os métodos de banco e o serviço com as assinaturas acima.
- [ ] **Passo 4:** rodar o mesmo comando e confirmar que PASSA. Rodar também `tests/test_migrations.py tests/test_users.py tests/test_redefinir_senha.py -q`, que precisam continuar verdes.
- [ ] **Passo 5: commit** `feat(login): confirmação de e-mail corporativo por código (dados e serviço)`.

### Task 2: login por e-mail e etapa `cadastro_email`

**Arquivos:**
- Modificar `backend/app/core/security.py`: `create_mfa_token` passa a aceitar o propósito `"email"`, com 15 minutos (`MFA_CADASTRO_MINUTOS`).
- Modificar `backend/app/schemas/auth.py`:
  - `LoginResponse.etapa` passa a ser `Literal["ok", "mfa", "cadastro_2fa", "cadastro_email"]`, com o campo novo `email_sugerido: Optional[str] = None`;
  - `EmailIniciarRequest(mfa_token: str, email: str)`;
  - `UserMe` ganha `email: Optional[str]`.
  - `MfaConfirmarRequest` já serve para o confirmar.
- Modificar `backend/app/routers/auth.py`, `login`, nesta ordem:
  1. `get_user_para_login(body.username.strip())`;
  2. bcrypt como hoje;
  3. se não achou a conta ou a senha está errada → 401 genérico;
  4. se o identificador não tem `@` e a conta tem `email_confirmado` → 401 com a mensagem específica;
  5. se `not email_confirmado` → `etapa="cadastro_email"`, `mfa_token=create_mfa_token(dados_do_token(user), "email")`, `email_sugerido=user.email`;
  6. senão, o fluxo de 2FA e de sessão de hoje.
- Extrair um helper `_proxima_etapa(response, usuario) -> LoginResponse` que aplica a ordem `mfa` / `cadastro_2fa` / `ok`. O `login` usa ele depois do item 5 e o confirmar de e-mail usa ele também.
- Novas rotas, com `@limiter.limit(settings.LOGIN_RATE_LIMIT)`:
  - `POST /email/iniciar` → `{"detail": "Enviamos um código para <email>."}`. Não gasta o token.
  - `POST /email/confirmar` → `_decodificar_mfa(..., "email", ...)`. Chama `confirmar`. Com `CodigoEsgotado`, chama `_revogar_mfa` e devolve 400. Com sucesso, chama `_revogar_mfa` e devolve `_proxima_etapa`, relendo o usuário do banco.
- `/me` devolve `email`.
- Modificar `backend/app/routers/auth_senha.py`, `_processar_pedido`: só envia o link quando `usuario.get("email_confirmado")`. A resposta continua a mesma neutra de hoje.
- Testes:
  - novo `backend/tests/test_login_email.py`;
  - ajustar os testes existentes que fazem login (`fixtures_auth.py` e `conftest.py`, nos clientes autenticados). Os usuários de teste são criados com `email_confirmado=1` e e-mail `<username>@revalle.com.br`, para os testes antigos continuarem iguais.

**Interfaces:**
- Consome a Task 1.
- Produz o contrato HTTP usado pela Task 4: `POST /api/auth/login {username, password}` (o campo `username` aceita e-mail), `POST /api/auth/email/iniciar {mfa_token, email}` e `POST /api/auth/email/confirmar {mfa_token, codigo}` → `LoginResponse`.

- [ ] **Passo 1: escrever os testes que falham**, em `test_login_email.py`:
  - `test_conta_antiga_por_usuario_vai_para_cadastro_email`: a resposta tem `etapa == "cadastro_email"`, nenhum `access_token` e nenhum cookie de refresh.
  - `test_fluxo_tecnico_confirma_e_entra`: iniciar, pegar o código do e-mail em memória, confirmar → `etapa == "ok"` com `access_token`. Depois disso, o login com e-mail `JOAO@REVALLE.COM.BR` e a senha → `ok`.
  - `test_gestor_confirma_email_e_cai_em_cadastro_2fa`: com `TOTP_ENCRYPTION_KEY` definida, confirmar → `etapa == "cadastro_2fa"` sem `access_token`.
  - `test_usuario_antigo_apos_confirmar_recebe_401_especifico`: com a senha certa, a mensagem específica; com a senha errada, a genérica.
  - `test_token_de_email_nao_serve_no_2fa_e_vice_versa`: os dois casos → 400.
  - `test_confirmar_esgotado_revoga_token`: depois de 5 erros, o próprio `iniciar` com o mesmo token → 400 de etapa expirada.
  - `test_me_traz_email`.
  - `test_forgot_ignora_email_nao_confirmado`: nenhum e-mail enviado; resposta neutra.
- [ ] **Passo 2:** rodar `tests/test_login_email.py` e confirmar que FALHA.
- [ ] **Passo 3:** implementar as mudanças de schema, segurança e rotas acima e ajustar as fixtures.
- [ ] **Passo 4:** rodar `tests/test_login_email.py` e depois a suíte completa (`-q -m "not mudanca_esperada"`, em segundo plano). Tudo verde.
- [ ] **Passo 5: commit** `feat(login): entrar com e-mail; contas antigas confirmam o e-mail no próximo acesso`.

### Task 3: Gestor cria contas por convite e envia links de senha

> Revisado em 2026-11-17 (mudança do responsável): ninguém sabe a senha de outra pessoa. A conta nasce sem senha utilizável e recebe um convite por e-mail; o Gestor não define mais senhas, só envia links.

**Arquivos:**
- Modificar `backend/migrations/0900_usuarios_email_login.sql` e `.down.sql`: coluna `usuarios.senha_definida TINYINT(1) NOT NULL DEFAULT 1` (as contas atuais ficam com 1), no mesmo padrão idempotente.
- Modificar `backend/app/db/user_manager_db.py`:
  - `add_user(..., senha_definida: bool = True)`; `_CAMPOS_PUBLICOS`, `_CAMPOS_LOGIN` e `get_all_users` incluem `senha_definida`;
  - `definir_senha_pelo_link(user_id, nova_senha)`: grava a senha com `senha_definida=1` e `email_confirmado=1`;
  - `update_email` com e-mail diferente invalida os links de senha abertos da conta.
- Modificar `backend/app/routers/auth.py`, `login`: conta com `senha_definida=0` → 401 genérico (depois do bcrypt).
- Modificar `backend/app/routers/auth_senha.py`:
  - `CONVITE_VALIDADE_HORAS = 72`; `emitir_link(usuario, convite: bool) -> str` invalida os links abertos da conta, grava o novo (72 h no convite, 30 min no link normal) e devolve o token;
  - `enviar_convite(email, token)` e `enviar_link_senha(email, username, token)`, síncronos, que levantam `MailError`;
  - convite: assunto "Seu acesso ao Patrimônio Revalle", corpo "Você recebeu acesso ao sistema de Patrimônio da Revalle. Crie sua senha pelo botão abaixo. O link vale por 72 horas.", botão "Criar minha senha", link `APP_PUBLIC_URL/redefinir-senha?token=…&convite=1`;
  - `/reset` passa a usar `definir_senha_pelo_link`.
- Modificar `backend/app/schemas/users.py`:
  - `UserCreate`: só `email` (obrigatório, `validar_email_corporativo`) e `role`;
  - `UserUpdate.email`, quando preenchido, também passa por `validar_email_corporativo`;
  - `UserResponse` ganha `email_confirmado: bool` e `senha_definida: bool`.
- Modificar `backend/app/routers/users.py`:
  - `create_user`: `username = gerar_username(email)`, senha = bcrypt de `secrets.token_urlsafe(32)`, `senha_definida=0`, `email_confirmado=0`; envia o convite; auditoria `criar` e `enviar_convite`. Falha no envio → 200 com `"Conta criada, mas não foi possível enviar o e-mail. Use “Reenviar convite”."` e `convite_enviado: false`;
  - remover `PUT /api/users/{id}/password`;
  - novo `POST /api/users/{id}/enviar-link-senha` (Gestor): convite se `senha_definida=0`, senão link normal, para o e-mail atual. Sem e-mail → 400 `"Cadastre o e-mail desta conta primeiro."`. Auditoria `enviar_convite` / `enviar_link_senha`. Falha no envio → 503.
- Testes em `backend/tests/test_users.py` (e ajustes em `test_usuarios_email.py`, `test_audit_fluxos.py`, `test_rbac.py`):
  - criar sem senha → 1 convite enviado, conta com `senha_definida=0` e `email_confirmado=0`, `username` gerado;
  - criar com e-mail de fora → 422;
  - login antes de concluir o convite → 401 genérico;
  - concluir o convite → `email_confirmado=1`; login pelo e-mail → `ok` (Técnico) ou `cadastro_2fa` (Gestor);
  - convite com mais de 72 h → recusado;
  - reenviar invalida o link anterior;
  - endpoint removido → 404/405;
  - `enviar-link-senha` por quem não é Gestor → 403; conta sem e-mail → 400;
  - editar o e-mail zera a confirmação; a lista traz `email_confirmado` e `senha_definida`;
  - falha no envio do convite mantém a conta criada.

**Interfaces:**
- Consome `gerar_username` e `validar_email_corporativo` da Task 1.
- Produz para a Task 4: `GET /api/users` → itens com `{id, username, role, email, email_confirmado, senha_definida, totp_ativo}`; `POST /api/users {email, role}` → `{detail, convite_enviado}`; `POST /api/users/{id}/enviar-link-senha` → `{detail}`.

- [ ] **Passo 1:** escrever os testes acima.
- [ ] **Passo 2:** rodá-los e confirmar que FALHAM.
- [ ] **Passo 3:** implementar.
- [ ] **Passo 4:** rodar `tests/test_users.py` e a suíte completa.
- [ ] **Passo 5: commit** `feat(usuarios): contas criadas por convite; Gestor envia link em vez de definir senha`.

### Task 4: frontend

**Arquivos:**
- Modificar `frontend/src/api/auth.ts`:
  - o tipo `EtapaLogin` inclui `'cadastro_email'`;
  - `LoginResposta.email_sugerido?: string | null`;
  - `iniciarCadastroEmail(mfaToken: string, email: string)` e `confirmarCadastroEmail(mfaToken: string, codigo: string)`, este devolvendo a mesma `LoginResposta`.
- Modificar `frontend/src/pages/LoginPage.tsx`:
  - o campo passa a ter o rótulo "E-mail", `id`/`name` `"email"` e `autoComplete="username"`, com a dica abaixo dele: "Primeiro acesso depois da mudança? Entre uma vez com o seu usuário antigo.";
  - nova etapa `{ tipo: 'email', mfaToken, emailSugerido }` com o componente `EmailEtapa`, que tem:
    - o campo de e-mail, preenchido com o `emailSugerido`;
    - o botão "Enviar código" (`loading`);
    - depois do envio, o campo "Código" com 6 dígitos (`inputMode="numeric"`), o botão "Confirmar" e o botão "Reenviar código", desabilitado nos 60 segundos seguintes ao envio;
    - os erros da API exibidos em `MensagemErro`.
  - Depois de confirmar, a resposta segue o mesmo tratamento do login: `ok` entra; `cadastro_2fa` vai para a etapa de cadastro; `mfa` vai para a etapa de verificação.
  - Com 400 de etapa expirada ou "Muitas tentativas", volta para a senha com o aviso.
  - O título da etapa no `CABECALHOS`: "Confirme seu e-mail".
- Modificar `frontend/src/pages/EsqueciSenhaPage.tsx`: o rótulo do campo passa a ser "E-mail".
- Modificar `frontend/src/pages/RedefinirSenhaPage.tsx`: com `?convite=1` no link, o título é "Criar sua senha".
- Modificar `frontend/src/api/users.ts` e `frontend/src/pages/UsersPage.tsx`:
  - o formulário de criação tem só E-mail (obrigatório, com a dica "@revalle.com.br") e Papel;
  - a lista mostra o e-mail como coluna principal e um `StatusPill` "Convite pendente", "Aguardando confirmação" ou "Confirmado";
  - o `username` aparece numa coluna secundária "Usuário (histórico)";
  - a troca de senha pelo Gestor sai; a ação da linha é "Reenviar convite" (convite pendente) ou "Enviar link para criar nova senha".
- Modificar `frontend/src/components/layout/TopBar.tsx`: mostra `user.email ?? user.username`. O tipo do usuário em `AuthContext` ganha `email?: string | null`.
- Testes, com handlers do MSW em `frontend/src/test/handlers.ts`:
  - `LoginPage.test.tsx`:
    - "conta antiga: confirma o e-mail e entra";
    - "Gestor: confirma o e-mail e segue para o cadastro do 2FA";
    - "reenviar fica desabilitado logo após o envio";
    - "erro 409 de e-mail em uso aparece na etapa".
  - `UsersPage.test.tsx`:
    - "cria conta só com e-mail e papel";
    - "mostra Convite pendente, Aguardando confirmação e Confirmado";
    - "reenviar convite chama enviar-link-senha".
  - `RedefinirSenhaPage.test.tsx`: "convite mostra Criar sua senha".
  - Ajustar as asserções antigas que procuravam o rótulo "Usuário" ou a troca de senha pelo Gestor.

**Interfaces:** consome o contrato HTTP das Tasks 2 e 3.

- [ ] **Passo 1:** escrever os testes acima.
- [ ] **Passo 2:** rodar `npx vitest run src/pages/LoginPage.test.tsx src/pages/UsersPage.test.tsx src/pages/RedefinirSenhaPage.test.tsx` e confirmar que FALHAM.
- [ ] **Passo 3:** implementar.
- [ ] **Passo 4:** rodar `npx vitest run`, `npx tsc --noEmit` e `npm run build`. Tudo verde.
- [ ] **Passo 5: commit** `feat(login): tela de login por e-mail e etapa de confirmação do e-mail`.
