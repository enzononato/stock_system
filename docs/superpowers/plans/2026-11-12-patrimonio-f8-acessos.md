# Fase 8 — Acessos por função — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Módulo "Acessos": catálogo de sistemas, pacotes de sistemas por cargo, ficha de acessos por colaborador (login, status, equipamento opcional, quem criou/desativou e quando) e eventos registrados à mão (admissão, desligamento, mudança de função) que geram uma lista de tarefas — sem nunca pedir nem guardar senha.

**Architecture:** Módulo isolado sobre a F0: migração `0801`, schemas próprios (`schemas/acessos.py`, com uma base que recusa qualquer campo de senha), dois managers de SQL cru (`db/acesso_db.py` para catálogo/pacotes/ficha; `db/acesso_evento_db.py` para eventos e tarefas, reaproveitando o primeiro no mesmo cursor) e dois routers sob `/api/acessos`. Toda escrita grava `audit.record(..., cur=cur)` na mesma transação. No frontend, um grupo "Acessos" no menu com três telas novas em `pages/acessos/`, componentes em `components/acessos/` e duas abas plugáveis (colaborador ordem 50, equipamento ordem 60).

**Tech Stack:** FastAPI 0.115 + PyMySQL/DBUtils (SQL cru, `get_cursor`) + Pydantic 2.9 + MySQL 8.0; React 18 + TS 5.6 + Tailwind 3.4 + TanStack Query 5 + Vitest/MSW 2. Nenhuma dependência nova.

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md` — seção 6 "F8 — Acessos", matriz 14.1–14.12 (14.13 fora), seção 5 "F8", seção 8. Contexto comum: `C:/Users/enzo.jz/dev/patrimonio/_tools/contexto-planejamento.md`. Interfaces da F0: `docs/superpowers/plans/2026-11-12-patrimonio-f0-base.md` (Tasks 3, 6, 8, 11, 12, 14).

**Ambiente:** worktree `C:/Users/enzo.jz/dev/patrimonio/f8`, branch `fase/f8-acessos` (a partir de `feat/patrimonio-melhorias` com a F0 integrada). Neste plano, `RBT` = `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh C:/Users/enzo.jz/dev/patrimonio/f8 f8`. Comandos de frontend rodam em `C:/Users/enzo.jz/dev/patrimonio/f8/frontend`.

## Global Constraints

- Migrações só aditivas, em `backend/migrations/NNNN_descricao.sql` + `NNNN_descricao.down.sql`, faixa da F8: `0800–0899`. O executor divide comandos por `;` no fim da linha: nada de `; -- comentário` na mesma linha nem `;` dentro de literais.
- Tabelas do spec (seção 5 "F8"): `sistemas` com carga inicial **AD, E-mail, VPN, Promax, Biz Deliveries, WMS**; `pacotes_funcao`; `colaborador_acessos`; `acesso_eventos`; `acesso_tarefas`. **Não existe coluna de senha.**
- "O sistema nunca pede nem guarda senha; os formulários deixam isso explícito." / "Nenhuma senha de sistema de terceiros no módulo de Acessos." Nenhum schema de entrada aceita chave com `senha`, `password`, `passwd` ou `pwd`.
- Catálogo de sistemas: Gestor cadastra, edita e desativa. Pacotes: Gestor edita. Ficha e eventos: Gestor e Técnico. Jovem Aprendiz não usa o módulo. Mensagens 403 padrão: `Acesso restrito ao Gestor.` / `Acesso restrito a Gestor ou Técnico.`
- Cargos vêm do cadastro de colaboradores da F0 (`GET /api/colaboradores/cargos`, coluna `colaboradores.cargo`). A revisão periódica de acessos pelo superior (14.13) está fora.
- Toda escrita grava `audit.record(ctx, entidade, entidade_id, acao, antes, depois, cur=cur)` (`app/services/audit.py`) com `ctx: AuditContext = Depends(get_audit_context)`. Entidades desta fase: `acesso`, `acesso_evento`, `acesso_sistema`, `acesso_pacote`.
- Arquivos compartilhados recebem só linhas novas, sem reordenar nem reformatar: `backend/app/main.py`, `backend/app/dependencies.py`, `frontend/src/App.tsx`, `frontend/src/components/layout/navigation.ts`. Não editar componentes nem registros de abas da F0 (StatusPill, Pagination, Button, `lib/status.ts`, `registry.ts`).
- Listagens que crescem com o uso usam `Pagination` + `usePageSize` (25/50/100), `StatusPill`, `Button loading` da F0. Texto azul: `text-brand-text` (nunca `text-primary`). Unidade/revenda exibidas com `upper()`.
- Textos de interface, mensagens e commits em português; commits `tipo(escopo): descrição` terminando com linha em branco + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- Abas: `features/colaborador-detail/tabs/acessos.tsx` ordem **50**; `features/item-detail/tabs/acessos.tsx` ordem **60**; ambas `roles: ['Gestor', 'Técnico']`.

## Review Focus

- Colaborador desligado que recebeu equipamento antes da F0 (só `items.cpf`, às vezes gravado com máscara `111.444.777-35`, sem `colaborador_id`) → o desligamento lista o equipamento e não conclui (teste na Task 4).
- Técnico registra de novo um login que já foi desativado, digitado com outra caixa ou espaços (`"  maria.silva "` × `Maria.Silva`) → reativa o mesmo registro, sem 409 nem linha duplicada (teste na Task 4).
- Dois cliques ou dois técnicos resolvendo a mesma tarefa ao mesmo tempo → o segundo recebe 409 "Esta tarefa já foi resolvida." e só um acesso é criado (testes nas Tasks 6 e 11).
- Cargo vindo do RH com caixa/espaços diferentes do cargo do pacote (`"VENDEDOR  "` × `Vendedor`) → a admissão acha o pacote (teste na Task 5).
- Campo de senha com outra grafia (`Senha`, `senha_vpn`, `PASSWORD`, `pwd`) num payload qualquer do módulo → 422 com a mensagem clara, nada gravado (testes nas Tasks 2 e 4).

## Contrato da API (todas sob `/api/acessos`)

| Método e caminho | Papel | Entrada | Resposta |
|---|---|---|---|
| `GET /sistemas` | G, T | `include_inactive: bool = False` | `list[SistemaResponse]` por nome |
| `POST /sistemas` | G | `SistemaCreate` | 201 `SistemaResponse` |
| `PUT /sistemas/{sistema_id}` | G | `SistemaUpdate` (chave ausente = não muda) | `SistemaResponse` |
| `GET /pacotes` | G, T | — | `list[PacoteResponse]`: cargos com ≥1 sistema, por cargo; sistemas por nome, inclusive desativados |
| `PUT /pacotes` | G | `PacoteUpdate` (substitui o pacote; lista vazia apaga) | `PacoteResponse` |
| `GET /ficha` | G, T | `colaborador_id`, `item_id`, `status` (exige `colaborador_id` ou `item_id`) | `list[AcessoResponse]`: ativos primeiro, depois sistema e login |
| `POST /ficha` | G, T | `AcessoCreate` | 201 `AcessoResponse` |
| `PUT /ficha/{acesso_id}` | G, T | `AcessoUpdate` (chave ausente = não muda; `item_id: null` limpa) | `AcessoResponse` |
| `POST /ficha/{acesso_id}/desativar` | G, T | — | `AcessoResponse` |
| `GET /colaboradores/{colaborador_id}/equipamentos` | G, T | — | `list[EquipamentoResumo]` |
| `GET /eventos` | G, T | `status`, `tipo`, `colaborador_id`, `busca`, `limit` (25, 1–100), `offset` | `Paginated[EventoResumo]`, `aberto_em DESC, id DESC` |
| `POST /eventos` | G, T | `EventoCreate` | 201 `EventoDetalhe` |
| `GET /eventos/{evento_id}` | G, T | — | `EventoDetalhe` |
| `POST /eventos/{evento_id}/concluir` | G, T | — | `EventoDetalhe` |
| `POST /eventos/{evento_id}/cancelar` | G, T | `EventoCancelar` | `EventoDetalhe` |
| `POST /tarefas/{tarefa_id}/resolver` | G, T | `TarefaResolver` | `EventoDetalhe` |

Erros de domínio: `NaoEncontrado` → 404, `Conflito` → 409, `Invalido` → 422, sempre `{"detail": "<mensagem>"}`. Mensagens exatas (usadas nos testes):

- Recusa de senha: `MSG_SEM_SENHA = "O Patrimônio não guarda senhas: remova o campo de senha e registre só o login."`
- 404: `Sistema não encontrado.` · `Colaborador não encontrado.` · `Equipamento não encontrado.` · `Acesso não encontrado.` · `Evento não encontrado.` · `Tarefa não encontrada.`
- 409: `Já existe um sistema com esse nome.` · `Este login já está ativo para o colaborador neste sistema.` · `Já existe um acesso com esse login para o colaborador neste sistema.` (edição) · `O acesso já está desativado.` · `Já existe um evento de acesso aberto para este colaborador.` · `O evento não está aberto.` · `Esta tarefa já foi resolvida.` · `Não é possível concluir: <pendências separadas por "; ">.`
- 422: `Sistema inválido ou desativado.` (pacote) · `Sistema desativado: não é possível registrar acesso nele.` · `Informe colaborador_id ou item_id.` · `Informe o cargo: o colaborador não tem cargo cadastrado.` · `O cargo "<cargo>" não tem pacote de acessos definido.` · `Informe o cargo anterior.` · `Informe o cargo novo.` · `O cargo novo é igual ao anterior.` · `Informe o login criado no sistema.` · `Informe a observação para marcar como não se aplica.` · `Informe o motivo para manter o acesso.` · `Manter só vale para tarefas de desativar em mudança de função.`
- 400: `Nenhum campo para atualizar.`
- Pendências: `1 tarefa pendente` / `N tarefas pendentes`; `1 equipamento ainda com o colaborador` / `N equipamentos ainda com o colaborador`.
- Observações automáticas: tarefa "criar" de sistema em que o colaborador já tem acesso ativo → `Acesso já ativo: <login>`; "manter" → `Acesso mantido: <observação digitada>`.

---

### Task 1: Migração 0801 (tabelas de acessos e carga de sistemas)

**Files:**
- Create: `backend/migrations/0801_acessos.sql`, `backend/migrations/0801_acessos.down.sql`, `backend/tests/acessos_helpers.py`, `backend/tests/test_acessos_migracao.py`

**Interfaces:**
- Consumes: `colaboradores` (F0 `0001`), `items.colaborador_id` (F0 `0004`), `app.db.migrations.MIGRATIONS_DIR` e `executar_sql(cur, texto)`.
- Produces: as cinco tabelas abaixo; helpers de teste em `tests/acessos_helpers.py` (todos importam `app.db.database_mysql.get_cursor` **dentro** da função, regra do `conftest.py`):
  - `CPF_MARIA = "11144477735"`, `CPF_JOAO = "52998224725"`
  - `criar_colaborador(nome: str = "Maria Silva", cpf: str = CPF_MARIA, cargo: str | None = "Vendedor", situacao: str = "ativo") -> int` (INSERT direto com `origem='manual'`, `criado_em`/`atualizado_em` = `NOW()`)
  - `criar_sistema(nome: str, ativo: bool = True) -> int`; `definir_pacote(cargo: str, *sistema_ids: int) -> None`
  - `criar_acesso(colaborador_id: int, sistema_id: int, login: str, status: str = "ativo", item_id: int | None = None) -> int` (`criado_por='fixture'`, `criado_em=NOW()`)
  - `vincular_item(item_id: int, *, colaborador_id: int | None = None, cpf: str | None = None, status: str = "Indisponível") -> None` (UPDATE de `colaborador_id`, `cpf`, `status`)
  - `contar(tabela: str, where: str = "1=1", params: tuple = ()) -> int`; `auditoria(entidade: str, acao: str | None = None) -> list[dict]` (linhas de `audit_log` em ordem de `id`, `antes`/`depois` já com `json.loads`)

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_acessos_migracao.py` (`pytestmark = pytest.mark.integration`):
  - `test_tabelas_de_acessos_existem(limpar_banco)`: `information_schema.tables` do schema atual contém `sistemas`, `pacotes_funcao`, `colaborador_acessos`, `acesso_eventos`, `acesso_tarefas`.
  - `test_nenhuma_coluna_de_senha(limpar_banco)`: nenhuma coluna dessas cinco tabelas casa com `re.compile(r"senha|pass|pwd", re.I)`.
  - `test_carga_inicial_de_sistemas_idempotente(limpar_banco)`: lê `MIGRATIONS_DIR / "0801_acessos.sql"`, executa duas vezes com `from app.db.migrations import executar_sql as executar_migracao` (`with get_cursor() as cur: executar_migracao(cur, texto)`; o nome evita confusão com a fixture `executar_sql` do conftest; o `limpar_banco` esvaziou `sistemas`); `SELECT nome, ativo FROM sistemas ORDER BY nome` == `[("AD",1), ("Biz Deliveries",1), ("E-mail",1), ("Promax",1), ("VPN",1), ("WMS",1)]`.
  - `test_down_remove_e_up_recria(limpar_banco)`: executa o `.down.sql` → as cinco tabelas somem; executa o `.sql` → voltam e `contar("sistemas") == 6`.
  - `test_login_unico_ignora_caixa(limpar_banco)`: `criar_acesso(c, s, "Maria.Silva")` e depois `criar_acesso(c, s, "maria.silva")` levanta `pymysql.err.IntegrityError` com `args[0] == 1062`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_acessos_migracao.py -q` → FAIL (tabelas inexistentes).
- [ ] **Step 3: Implementar a migração** exatamente com estas definições (todas com `CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci` — a comparação sem caixa/acento de `nome`, `cargo` e `login` depende disso — e `CREATE TABLE IF NOT EXISTS`):

```sql
-- sistemas
id INT AUTO_INCREMENT PK, nome VARCHAR(100) NOT NULL, descricao VARCHAR(255) NULL,
ativo TINYINT(1) NOT NULL DEFAULT 1, UNIQUE KEY uq_sistemas_nome (nome)
-- pacotes_funcao
id INT AUTO_INCREMENT PK, cargo VARCHAR(150) NOT NULL, sistema_id INT NOT NULL,
UNIQUE KEY uq_pacotes_funcao (cargo, sistema_id), FK sistema_id → sistemas(id)
-- colaborador_acessos
id INT AUTO_INCREMENT PK, colaborador_id INT NOT NULL, sistema_id INT NOT NULL,
login VARCHAR(150) NOT NULL, status ENUM('ativo','desativado') NOT NULL DEFAULT 'ativo',
item_id INT NULL, criado_por VARCHAR(100) NOT NULL, criado_em DATETIME NOT NULL,
desativado_por VARCHAR(100) NULL, desativado_em DATETIME NULL,
UNIQUE KEY uq_colaborador_acessos (colaborador_id, sistema_id, login), INDEX (item_id),
FK colaborador_id → colaboradores(id), FK sistema_id → sistemas(id),
FK item_id → items(id) ON DELETE SET NULL
-- acesso_eventos
id INT AUTO_INCREMENT PK, colaborador_id INT NOT NULL,
tipo ENUM('admissao','desligamento','mudanca_funcao') NOT NULL,
cargo_anterior VARCHAR(150) NULL, cargo_novo VARCHAR(150) NULL,
status ENUM('aberto','concluido','cancelado') NOT NULL DEFAULT 'aberto',
observacao VARCHAR(500) NULL, aberto_por VARCHAR(100) NOT NULL, aberto_em DATETIME NOT NULL,
concluido_por VARCHAR(100) NULL, concluido_em DATETIME NULL,
cancelado_por VARCHAR(100) NULL, cancelado_em DATETIME NULL, motivo_cancelamento VARCHAR(500) NULL,
INDEX (colaborador_id, status), INDEX (status, aberto_em), FK colaborador_id → colaboradores(id)
-- acesso_tarefas
id INT AUTO_INCREMENT PK, evento_id INT NOT NULL, sistema_id INT NOT NULL, acesso_id INT NULL,
acao ENUM('criar','desativar') NOT NULL,
status ENUM('pendente','feito','nao_se_aplica') NOT NULL DEFAULT 'pendente',
login VARCHAR(150) NULL, feito_por VARCHAR(100) NULL, feito_em DATETIME NULL, observacao VARCHAR(500) NULL,
INDEX (evento_id), FK evento_id → acesso_eventos(id) ON DELETE CASCADE, FK sistema_id → sistemas(id),
FK acesso_id → colaborador_acessos(id) ON DELETE SET NULL
```

  Colunas além do spec (decisão deste plano): `acesso_eventos.observacao`, `concluido_por`, `cancelado_por`, `cancelado_em`, `motivo_cancelamento`; `acesso_tarefas.acesso_id` (qual acesso a tarefa desativa/criou — um colaborador pode ter dois logins no mesmo sistema). Carga, num único `INSERT IGNORE INTO sistemas (nome, descricao, ativo) VALUES ... ;` com `;` só no fim: `('AD', 'Active Directory (usuário de rede)', 1)`, `('E-mail', 'E-mail corporativo', 1)`, `('VPN', 'Acesso remoto (VPN)', 1)`, `('Promax', 'ERP Promax', 1)`, `('Biz Deliveries', 'Biz Deliveries (entregas)', 1)`, `('WMS', 'WMS (armazém)', 1)`. O `.down.sql` faz `DROP TABLE IF EXISTS` na ordem `acesso_tarefas`, `acesso_eventos`, `colaborador_acessos`, `pacotes_funcao`, `sistemas`. Escrever também `tests/acessos_helpers.py`.
- [ ] **Step 4: Rodar** `RBT tests/test_acessos_migracao.py -q` → PASS; depois `RBT -q -m "not mudanca_esperada"` → suíte inteira verde (o `test_reverter_ultima_e_reaplicar` da F0 passa a reverter a 0801).
- [ ] **Step 5: Commit** `feat(acessos): tabelas do módulo de acessos com carga inicial de sistemas`.

### Task 2: Schemas do módulo (recusa de senha) e rótulos do 422

**Files:**
- Create: `backend/app/schemas/acessos.py`, `backend/tests/test_acessos_schemas.py`
- Modify: `backend/app/main.py` (só acrescentar chaves em `_ROTULOS_DE_CAMPO`)

**Interfaces:**
- Produces (em `app/schemas/acessos.py`):
  - `MSG_SEM_SENHA` (texto do Contrato); `PALAVRAS_SENHA = ("senha", "password", "passwd", "pwd")`
  - `normalizar_cargo(valor: str | None) -> str | None` (`" ".join(valor.split())`; vazio → `None`); `chave_cargo(valor: str | None) -> str` (normalizado, sem acento via `unicodedata` NFKD, `casefold()`; `None` → `""`)
  - `class EntradaSemSenha(BaseModel)`: `model_config = ConfigDict(extra="forbid")` + `@model_validator(mode="before")` que levanta `ValueError(MSG_SEM_SENHA)` quando alguma chave do dict, em minúsculas, contém um item de `PALAVRAS_SENHA`. Todos os schemas de entrada abaixo herdam dela.
  - Literais: `TipoEvento = Literal["admissao", "desligamento", "mudanca_funcao"]`, `StatusEvento = Literal["aberto", "concluido", "cancelado"]`, `AcaoTarefa = Literal["criar", "desativar"]`, `StatusTarefa = Literal["pendente", "feito", "nao_se_aplica"]`, `StatusAcesso = Literal["ativo", "desativado"]`.
  - Entrada: `SistemaCreate(nome: str, descricao: str | None = None)`; `SistemaUpdate(nome: str | None = None, descricao: str | None = None, ativo: bool | None = None)`; `PacoteUpdate(cargo: str, sistema_ids: list[int])`; `AcessoCreate(colaborador_id: int, sistema_id: int, login: str, item_id: int | None = None)`; `AcessoUpdate(login: str | None = None, item_id: int | None = None)`; `EventoCreate(colaborador_id: int, tipo: TipoEvento, cargo_anterior: str | None = None, cargo_novo: str | None = None, observacao: str | None = None)`; `TarefaResolver(status: Literal["feito", "nao_se_aplica", "manter"], login: str | None = None, item_id: int | None = None, observacao: str | None = None)`; `EventoCancelar(motivo_cancelamento: str)`.
  - Validações: textos aparados com `blank_to_none` (`app/schemas/validators.py`); `nome` obrigatório ≤100 ("Informe o nome do sistema."); `descricao` ≤255; `cargo`/`cargo_anterior`/`cargo_novo` passam por `normalizar_cargo`, ≤150 (`PacoteUpdate.cargo` vazio → "Informe o cargo."); `login` obrigatório em `AcessoCreate` ("Informe o login."), ≤150 ("Login muito longo (máximo 150 caracteres)."), em `AcessoUpdate`/`TarefaResolver` opcional com o mesmo limite; `observacao` ≤500; `motivo_cancelamento` obrigatório ("Informe o motivo do cancelamento.") ≤500; `sistema_ids` sem repetidos, mantendo a ordem.
  - Respostas (`BaseModel` comum): `SistemaResponse(id, nome, descricao: str | None, ativo: bool)`; `PacoteResponse(cargo: str, sistemas: list[SistemaResponse])`; `AcessoResponse(id, colaborador_id, colaborador_nome, sistema_id, sistema_nome, login, status: StatusAcesso, item_id: int | None, item_descricao: str | None, criado_por, criado_em: datetime, desativado_por: str | None, desativado_em: datetime | None)`; `EquipamentoResumo(id, tipo, brand, model, identificador: str | None, status: str, revenda: str | None)`; `TarefaResponse(id, evento_id, sistema_id, sistema_nome, acao: AcaoTarefa, status: StatusTarefa, login, acesso_id: int | None, feito_por, feito_em: datetime | None, observacao)`; `EventoResumo(id, colaborador_id, colaborador_nome, tipo, cargo_anterior, cargo_novo, status, observacao, aberto_por, aberto_em, concluido_por, concluido_em, cancelado_por, cancelado_em, motivo_cancelamento, total_tarefas: int, tarefas_pendentes: int)`; `EventoDetalhe(EventoResumo)` + `colaborador_situacao: str`, `tarefas: list[TarefaResponse]`, `equipamentos: list[EquipamentoResumo]`, `pode_concluir: bool`, `pendencias: list[str]`.
  - `_ROTULOS_DE_CAMPO` ganha só: `"login": "Login"`, `"sistema_id": "Sistema"`, `"sistema_ids": "Sistemas"`, `"cargo_anterior": "Cargo anterior"`, `"cargo_novo": "Cargo novo"`, `"motivo_cancelamento": "Motivo do cancelamento"`.

- [ ] **Step 1: Escrever os testes** (`unit`) em `backend/tests/test_acessos_schemas.py`:
  - `test_campos_de_senha_recusados` parametrizado: `(SistemaCreate, {"nome": "X", "senha": "1"})`, `(AcessoCreate, {"colaborador_id": 1, "sistema_id": 1, "login": "a", "password": "x"})`, `(AcessoUpdate, {"Senha_AD": "x"})`, `(PacoteUpdate, {"cargo": "Vendedor", "sistema_ids": [], "senha_vpn": "x"})`, `(EventoCreate, {"colaborador_id": 1, "tipo": "admissao", "pwd": "x"})`, `(TarefaResolver, {"status": "feito", "login": "a", "PASSWORD": "x"})`, `(EventoCancelar, {"motivo_cancelamento": "x", "passwd": "x"})` → `pytest.raises(ValidationError)` e `MSG_SEM_SENHA in str(exc.value)`.
  - `test_campo_desconhecido_recusado`: `AcessoCreate(colaborador_id=1, sistema_id=1, login="a", apelido="x")` → `ValidationError`.
  - `test_respostas_sem_campo_de_senha`: nenhum nome em `model_fields` de nenhuma classe `*Response`/`EventoResumo`/`EventoDetalhe`/`EquipamentoResumo` contém um item de `PALAVRAS_SENHA`.
  - `test_login_aparado_e_obrigatorio`: `"  maria.silva "` → `"maria.silva"`; `"   "` → erro com "Informe o login."; 151 caracteres → erro com "Login muito longo".
  - `test_pacote_normaliza_cargo_e_remove_repetidos`: `cargo="  Vendedor   Externo "` → `"Vendedor Externo"`; `sistema_ids=[3, 1, 3]` → `[3, 1]`; `cargo="  "` → erro "Informe o cargo.".
  - `test_chave_cargo`: `chave_cargo("VENDEDOR  ") == chave_cargo("vendedor") == chave_cargo("Vendédor")`; `chave_cargo(None) == ""`.
  - `test_resolver_status_validos`: `"manter"` aceito; `"outro"` → `ValidationError`.
  - `test_cancelar_exige_motivo`: `motivo_cancelamento="  "` → erro "Informe o motivo do cancelamento.".
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_acessos_schemas.py -q` → FAIL (módulo inexistente).
- [ ] **Step 3: Implementar** `app/schemas/acessos.py` e as chaves novas em `_ROTULOS_DE_CAMPO`.
- [ ] **Step 4: Rodar** `RBT tests/test_acessos_schemas.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(acessos): schemas do módulo de acessos sem nenhum campo de senha`.

### Task 3: Catálogo de sistemas e pacotes por função (manager e API)

**Files:**
- Create: `backend/app/db/acesso_db.py`, `backend/app/routers/acessos.py`, `backend/tests/test_acessos_catalogo.py`
- Modify: `backend/app/dependencies.py` (bloco novo no fim do arquivo), `backend/app/main.py` (uma linha de import e uma de `include_router`, depois das existentes)

**Interfaces:**
- Consumes: schemas da Task 2; `AuditContext`, `audit.record`, `audit.diff`; `gestor_only`, `gestor_or_tecnico`, `get_audit_context`, `CurrentUser`.
- Produces (em `app/db/acesso_db.py`):
  - `class AcessoErro(Exception)` com atributo de classe `status_code = 400`; `NaoEncontrado(AcessoErro)` 404; `Conflito(AcessoErro)` 409; `Invalido(AcessoErro)` 422.
  - `@contextmanager cursor_ou_novo(cur=None, commit: bool = True)` — devolve `cur` quando recebido; senão abre `get_cursor(commit=commit)`.
  - `AcessoDBManager` (não cria tabelas no `__init__`; a migração cria): `listar_sistemas(*, include_inactive: bool = False) -> list[dict]`; `obter_sistema(sistema_id: int, cur=None) -> dict | None`; `criar_sistema(ctx, *, nome: str, descricao: str | None) -> dict`; `atualizar_sistema(ctx, sistema_id: int, dados: dict) -> dict`; `listar_pacotes() -> list[dict]` (`[{"cargo", "sistemas": [sistema]}]`); `sistemas_do_pacote(cargo: str, *, somente_ativos: bool = True, cur=None) -> list[dict]` (`WHERE p.cargo = %s` com `normalizar_cargo(cargo)`; a collation resolve caixa/acento); `definir_pacote(ctx, *, cargo: str, sistema_ids: list[int]) -> dict` (`DELETE` do cargo + `INSERT` dos ids, numa transação).
  - Erros: nome repetido (MySQL 1062) → `Conflito("Já existe um sistema com esse nome.")`; sistema inexistente → `NaoEncontrado("Sistema não encontrado.")`; id de pacote inexistente ou com `ativo=0` → `Invalido("Sistema inválido ou desativado.")`.
  - Auditoria (com `cur`): `acesso_sistema`/`criar` (`depois` = sistema); `atualizar_sistema` grava `desativar` ou `reativar` quando `ativo` muda, senão `editar` (com `audit.diff`); `acesso_pacote`/`editar` com `entidade_id=None`, `antes={"cargo", "sistemas": [nomes ordenados]}`, `depois` idem.
  - `app.dependencies.get_acesso_db()` (`@lru_cache`, import de `AcessoDBManager` dentro da função) — bloco no fim do arquivo com o comentário `# ── F8: Acessos ──`.
  - `app/routers/acessos.py`: `router = APIRouter(prefix="/api/acessos", tags=["acessos"])`; `erro_http(e: AcessoErro) -> HTTPException` (usa `e.status_code` e `str(e)`; também usado pelo router de eventos); rotas `/sistemas` e `/pacotes` do Contrato. `PUT /sistemas/{id}` usa `body.model_dump(exclude_unset=True)`; vazio → 400 "Nenhum campo para atualizar.".

- [ ] **Step 1: Escrever os testes** (`integration`) em `backend/tests/test_acessos_catalogo.py`, criando sistemas com `criar_sistema` dos helpers:
  - `test_tecnico_lista_so_sistemas_ativos`: AD ativo e Promax inativo → `client_tecnico.get("/api/acessos/sistemas")` traz só AD; com `?include_inactive=true` traz os dois, Promax com `ativo is False`.
  - `test_aprendiz_nao_acessa_o_catalogo`: `client_aprendiz.get(...)` → 403, `detail == "Acesso restrito a Gestor ou Técnico."`.
  - `test_criar_sistema_exige_gestor`: `client_tecnico.post("/api/acessos/sistemas", json={"nome": "GLPI"})` → 403 "Acesso restrito ao Gestor.".
  - `test_criar_sistema`: Gestor posta `{"nome": " GLPI ", "descricao": ""}` → 201 `{"nome": "GLPI", "descricao": None, "ativo": True}`; `auditoria("acesso_sistema", "criar")` tem uma linha com `username == "gestor.teste"`.
  - `test_sistema_duplicado_409`: depois de "GLPI", `{"nome": "glpi"}` → 409 "Já existe um sistema com esse nome.".
  - `test_desativar_e_reativar_sistema`: `PUT {"ativo": false}` → `ativo False` e auditoria `desativar`; `GET` padrão não o lista; `PUT {"ativo": true}` → auditoria `reativar`.
  - `test_editar_sistema_404_e_corpo_vazio`: `PUT /sistemas/999` → 404 "Sistema não encontrado."; `PUT {}` → 400 "Nenhum campo para atualizar.".
  - `test_senha_no_corpo_do_sistema_422`: `{"nome": "X", "senha": "123"}` → 422, `detail == MSG_SEM_SENHA`, `contar("sistemas", "nome='X'") == 0`.
  - `test_definir_pacote_normaliza_e_substitui`: `PUT /pacotes {"cargo": " Vendedor ", "sistema_ids": [ad, email, ad]}` → 200 `cargo == "Vendedor"`, nomes `["AD", "E-mail"]`; novo `PUT` com `[vpn]` → `GET /pacotes` traz `[{"cargo": "Vendedor", "sistemas": [VPN]}]`.
  - `test_pacote_vazio_remove_o_cargo`: `sistema_ids: []` → `GET /pacotes == []`.
  - `test_pacote_com_sistema_desativado_ou_inexistente_422`: id de sistema inativo ou `9999` → 422 "Sistema inválido ou desativado.", pacote inalterado.
  - `test_pacote_tecnico_le_mas_nao_edita`: `GET` 200, `PUT` 403 "Acesso restrito ao Gestor.".
  - `test_pacote_audita_antes_e_depois`: pacote `[AD]` → `[E-mail, VPN]` gera `acesso_pacote`/`editar` com `antes == {"cargo": "Vendedor", "sistemas": ["AD"]}` e `depois == {"cargo": "Vendedor", "sistemas": ["E-mail", "VPN"]}`.
  - `test_pacote_lista_sistema_desativado_depois`: pacote com AD; AD desativado → `GET /pacotes` ainda traz AD com `ativo False`; `AcessoDBManager().sistemas_do_pacote("vendedor")` (padrão `somente_ativos=True`) == `[]`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_acessos_catalogo.py -q` → FAIL (404 nas rotas).
- [ ] **Step 3: Implementar** `acesso_db.py` (erros, `cursor_ou_novo`, sistemas, pacotes), `routers/acessos.py`, `get_acesso_db`, e em `main.py` acrescentar `from app.routers import acessos` e `app.include_router(acessos.router)` depois das linhas existentes.
- [ ] **Step 4: Rodar** `RBT tests/test_acessos_catalogo.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(acessos): catálogo de sistemas e pacotes por função`.

### Task 4: Ficha de acessos e equipamentos com o colaborador

**Files:**
- Modify: `backend/app/db/acesso_db.py`, `backend/app/routers/acessos.py`
- Create: `backend/tests/test_acessos_ficha.py`

**Interfaces:**
- Consumes: Task 3 (`AcessoDBManager`, erros, `cursor_ou_novo`, `erro_http`).
- Produces (em `acesso_db.py`):
  - `STATUS_COM_COLABORADOR = ("Pendente", "Indisponível", "Pendente Devolução")`
  - `descrever_item(linha: dict) -> str | None` — `" ".join` de `tipo`, `brand`, `model` não vazios, mais `" · " + identificador` quando houver; tudo vazio → `None`. Ex.: `"Notebook Dell Latitude 5420 · SN-AB12"`.
  - `obter_colaborador(cur, colaborador_id: int, *, travar: bool = False) -> dict | None` — `SELECT id, nome, cpf, cargo, situacao FROM colaboradores WHERE id=%s` (+ `FOR UPDATE` com `travar`).
  - `AcessoDBManager.listar_acessos(*, colaborador_id: int | None = None, item_id: int | None = None, status: str | None = None, cur=None) -> list[dict]` (JOIN `sistemas` e `colaboradores`, LEFT JOIN `items`; `item_descricao` via `descrever_item`; `ORDER BY a.status = 'ativo' DESC, s.nome, a.login`); `obter_acesso(acesso_id: int, cur=None) -> dict | None`.
  - `criar_acesso(ctx, *, colaborador_id: int, sistema_id: int, login: str, item_id: int | None = None, cur=None) -> dict`: valida colaborador (`NaoEncontrado`), sistema (`NaoEncontrado`; `ativo=0` → `Invalido("Sistema desativado: não é possível registrar acesso nele.")`), item quando informado (existe e `is_active=1`, senão `NaoEncontrado("Equipamento não encontrado.")`); `SELECT ... FOR UPDATE` pelo trio `(colaborador_id, sistema_id, login)` (a collation ignora caixa/acento): `ativo` → `Conflito("Este login já está ativo para o colaborador neste sistema.")`; `desativado` → UPDATE `login` (grafia nova), `status='ativo'`, `item_id`, `criado_por=ctx.username`, `criado_em=agora`, `desativado_por=NULL`, `desativado_em=NULL` e auditoria `acesso`/`reativar`; senão INSERT e auditoria `acesso`/`criar`. Datas com `datetime.now()` (mesmo padrão do `history`). Devolve `obter_acesso`.
  - `atualizar_acesso(ctx, acesso_id: int, dados: dict) -> dict` (chaves `login`, `item_id`; 1062 → `Conflito("Já existe um acesso com esse login para o colaborador neste sistema.")`; auditoria `editar` com `audit.diff`).
  - `desativar_acesso(ctx, acesso_id: int, cur=None) -> dict` (`NaoEncontrado("Acesso não encontrado.")`; já desativado → `Conflito("O acesso já está desativado.")`; grava `desativado_por=ctx.username`, `desativado_em=agora`; auditoria `desativar`).
  - `equipamentos_com_colaborador(colaborador_id: int, cur=None) -> list[dict]` — itens com o colaborador **ou**, nos registros antigos sem `colaborador_id`, com o CPF dele. O filtro de status protege contra `colaborador_id` esquecido num item já devolvido:

```sql
SELECT i.id, i.tipo, i.brand, i.model, i.identificador, i.status, i.revenda
FROM items i
WHERE i.is_active = 1
  AND i.status IN ('Pendente', 'Indisponível', 'Pendente Devolução')
  AND (i.colaborador_id = %s
       OR (i.colaborador_id IS NULL AND REGEXP_REPLACE(COALESCE(i.cpf, ''), '[^0-9]', '') = %s))
ORDER BY i.id
```
    Quando o CPF do colaborador não tem 11 dígitos, usar só a primeira condição (senão `''` casaria itens sem CPF).
  - Rotas `/ficha`, `/ficha/{acesso_id}`, `/ficha/{acesso_id}/desativar` e `/colaboradores/{colaborador_id}/equipamentos` do Contrato (todas `gestor_or_tecnico`); `GET /ficha` sem `colaborador_id` nem `item_id` → 422 "Informe colaborador_id ou item_id."; `PUT` vazio → 400 "Nenhum campo para atualizar."; equipamentos de colaborador inexistente → 404 "Colaborador não encontrado.".

- [ ] **Step 1: Escrever os testes** (`integration`) em `backend/tests/test_acessos_ficha.py` (usar `criar_item` do conftest e os helpers):
  - `test_registrar_acesso_com_equipamento`: Técnico posta `{"colaborador_id": c, "sistema_id": ad, "login": "maria.silva", "item_id": item}` → 201 com `status "ativo"`, `criado_por "tecnico.teste"`, `colaborador_nome "Maria Silva"`, `sistema_nome "AD"`, `item_descricao` começando com `"Notebook Dell Latitude 5420 · SN-"`; auditoria `acesso`/`criar`.
  - `test_login_ativo_repetido_409`: o mesmo trio de novo → 409 "Este login já está ativo para o colaborador neste sistema.".
  - `test_criar_acesso_reativa_login_desativado_ignorando_caixa_e_espacos` (Review Focus): `criar_acesso(c, ad, "Maria.Silva", status="desativado")` e POST com `login "  maria.silva "` → 201 com o **mesmo id**, `login "maria.silva"`, `status "ativo"`, `desativado_em None`; `contar("colaborador_acessos") == 1`; auditoria `acesso`/`reativar`.
  - `test_sistema_desativado_422`, `test_colaborador_inexistente_404`, `test_equipamento_inexistente_404` (mensagens do Contrato).
  - `test_senha_em_qualquer_grafia_422` (Review Focus): POST com `"Senha"` e PUT com `"senha_vpn"` → 422 `detail == MSG_SEM_SENHA`; `contar("colaborador_acessos") == 0` após o POST.
  - `test_listar_ficha_exige_filtro_e_ordena`: `GET /ficha` → 422 "Informe colaborador_id ou item_id."; com acessos VPN ativo, AD desativado e AD ativo → ordem `[AD ativo, VPN ativo, AD desativado]`; `?item_id=` traz só os do item.
  - `test_editar_login_e_limpar_equipamento`: `PUT {"item_id": null}` → `item_id None`; `PUT {"login": "<login de outro acesso do mesmo colaborador e sistema>"}` → 409 "Já existe um acesso com esse login para o colaborador neste sistema."; `PUT {}` → 400.
  - `test_desativar_acesso`: → `status "desativado"`, `desativado_por "tecnico.teste"`, `desativado_em` preenchido; de novo → 409 "O acesso já está desativado."; auditoria `acesso`/`desativar`.
  - `test_aprendiz_nao_ve_a_ficha`: 403 "Acesso restrito a Gestor ou Técnico.".
  - `test_equipamentos_inclui_legado_por_cpf_com_mascara` (Review Focus): item A `vincular_item(colaborador_id=c)`; item B `vincular_item(cpf="111.444.777-35", status="Pendente Devolução")` sem `colaborador_id`; item C `colaborador_id=c` e `status="Disponível"`; item D com CPF de João; item E `colaborador_id=c` e `is_active=0` → `GET /colaboradores/{c}/equipamentos` devolve ids `[A, B]`; colaborador 999 → 404.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_acessos_ficha.py -q` → FAIL.
- [ ] **Step 3: Implementar** os métodos e as rotas.
- [ ] **Step 4: Rodar** `RBT tests/test_acessos_ficha.py tests/test_acessos_catalogo.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(acessos): ficha de acessos do colaborador e equipamentos em posse dele`.

### Task 5: Abertura e consulta de eventos (admissão, desligamento, mudança de função)

**Files:**
- Create: `backend/app/db/acesso_evento_db.py`, `backend/app/routers/acesso_eventos.py`, `backend/tests/test_acesso_eventos.py`
- Modify: `backend/app/dependencies.py` (no bloco da F8), `backend/app/main.py` (mais um import e um `include_router`)

**Interfaces:**
- Consumes: `AcessoDBManager` (`sistemas_do_pacote`, `listar_acessos`, `equipamentos_com_colaborador`, `obter_colaborador`, `cursor_ou_novo`, erros), `normalizar_cargo`, `chave_cargo`, `erro_http`.
- Produces (em `app/db/acesso_evento_db.py`):
  - `diferenca_pacotes(anterior: list[dict], novo: list[dict]) -> tuple[list[dict], list[dict]]` — `(entram, saem)` comparando por `id`, cada lista ordenada por `nome.casefold()`.
  - `pendencias_do_evento(tipo: str, tarefas_pendentes: int, equipamentos: int) -> list[str]` — textos do Contrato; equipamentos só contam quando `tipo == "desligamento"`.
  - `AcessoEventoDBManager(acessos: AcessoDBManager)`:
    - `abrir(ctx, *, colaborador_id: int, tipo: str, cargo_anterior: str | None, cargo_novo: str | None, observacao: str | None) -> dict` (devolve `obter(...)`).
    - `listar(*, status: str | None = None, tipo: str | None = None, colaborador_id: int | None = None, busca: str | None = None, limit: int = 25, offset: int = 0) -> tuple[list[dict], int]` — `busca` casa `c.nome LIKE %busca%` ou, se tiver dígitos, `c.cpf LIKE %<só os dígitos>%`; `total_tarefas` e `tarefas_pendentes` por subconsulta.
    - `obter(evento_id: int, cur=None) -> dict | None` — campos de `EventoDetalhe`; tarefas `ORDER BY t.acao, s.nome, t.id` (ENUM: `criar` antes de `desativar`); `equipamentos` = `equipamentos_com_colaborador` só para `desligamento`, senão `[]`; `pendencias` só quando `status == "aberto"`; `pode_concluir = status == "aberto" and not pendencias`.
    - `_tentar_concluir(ctx, evento_id: int, cur) -> bool` — sem pendências: `status='concluido'`, `concluido_por=ctx.username`, `concluido_em=agora`, auditoria `acesso_evento`/`concluir`.
  - `app.dependencies.get_acesso_evento_db()` (`@lru_cache`, devolve `AcessoEventoDBManager(get_acesso_db())`).
  - `app/routers/acesso_eventos.py`: `router = APIRouter(prefix="/api/acessos", tags=["acessos"])` com `GET /eventos`, `POST /eventos` (201) e `GET /eventos/{evento_id}` (404 "Evento não encontrado.").
- Algoritmo de `abrir` (uma transação):
  1. `obter_colaborador(cur, id, travar=True)` → `NaoEncontrado("Colaborador não encontrado.")`. Existe evento `aberto` do colaborador → `Conflito("Já existe um evento de acesso aberto para este colaborador.")`.
  2. Cargos: **admissão** — `cargo_novo = normalizar_cargo(cargo_novo) or normalizar_cargo(colaborador.cargo)`, vazio → `Invalido("Informe o cargo: o colaborador não tem cargo cadastrado.")`, `cargo_anterior = None`; **mudança** — `cargo_anterior` obrigatório ("Informe o cargo anterior."), `cargo_novo` = informado ou o cargo atual do colaborador (vazio → "Informe o cargo novo."), `chave_cargo` iguais → "O cargo novo é igual ao anterior."; **desligamento** — `cargo_anterior = normalizar_cargo(colaborador.cargo)`, `cargo_novo = None`, campos do corpo ignorados.
  3. INSERT do evento (`aberto_por=ctx.username`, `aberto_em=agora`) e das tarefas:
     - admissão: um "criar" por sistema de `sistemas_do_pacote(cargo_novo)` (só ativos); pacote vazio → `Invalido('O cargo "<cargo_novo>" não tem pacote de acessos definido.')`.
     - mudança: `entram, saem = diferenca_pacotes(pacote(cargo_anterior), pacote(cargo_novo))`; "criar" para cada `entram`; para cada `saem`, um "desativar" por acesso **ativo** do colaborador naquele sistema (com `acesso_id` e `login`); sem acesso ativo → nenhuma tarefa.
     - desligamento: um "desativar" por acesso ativo do colaborador (qualquer sistema, mesmo desativado no catálogo), com `acesso_id` e `login`.
     - Toda tarefa "criar" de sistema em que o colaborador já tem acesso ativo nasce `nao_se_aplica`, `acesso_id`/`login` desse acesso, `observacao "Acesso já ativo: <login>"`, `feito_por=ctx.username`, `feito_em=agora`.
  4. Auditoria `acesso_evento`/`criar` com `depois={"tipo", "colaborador_id", "cargo_anterior", "cargo_novo", "tarefas": <quantidade>}`; `_tentar_concluir` (desligamento sem acessos nem equipamentos e mudança sem diferença concluem na hora).

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_acesso_eventos.py`:
  - `unit` `test_diferenca_pacotes`: anterior `[AD, E-mail, Promax]`, novo `[AD, E-mail, VPN, WMS]` → entram nomes `["VPN", "WMS"]`, saem `["Promax"]`.
  - `unit` `test_pendencias_do_evento`: `("admissao", 2, 0) == ["2 tarefas pendentes"]`; `("desligamento", 1, 1) == ["1 tarefa pendente", "1 equipamento ainda com o colaborador"]`; `("mudanca_funcao", 0, 3) == []`.
  - `test_admissao_gera_tarefas_do_pacote`: sistemas AD, E-mail, Promax (inativo); pacote "Vendedor" = AD, E-mail, Promax; Maria cargo "Vendedor"; Técnico posta `{"colaborador_id": c, "tipo": "admissao"}` → 201, `cargo_novo "Vendedor"`, `status "aberto"`, tarefas `[(sistema_nome, acao, status)] == [("AD", "criar", "pendente"), ("E-mail", "criar", "pendente")]`, `pendencias == ["2 tarefas pendentes"]`, `pode_concluir False`, `equipamentos == []`; auditoria `acesso_evento`/`criar`.
  - `test_admissao_acha_pacote_com_cargo_em_outra_caixa_e_espacos` (Review Focus): pacote salvo por `PUT /pacotes` com `cargo "Vendedor"`; colaborador com `cargo "VENDEDOR  "` → mesmas duas tarefas.
  - `test_admissao_com_acesso_ja_ativo`: Maria já tem E-mail ativo `"maria@revalle"` → tarefa E-mail `nao_se_aplica`, `observacao "Acesso já ativo: maria@revalle"`; só AD pendente.
  - `test_admissao_sem_pacote_ou_sem_cargo_422`: cargo "Motorista" sem pacote → 422 `'O cargo "Motorista" não tem pacote de acessos definido.'`; colaborador `cargo=None` sem `cargo_novo` → 422 "Informe o cargo: o colaborador não tem cargo cadastrado."; nada gravado em `acesso_eventos`.
  - `test_segundo_evento_aberto_409`: segundo POST para Maria → 409 "Já existe um evento de acesso aberto para este colaborador.".
  - `test_desligamento_gera_desativacoes_e_lista_equipamentos`: Maria com AD `"maria.silva"` ativo, VPN `"msilva"` ativo, E-mail desativado; item A com `colaborador_id`, item B legado por CPF → tarefas `[("AD", "desativar", "maria.silva"), ("VPN", "desativar", "msilva")]` com `acesso_id` preenchido; `equipamentos` ids `[A, B]`; `pendencias == ["2 tarefas pendentes", "2 equipamentos ainda com o colaborador"]`; `cargo_anterior "Vendedor"`.
  - `test_desligamento_sem_nada_conclui_na_hora`: sem acessos nem equipamentos → `status "concluido"`, `concluido_em` preenchido, auditoria `acesso_evento`/`concluir`.
  - `test_mudanca_de_funcao_gera_a_diferenca`: pacotes "Vendedor" = AD, E-mail, Promax; "Supervisor" = AD, E-mail, VPN, WMS; Maria com `cargo "Supervisor"` (RH já atualizou), AD `"maria.silva"` e Promax `"MSILVA"` ativos; POST `{"tipo": "mudanca_funcao", "cargo_anterior": "Vendedor"}` → `cargo_novo "Supervisor"`, tarefas `[("VPN", "criar"), ("WMS", "criar"), ("Promax", "desativar")]`, a de Promax com `login "MSILVA"`.
  - `test_mudanca_valida_cargos`: sem `cargo_anterior` → 422 "Informe o cargo anterior."; `cargo_anterior "supervisor "` com cargo atual "Supervisor" → 422 "O cargo novo é igual ao anterior.".
  - `test_listar_eventos_filtra_e_pagina`: desligamento de Maria sem acessos nem equipamentos (conclui na hora), depois admissão de Maria e admissão de João (abertas) → `?status=aberto&limit=1` traz `total == 2`, um item, com `total_tarefas`/`tarefas_pendentes`; `?busca=111.444` traz só eventos de Maria; `?tipo=desligamento` traz `total == 1`.
  - `test_eventos_404_e_aprendiz_403`: `GET /eventos/999` → 404 "Evento não encontrado."; Aprendiz → 403.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_acesso_eventos.py -q` → FAIL.
- [ ] **Step 3: Implementar** o manager, o router, `get_acesso_evento_db` e as linhas em `main.py` (`from app.routers import acesso_eventos` e `app.include_router(acesso_eventos.router)`).
- [ ] **Step 4: Rodar** `RBT tests/test_acesso_eventos.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(acessos): eventos de admissão, desligamento e mudança de função com tarefas`.

### Task 6: Resolução de tarefas, conclusão e cancelamento de eventos

**Files:**
- Modify: `backend/app/db/acesso_evento_db.py`, `backend/app/routers/acesso_eventos.py`
- Create: `backend/tests/test_acesso_tarefas.py`

**Interfaces:**
- Consumes: Task 5 (`obter`, `_tentar_concluir`), Task 4 (`criar_acesso`, `desativar_acesso`, `obter_acesso` com `cur`).
- Produces: `AcessoEventoDBManager.resolver_tarefa(ctx, tarefa_id: int, *, status: str, login: str | None, item_id: int | None, observacao: str | None) -> dict`; `concluir(ctx, evento_id: int) -> dict`; `cancelar(ctx, evento_id: int, *, motivo: str) -> dict` (todos devolvem `obter(...)`); rotas `POST /tarefas/{tarefa_id}/resolver`, `POST /eventos/{evento_id}/concluir`, `POST /eventos/{evento_id}/cancelar`.
- Algoritmo de `resolver_tarefa` (uma transação):
  1. Achar a tarefa (`NaoEncontrado("Tarefa não encontrada.")`); `SELECT ... FROM acesso_eventos WHERE id=%s FOR UPDATE`, depois reler a tarefa com `FOR UPDATE`.
  2. Evento não `aberto` → `Conflito("O evento não está aberto.")`; tarefa não `pendente` → `Conflito("Esta tarefa já foi resolvida.")`.
  3. `manter`: só para `acao='desativar'` em evento `mudanca_funcao` (senão `Invalido("Manter só vale para tarefas de desativar em mudança de função.")`); sem observação → `Invalido("Informe o motivo para manter o acesso.")`; grava `status='nao_se_aplica'`, `observacao = "Acesso mantido: " + observacao`; o acesso continua ativo.
  4. `nao_se_aplica`: sem observação → `Invalido("Informe a observação para marcar como não se aplica.")`.
  5. `feito` + `criar`: sem login → `Invalido("Informe o login criado no sistema.")`; se já existe acesso **ativo** do colaborador no sistema com esse login, só vincula (`acesso_id`); senão `criar_acesso(ctx, ..., item_id=item_id, cur=cur)` (cria ou reativa; repassa `Invalido`/`NaoEncontrado`). Grava `login` e `acesso_id` na tarefa.
  6. `feito` + `desativar`: se `acesso_id` aponta para acesso ativo → `desativar_acesso(ctx, acesso_id, cur=cur)`; já desativado → só marca a tarefa.
  7. UPDATE da tarefa (`feito_por=ctx.username`, `feito_em=agora`, `observacao` quando informada); auditoria `acesso_evento`/`resolver_tarefa` com `entidade_id=evento_id` e `depois={"tarefa_id", "sistema", "acao", "status", "login", "observacao"}`; `_tentar_concluir`.
- `concluir`: não aberto → `Conflito("O evento não está aberto.")`; com pendências → `Conflito("Não é possível concluir: " + "; ".join(pendencias) + ".")`; senão `_tentar_concluir`.
- `cancelar`: não aberto → `Conflito("O evento não está aberto.")`; grava `status='cancelado'`, `cancelado_por`, `cancelado_em`, `motivo_cancelamento`; tarefas pendentes ficam como estão; auditoria `acesso_evento`/`cancelar` com `depois={"motivo_cancelamento": motivo}`.

- [ ] **Step 1: Escrever os testes** (`integration`) em `backend/tests/test_acesso_tarefas.py` (eventos abertos pela API como na Task 5):
  - `test_criar_feito_registra_login_e_cria_acesso`: admissão com AD e E-mail; resolver AD `{"status": "feito", "login": "  maria.silva ", "item_id": item}` → tarefa `feito`, `login "maria.silva"`, `feito_por "tecnico.teste"`, `acesso_id` preenchido; `GET /ficha?colaborador_id=` traz AD `maria.silva` ativo com `item_id`; auditoria com uma linha `acesso`/`criar` e uma `acesso_evento`/`resolver_tarefa`; evento ainda `aberto`.
  - `test_ultima_tarefa_conclui_a_admissao`: resolver E-mail como `nao_se_aplica` com `"observacao": "Sem licença"` → evento `concluido`, `concluido_por "tecnico.teste"`, `pendencias == []`.
  - `test_criar_feito_sem_login_422`: → 422 "Informe o login criado no sistema.", tarefa continua `pendente`.
  - `test_criar_feito_reativa_login_desativado`: Maria tem AD `maria.silva` desativado → resolver com `"Maria.Silva"` reaproveita o mesmo id de acesso, agora ativo.
  - `test_criar_feito_em_sistema_desativado_depois_422`: AD desativado no catálogo após abrir → `feito` dá 422 "Sistema desativado: não é possível registrar acesso nele."; `nao_se_aplica` com observação funciona e o evento conclui.
  - `test_nao_se_aplica_exige_observacao`: → 422 "Informe a observação para marcar como não se aplica.".
  - `test_resolver_duas_vezes_409_e_um_so_acesso` (Review Focus): o mesmo `feito` duas vezes → segunda 409 "Esta tarefa já foi resolvida."; `contar("colaborador_acessos") == 1`.
  - `test_desativar_feito_desativa_o_acesso`: desligamento de Maria com AD ativo e sem equipamento → resolver AD `{"status": "feito"}` → acesso `desativado`, `desativado_por "tecnico.teste"`; evento `concluido`.
  - `test_desligamento_nao_conclui_com_equipamento` (aceite do spec): desligamento com AD ativo e item A com Maria → resolver AD → evento continua `aberto`, `pendencias == ["1 equipamento ainda com o colaborador"]`, `pode_concluir False`; `POST /concluir` → 409 "Não é possível concluir: 1 equipamento ainda com o colaborador."; depois de `executar_sql("UPDATE items SET status='Disponível', colaborador_id=NULL, cpf=NULL WHERE id=%s", (A,))`, `POST /concluir` → 200 `concluido`.
  - `test_manter_vira_nao_se_aplica_com_observacao`: mudança com tarefa "desativar Promax" → `{"status": "manter", "observacao": "continua atendendo a rota antiga"}` → tarefa `nao_se_aplica`, `observacao "Acesso mantido: continua atendendo a rota antiga"`; o acesso Promax segue `ativo`.
  - `test_manter_fora_de_desativar_em_mudanca_422`: `manter` numa tarefa "criar" → 422 "Manter só vale para tarefas de desativar em mudança de função."; sem observação → 422 "Informe o motivo para manter o acesso.".
  - `test_cancelar_evento`: `{"motivo_cancelamento": "  "}` → 422 `detail == "Motivo do cancelamento: Informe o motivo do cancelamento."`; com `{"motivo_cancelamento": "Aberto por engano"}` → `cancelado`, `cancelado_por "tecnico.teste"`; resolver tarefa dele → 409 "O evento não está aberto."; cancelar de novo → 409; um novo evento para Maria é aceito (201).
  - `test_tarefa_inexistente_404_e_aprendiz_403`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_acesso_tarefas.py -q` → FAIL.
- [ ] **Step 3: Implementar** os três métodos e as rotas.
- [ ] **Step 4: Rodar** `RBT -q -m "not mudanca_esperada"` → suíte inteira verde.
- [ ] **Step 5: Commit** `feat(acessos): resolução de tarefas, conclusão e cancelamento de eventos`.

### Task 7: Cliente da API, rótulos e aviso de senha (frontend)

**Files:**
- Create: `frontend/src/api/acessos.ts`, `frontend/src/api/acessos.test.ts`, `frontend/src/components/acessos/rotulos.ts`, `frontend/src/components/acessos/rotulos.test.ts`, `frontend/src/components/acessos/AvisoSemSenha.tsx`

**Interfaces:**
- Consumes: Contrato da API (MSW); `api` e `Paginated` de `api/client.ts`; `StatusTone` de `lib/status.ts`.
- Produces:
  - Tipos espelhando os schemas: `TipoEvento`, `StatusEvento`, `AcaoTarefa`, `StatusTarefa`, `StatusAcesso`; `Sistema { id; nome; descricao: string | null; ativo: boolean }`; `Pacote { cargo: string; sistemas: Sistema[] }`; `Acesso` (campos de `AcessoResponse`, datas `string`); `EquipamentoResumo`; `Tarefa`; `EventoResumo`; `EventoDetalhe extends EventoResumo`.
  - Funções (cada uma devolve `res.data`): `listSistemas(params?: { include_inactive?: boolean }): Promise<Sistema[]>`; `createSistema(data: { nome: string; descricao: string | null }): Promise<Sistema>`; `updateSistema(id: number, data: { nome?: string; descricao?: string | null; ativo?: boolean }): Promise<Sistema>`; `listPacotes(): Promise<Pacote[]>`; `savePacote(data: { cargo: string; sistema_ids: number[] }): Promise<Pacote>`; `listAcessos(params: { colaborador_id?: number; item_id?: number; status?: StatusAcesso }): Promise<Acesso[]>`; `createAcesso(data: { colaborador_id: number; sistema_id: number; login: string; item_id: number | null }): Promise<Acesso>`; `updateAcesso(id: number, data: { login?: string; item_id?: number | null }): Promise<Acesso>`; `desativarAcesso(id: number): Promise<Acesso>`; `listEquipamentosDoColaborador(colaboradorId: number): Promise<EquipamentoResumo[]>`; `listEventos(params: { status?: StatusEvento; tipo?: TipoEvento; colaborador_id?: number; busca?: string; limit?: number; offset?: number }): Promise<Paginated<EventoResumo>>`; `getEvento(id: number): Promise<EventoDetalhe>`; `createEvento(data: { colaborador_id: number; tipo: TipoEvento; cargo_anterior: string | null; cargo_novo: string | null; observacao: string | null }): Promise<EventoDetalhe>`; `resolverTarefa(id: number, data: { status: 'feito' | 'nao_se_aplica' | 'manter'; login?: string; item_id?: number | null; observacao?: string }): Promise<EventoDetalhe>`; `concluirEvento(id: number): Promise<EventoDetalhe>`; `cancelarEvento(id: number, data: { motivo_cancelamento: string }): Promise<EventoDetalhe>`.
  - `rotulos.ts`: `ROTULO_TIPO_EVENTO: Record<TipoEvento, string>` = Admissão / Desligamento / Mudança de função; `STATUS_EVENTO: Record<StatusEvento, { label: string; tone: StatusTone }>` = aberto `Aberto`/`pending`, concluido `Concluído`/`available`, cancelado `Cancelado`/`retired`; `STATUS_TAREFA` = pendente `Pendente`/`pending`, feito `Feito`/`available`, nao_se_aplica `Não se aplica`/`neutral`; `ROTULO_ACAO_TAREFA` = criar `Criar`, desativar `Desativar`; `STATUS_ACESSO` = ativo `Ativo`/`available`, desativado `Desativado`/`retired`; `descreverEquipamento(e: { id: number; tipo?: string | null; brand?: string | null; model?: string | null; identificador?: string | null }): string` (mesmo formato de `descrever_item`; tudo vazio → `Equipamento #<id>`). Os tons vão sempre pelo `tone` do `StatusPill` (não mexer em `lib/status.ts`).
  - `AvisoSemSenha.tsx`: `export const AVISO_SEM_SENHA = 'O Patrimônio nunca pede nem guarda senhas: registre só o login. Não digite senhas em nenhum campo.'` e `AvisoSemSenha({ className }: { className?: string })` — `<p role="note">` com ícone `ShieldAlert` (`aria-hidden`) e o texto, fundo `bg-brand-soft`, texto `text-brand-soft-foreground`.
- Query keys usados nas próximas tasks (sempre com prefixo `'acessos'`, invalidado inteiro após qualquer mutação): `['acessos', 'sistemas', { include_inactive }]`, `['acessos', 'pacotes']`, `['acessos', 'ficha', params]`, `['acessos', 'equipamentos', colaboradorId]`, `['acessos', 'eventos', params]`, `['acessos', 'evento', id]`.

- [ ] **Step 1: Escrever os testes:** `api/acessos.test.ts` (MSW) — `createAcesso({ colaborador_id: 7, sistema_id: 1, login: 'maria.silva', item_id: null })` faz `POST /api/acessos/ficha` com corpo exatamente igual ao objeto; `listEventos({ status: 'aberto', limit: 25, offset: 0 })` chama `/api/acessos/eventos?status=aberto&limit=25&offset=0`; `resolverTarefa(5, { status: 'feito', login: 'a', item_id: 11 })` faz `POST /api/acessos/tarefas/5/resolver`; `savePacote` usa `PUT /api/acessos/pacotes`; `desativarAcesso(3)` faz `POST /api/acessos/ficha/3/desativar`. `rotulos.test.ts` — `descreverEquipamento({ id: 4, tipo: 'Notebook', brand: 'Dell', model: 'X', identificador: 'NB-1' }) === 'Notebook Dell X · NB-1'`; `descreverEquipamento({ id: 4 }) === 'Equipamento #4'`; `STATUS_EVENTO.concluido.tone === 'available'`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/api/acessos.test.ts src/components/acessos` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** os mesmos testes e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(acessos): cliente da API e rótulos do módulo de acessos`.

### Task 8: Tela de catálogo de sistemas e grupo "Acessos" no menu

**Files:**
- Create: `frontend/src/components/acessos/SistemaFormDialog.tsx`, `frontend/src/pages/acessos/SistemasPage.tsx`, `frontend/src/pages/acessos/SistemasPage.test.tsx`
- Modify: `frontend/src/App.tsx` (linha de import nova depois da última e rota `acessos/sistemas`, só Gestor, antes do comentário `{/* Fallback */}`), `frontend/src/components/layout/navigation.ts` (novo grupo no fim de `navGroups` e uma linha de import **nova** `import { AppWindow } from 'lucide-react'` — não editar a linha de import existente; as Tasks 9 e 10 acrescentam `Boxes` e `KeyRound` nesta linha da F8)

**Interfaces:**
- Consumes: Task 7; `PageHeader`, `Button` (`loading`), `StatusPill`, `Input`, `Label`, `Dialog*`, `EmptyState`/`ErrorState`/`LoadingState`, `toast`, `getErrorMessage`.
- Produces:
  - Grupo `{ title: 'Acessos', items: [ { to: '/acessos/sistemas', label: 'Sistemas', icon: AppWindow, roles: ['Gestor'] } ] }` (as Tasks 9 e 10 acrescentam itens nele).
  - `SistemaFormDialog({ open, onOpenChange, sistema }: { open: boolean; onOpenChange: (open: boolean) => void; sistema: Sistema | null })` — campos "Nome" (obrigatório) e "Descrição"; cria com `createSistema({ nome: nome.trim(), descricao: descricao.trim() || null })` ou edita com `updateSistema(id, { nome, descricao })`; toasts "Sistema cadastrado." / "Sistema atualizado."; erro da API em toast de erro.
  - `SistemasPage` (default export, `/acessos/sistemas`): `PageHeader` (eyebrow "Acessos", título "Sistemas", descrição "Catálogo de sistemas usados nos pacotes por função e na ficha de acessos."), botão "Novo sistema"; tabela (rolagem horizontal no celular) com Nome, Descrição, Situação (`StatusPill` com `STATUS_ACESSO[ativo ? 'ativo' : 'desativado']`), ações "Editar" e "Desativar"/"Reativar" (`updateSistema(id, { ativo })`, `Button loading` na linha em andamento, toasts "Sistema desativado." / "Sistema reativado."). Usa `listSistemas({ include_inactive: true })`. Sem paginação: o catálogo é curto e limitado.
- [ ] **Step 1: Escrever os testes** (`SistemasPage.test.tsx`, `useAuth` mockado como Gestor, MSW): lista AD (Ativo) e Promax (Desativado); "Novo sistema" → preencher "GLPI" → `POST /api/acessos/sistemas` com `{ nome: 'GLPI', descricao: null }` e toast "Sistema cadastrado."; resposta 409 mostra "Já existe um sistema com esse nome."; "Desativar" em AD envia `PUT /api/acessos/sistemas/1` com `{ ativo: false }`; `getVisibleNavItems('Técnico')` não contém `/acessos/sistemas` e `getVisibleNavItems('Gestor')` contém.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/acessos/SistemasPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(acessos): tela do catálogo de sistemas e grupo Acessos no menu`.

### Task 9: Tela de pacotes por função

**Files:**
- Create: `frontend/src/pages/acessos/PacotesPage.tsx`, `frontend/src/pages/acessos/PacotesPage.test.tsx`
- Modify: `frontend/src/App.tsx` (rota `acessos/pacotes`, só Gestor), `frontend/src/components/layout/navigation.ts` (item `{ to: '/acessos/pacotes', label: 'Pacotes por função', icon: Boxes, roles: ['Gestor'] }` depois de "Sistemas"; `Boxes` entra na linha de import da F8)

**Interfaces:**
- Consumes: `listCargos()` (F0), `listPacotes`, `listSistemas`, `savePacote` (Task 7); `Checkbox`, `Input`, `Button loading`, `toast`.
- Produces: `PacotesPage` (default export). Coluna de cargos = união de `listCargos()` com os cargos de `listPacotes()`, ordenada com `localeCompare(…, 'pt-BR')`, filtrada por um campo "Buscar cargo"; cada cargo é um botão (`aria-pressed`) com o resumo `Sem pacote` / `1 sistema` / `N sistemas` e, quando só existe no pacote, a nota "sem colaboradores com este cargo". Ao escolher um cargo: `h2` com o cargo e um `Checkbox` por sistema **ativo** (`listSistemas()`), marcados conforme o pacote; se o pacote tem sistemas desativados, o texto `Sistemas desativados neste pacote (não geram tarefas): <nomes>`; botão "Salvar pacote" (`loading`) envia `savePacote({ cargo, sistema_ids: marcados.sort((a, b) => a - b) })` e mostra "Pacote salvo.". Sem cargo escolhido: "Selecione um cargo para definir os sistemas padrão.". Duas colunas a partir de `md`, uma no celular.
- [ ] **Step 1: Escrever os testes** (`PacotesPage.test.tsx`, Gestor, MSW com `/api/colaboradores/cargos` → `['Supervisor', 'Vendedor']`, pacotes `[{ cargo: 'Vendedor', sistemas: [AD, E-mail] }, { cargo: 'Motorista', sistemas: [Promax inativo] }]`, sistemas ativos AD(1), E-mail(2), VPN(3)): aparecem Motorista, Supervisor e Vendedor nessa ordem, Motorista com "sem colaboradores com este cargo"; escolher Vendedor mostra AD e E-mail marcados e VPN desmarcado; marcar VPN e salvar envia `{ cargo: 'Vendedor', sistema_ids: [1, 2, 3] }` e mostra "Pacote salvo."; escolher Motorista mostra "Sistemas desativados neste pacote (não geram tarefas): Promax"; buscar "sup" deixa só Supervisor.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(acessos): tela de pacotes de sistemas por cargo`.

### Task 10: Lista de eventos e registro de evento

**Files:**
- Create: `frontend/src/components/acessos/EventoFormDialog.tsx`, `frontend/src/pages/acessos/AcessoEventosPage.tsx`, `frontend/src/pages/acessos/AcessoEventosPage.test.tsx`
- Modify: `frontend/src/App.tsx` (rota `acessos`, Gestor e Técnico), `frontend/src/components/layout/navigation.ts` (item `{ to: '/acessos', label: 'Eventos', icon: KeyRound, roles: ['Gestor', 'Técnico'] }` como **primeiro** item do grupo "Acessos"; `KeyRound` entra na linha de import da F8)

**Interfaces:**
- Consumes: `listEventos`, `createEvento` (Task 7); `listCargos`, `ColaboradorSelect` (F0); `Pagination`, `usePageSize`, `useDebouncedValue`, `StatusPill`, `AvisoSemSenha`.
- Produces:
  - `EventoFormDialog({ open, onOpenChange, colaboradorInicial, onCreated }: { open: boolean; onOpenChange: (open: boolean) => void; colaboradorInicial?: Colaborador | null; onCreated: (evento: EventoDetalhe) => void })`. Campos: "Tipo de evento" (`<select>` Admissão / Desligamento / Mudança de função, padrão Admissão); "Colaborador" (`ColaboradorSelect somenteAtivos={false}`, pré-preenchido com `colaboradorInicial`); admissão → "Cargo" (opções de `listCargos()`, padrão o cargo do colaborador); mudança → "Cargo anterior" (obrigatório, sem padrão) e "Cargo novo" (padrão o cargo do colaborador); desligamento → texto "Serão geradas tarefas para desativar todos os acessos ativos do colaborador, e o evento só termina quando ele não tiver mais equipamentos."; "Observação" (opcional); `AvisoSemSenha`. Envia sempre as cinco chaves: `{ colaborador_id, tipo, cargo_anterior, cargo_novo, observacao }` (vazio → `null`; desligamento manda os dois cargos `null`). "Registrar evento" (`loading`), desabilitado sem colaborador ou, na mudança, sem cargo anterior. Sucesso: toast "Evento registrado." e `onCreated(evento)`.
  - `AcessoEventosPage` (default export, `/acessos`): `PageHeader` (eyebrow "Acessos", título "Eventos de acesso", descrição "Admissões, desligamentos e mudanças de função com as tarefas de criar e desativar acessos."), botão "Registrar evento" (abre o diálogo; `onCreated` navega para `/acessos/eventos/<id>`); filtros "Situação" (Abertos — padrão —, Concluídos, Cancelados, Todos → sem `status`), "Tipo" (Todos + 3) e busca "Buscar por nome ou CPF" (debounce 300 ms); qualquer mudança de filtro ou de tamanho volta à página 1, e a página é corrigida quando passa da última (como `features/colaborador-detail/tabs/equipamentos.tsx`). Tabela: Colaborador (`upper(colaborador_nome)`), Tipo, Cargo (`cargo_anterior → cargo_novo` na mudança; senão o que existir), Situação (`StatusPill` com `STATUS_EVENTO`), Tarefas (`<resolvidas>/<total> resolvidas`), Aberto em (`formatDateTime`) e link "Abrir" (`aria-label="Abrir evento #<id>"`, `text-brand-text`). `Pagination` com `usePageSize('acessos-eventos')`; vazio → "Nenhum evento encontrado.".
- [ ] **Step 1: Escrever os testes** (`AcessoEventosPage.test.tsx`, Técnico, MSW; Maria `id 7` com `cargo: 'Supervisor'`, cargos `['Supervisor', 'Vendedor']`): primeira chamada com `status=aberto&limit=25&offset=0`; linha mostra "MARIA SILVA", "Admissão", "Aberto" e "1/3 resolvidas"; trocar "Situação" para "Todos" refaz a chamada sem `status` e com `offset=0`; com `total: 60`, "Próxima" pede `offset=25`; "Registrar evento" → tipo "Mudança de função", escolher Maria no `ColaboradorSelect` (interação como em `ColaboradorSelect.test.tsx`), cargo anterior "Vendedor" → `POST /api/acessos/eventos` com `{ colaborador_id: 7, tipo: 'mudanca_funcao', cargo_anterior: 'Vendedor', cargo_novo: 'Supervisor', observacao: null }` e a rota muda para `/acessos/eventos/12`; resposta 409 mostra "Já existe um evento de acesso aberto para este colaborador."; o diálogo mostra o `AVISO_SEM_SENHA`; `getVisibleNavItems('Técnico')` contém `/acessos` e `getVisibleNavItems('Jovem Aprendiz')` não.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(acessos): lista de eventos de acesso e registro de evento`.

### Task 11: Detalhe do evento com checklist de tarefas

**Files:**
- Create: `frontend/src/components/acessos/EquipamentoDoColaboradorSelect.tsx`, `frontend/src/components/acessos/ResolverTarefaDialog.tsx`, `frontend/src/pages/acessos/AcessoEventoDetailPage.tsx`, `frontend/src/pages/acessos/AcessoEventoDetailPage.test.tsx`
- Modify: `frontend/src/App.tsx` (rota `acessos/eventos/:id`, Gestor e Técnico)

**Interfaces:**
- Consumes: `getEvento`, `resolverTarefa`, `concluirEvento`, `cancelarEvento`, `listEquipamentosDoColaborador` (Task 7); `AvisoSemSenha`, rótulos.
- Produces:
  - `EquipamentoDoColaboradorSelect({ colaboradorId, value, onChange, id, disabled }: { colaboradorId: number; value: number | null; onChange: (itemId: number | null) => void; id?: string; disabled?: boolean })` — `<select>` nativo com "Nenhum" (`''`) e os equipamentos de `listEquipamentosDoColaborador` rotulados por `descreverEquipamento`; `value` fora da lista vira a opção `Equipamento #<id>`. Reaproveitado na Task 12.
  - `type ModoResolucao = 'feito' | 'nao_se_aplica' | 'manter'`; `ResolverTarefaDialog({ tarefa, modo, colaboradorId, onClose, onResolved }: { tarefa: Tarefa | null; modo: ModoResolucao | null; colaboradorId: number; onClose: () => void; onResolved: (evento: EventoDetalhe) => void })` — aberto quando `tarefa && modo`. `feito` + criar: "Login" (obrigatório; vazio mostra "Informe o login."), "Equipamento (opcional)" com `EquipamentoDoColaboradorSelect`, `AvisoSemSenha`, botão "Confirmar" → `{ status: 'feito', login: login.trim(), item_id }`. `feito` + desativar: texto `Confirme que o acesso "<login>" no sistema <sistema> foi desativado.` e botão "Confirmar desativação" → `{ status: 'feito' }`. `nao_se_aplica`: "Observação" obrigatória → `{ status: 'nao_se_aplica', observacao }`. `manter`: título "Manter acesso", "Motivo" obrigatório → `{ status: 'manter', observacao }`. Botões com `loading` (clique único).
  - `AcessoEventoDetailPage` (default export): link "Voltar aos eventos"; `h1` com `ROTULO_TIPO_EVENTO[tipo]`; link do colaborador (`upper(colaborador_nome)` → `/colaboradores/<id>`); cargo (`Vendedor → Supervisor` na mudança); `StatusPill` do evento; "Aberto por <usuário> em <data>", "Concluído em …" ou "Cancelado por … em … — Motivo: …"; lista `<ul aria-label="Tarefas">` com sistema, ação, `StatusPill` (`STATUS_TAREFA`), login, "Feito por … em …" e observação. Com evento aberto e tarefa pendente: criar → "Marcar como feito" e "Não se aplica"; desativar → "Marcar como desativado" e ("Manter acesso" na mudança de função | "Não se aplica" nos demais). Desligamento: seção "Equipamentos com o colaborador" (link para `/items/<id>` com `descreverEquipamento`, `StatusPill` do status, `upper(revenda)`), link "Ir para Devolução" (`/return`) quando há equipamento, ou "Nenhum equipamento com o colaborador.". Evento aberto: pendências listadas (`role="status"`), "Concluir evento" (`loading`, desabilitado sem `pode_concluir`) e "Cancelar evento" (diálogo com "Motivo do cancelamento" obrigatório). Toda resposta das mutações vira `setQueryData(['acessos', 'evento', id], evento)` + invalidação de `['acessos']`; toasts "Tarefa atualizada.", "Evento concluído.", "Evento cancelado.". 404 → "Evento não encontrado".
- [ ] **Step 1: Escrever os testes** (`AcessoEventoDetailPage.test.tsx`, Técnico, MSW):
  - admissão aberta: título "Admissão", link "MARIA SILVA" → `/colaboradores/7`, `StatusPill` "Aberto", tarefas AD e E-mail "Pendente".
  - "Marcar como feito" em AD: o diálogo mostra `AVISO_SEM_SENHA` e `queryByLabelText(/senha/i)` é `null`; confirmar sem login mostra "Informe o login."; com login `maria.silva` e o equipamento 11 → corpo `{ status: 'feito', login: 'maria.silva', item_id: 11 }`; a tarefa passa a "Feito".
  - dois cliques rápidos em "Confirmar" geram **uma** requisição (Review Focus).
  - "Não se aplica" → corpo `{ status: 'nao_se_aplica', observacao: 'Sem licença disponível' }`.
  - mudança de função: tarefa desativar mostra "Manter acesso" (e não "Não se aplica") → corpo `{ status: 'manter', observacao: 'continua na rota' }`.
  - desligamento: "Marcar como desativado" → corpo `{ status: 'feito' }`; seção de equipamentos com link para `/items/11` e "Ir para Devolução" com `href="/return"`; pendência "1 equipamento ainda com o colaborador" e "Concluir evento" desabilitado.
  - `pode_concluir: true` → "Concluir evento" faz `POST /api/acessos/eventos/5/concluir`.
  - "Cancelar evento" com motivo "Aberto por engano" → corpo `{ motivo_cancelamento: 'Aberto por engano' }`.
  - evento concluído não mostra botões de ação; 404 mostra "Evento não encontrado".
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(acessos): detalhe do evento com checklist de tarefas`.

### Task 12: Aba "Acessos" no detalhe do colaborador

**Files:**
- Create: `frontend/src/components/acessos/AcessoFormDialog.tsx`, `frontend/src/features/colaborador-detail/tabs/acessos.tsx`, `frontend/src/features/colaborador-detail/tabs/acessos.test.tsx` (o registro já ignora `*.test.tsx`)
- Modify: `frontend/src/pages/ColaboradorDetailPage.test.tsx:81` (só esta linha)

**Interfaces:**
- Consumes: `listAcessos`, `createAcesso`, `updateAcesso`, `desativarAcesso`, `listSistemas`, `listEventos` (Task 7); `EquipamentoDoColaboradorSelect` (Task 11); `EventoFormDialog` (Task 10); `ColaboradorDetailTab` (F0); `AlertDialog*`.
- Produces:
  - `AcessoFormDialog({ open, onOpenChange, colaboradorId, acesso }: { open: boolean; onOpenChange: (open: boolean) => void; colaboradorId: number; acesso: Acesso | null })` — novo: "Sistema" (`<select>` com `listSistemas()` ativos), "Login", "Equipamento (opcional)", `AvisoSemSenha`; envia `createAcesso({ colaborador_id, sistema_id, login: login.trim(), item_id })`. Edição: sistema só leitura; envia `updateAcesso(id, { login: login.trim(), item_id })`. Toasts "Acesso registrado." / "Acesso atualizado.".
  - Aba default export `{ id: 'acessos', label: 'Acessos', order: 50, roles: ['Gestor', 'Técnico'], Component }`. Seção "Ficha de acessos": botão "Adicionar acesso"; lista de `listAcessos({ colaborador_id })` com Sistema, Login, `StatusPill` (`STATUS_ACESSO`), Equipamento (link `/items/<id>` com `item_descricao` ou "—"), "Criado por … em …", "Desativado por … em …" ou "—", e ações "Editar" e (só ativos) "Desativar" (confirmação `Desativar o acesso "<login>" no sistema <sistema>?` → `desativarAcesso`, toast "Acesso desativado."); vazio → "Nenhum acesso registrado para este colaborador.". Seção "Eventos de acesso": `listEventos({ colaborador_id, limit: 20, offset: 0 })` com tipo, `StatusPill`, data e link para `/acessos/eventos/<id>`; botão "Registrar evento" abre `EventoFormDialog` com `colaboradorInicial={colaborador}` e navega para o evento criado. Sem paginação: a ficha de um colaborador é limitada ao número de sistemas.
  - Linha 81 de `ColaboradorDetailPage.test.tsx` passa a ser exatamente `    expect(tabs.map((t) => t.textContent)).toEqual(expect.arrayContaining(['Dados', 'Equipamentos']))` (forma neutra que as outras fases que criam abas devem usar igual, para o merge não conflitar).
- [ ] **Step 1: Escrever os testes** (`tabs/acessos.test.tsx`, renderizando `ColaboradorDetailPage` em `/colaboradores/7?aba=acessos` como no teste da F0, Técnico, MSW): a aba "Acessos" vem depois de "Equipamentos"; a ficha mostra AD `maria.silva` "Ativo" com link do equipamento para `/items/11` e "Criado por tecnico.teste"; "Adicionar acesso" → escolher VPN, login `msilva` → `POST /api/acessos/ficha` com corpo exatamente `{ colaborador_id: 7, sistema_id: 3, login: 'msilva', item_id: null }` e o diálogo tem `AVISO_SEM_SENHA` e nenhum campo com rótulo `/senha/i`; "Desativar" + confirmar → `POST /api/acessos/ficha/1/desativar`; "Editar" → `PUT /api/acessos/ficha/1` com `{ login: 'm.silva', item_id: 11 }`; os eventos listam "Admissão" com link `/acessos/eventos/5`; `getColaboradorDetailTabs('Jovem Aprendiz')` não contém `acessos`.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar** a aba, o diálogo e o ajuste da linha 81.
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes (inclusive `ColaboradorDetailPage.test.tsx`).
- [ ] **Step 5: Commit** `feat(acessos): aba de acessos no detalhe do colaborador`.

### Task 13: Aba "Acessos" no detalhe do equipamento

**Files:**
- Create: `frontend/src/features/item-detail/tabs/acessos.tsx`, `frontend/src/features/item-detail/tabs/acessos.test.tsx`
- Modify: `frontend/src/pages/ItemDetailPage.test.tsx:61` (só esta linha)

**Interfaces:**
- Consumes: `listAcessos({ item_id })` (Task 7); `ItemDetailTab` (F0).
- Produces: default export `{ id: 'acessos', label: 'Acessos', order: 60, roles: ['Gestor', 'Técnico'], Component }` — lista somente leitura dos acessos instalados no equipamento: Sistema, Login, Colaborador (link `upper(colaborador_nome)` → `/colaboradores/<id>`), `StatusPill` (`STATUS_ACESSO`); vazio → "Nenhum acesso vinculado a este equipamento.". Linha 61 de `ItemDetailPage.test.tsx` passa a ser exatamente `    expect(tabs.map((t) => t.textContent)).toEqual(expect.arrayContaining(['Dados', 'Histórico']))`.
- [ ] **Step 1: Escrever os testes** (`ItemDetailPage` em `/items/5?aba=acessos`, Técnico, MSW): chama `/api/acessos/ficha?item_id=5`; mostra "AD", "maria.silva", link "MARIA SILVA" → `/colaboradores/7` e "Ativo"; lista vazia mostra a frase de vazio; a aba "Acessos" fica entre "Dados" e "Histórico"; `getItemDetailTabs('Jovem Aprendiz')` não contém `acessos`.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar** a aba e o ajuste da linha 61.
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(acessos): aba de acessos no detalhe do equipamento`.

### Task 14: Verificação final da fase

**Files:** nenhum novo (só correções que a verificação apontar).

- [ ] **Step 1: Backend completo:** `RBT -q -m "not mudanca_esperada"` → todos passam (os da F0 e os desta fase). Conferir também `RBT tests/test_migrations.py -q` (a 0801 vai e volta).
- [ ] **Step 2: Frontend completo** em `frontend/`: `npx vitest run`, `npx tsc --noEmit` e `npm run build` → verdes.
- [ ] **Step 3: Arquivos compartilhados só com acréscimos:** `git diff feat/patrimonio-melhorias -- backend/app/main.py backend/app/dependencies.py frontend/src/App.tsx frontend/src/components/layout/navigation.ts | grep -E '^-[^-]'` → vazio. As únicas linhas alteradas em arquivos da F0 são `ColaboradorDetailPage.test.tsx:81` e `ItemDetailPage.test.tsx:61`.
- [ ] **Step 4: Conferir o spec F8 item a item** (marcar cada um com o teste que o prova):
  - Item 1 — catálogo `/acessos/sistemas`, Gestor cadastra/edita/desativa → Tasks 3 e 8.
  - Item 2 — pacotes por cargo com cargos do cadastro de colaboradores → Tasks 3 e 9.
  - Item 3 — ficha com sistema, login, status, equipamento opcional, quem criou/desativou e quando; nunca pede nem guarda senha e os formulários dizem isso → Tasks 1, 2, 4, 7, 11, 12.
  - Item 4 — eventos à mão por Gestor e Técnico: admissão (tarefas "criar" do pacote; feito com login ou não se aplica), desligamento (desativa todos os ativos + equipamentos com atalho para a devolução; só conclui sem tarefas pendentes e sem equipamento), mudança de função (diferença dos pacotes; "manter" = não se aplica com observação) → Tasks 5, 6, 10, 11.
  - Item 5 — módulo próprio no menu, mesmo cadastro de colaboradores, aba no colaborador e no equipamento → Tasks 8–13.
  - Aceite: admissão gera as tarefas certas (`test_admissao_gera_tarefas_do_pacote`); desligamento não conclui com equipamento (`test_desligamento_nao_conclui_com_equipamento`); mudança gera a diferença (`test_mudanca_de_funcao_gera_a_diferenca`); nenhuma tabela ou payload aceita senha (`test_nenhuma_coluna_de_senha`, `test_campos_de_senha_recusados`, `test_senha_em_qualquer_grafia_422`).
  - Matriz: 14.1–14.12 cobertos pelos itens acima; 14.13 (revisão periódica pelo superior) sem tela nem endpoint. Seção 8: nenhuma senha de terceiros guardada.
  - Review Focus: os cinco testes listados no topo existem e passam.
- [ ] **Step 5: Commit** das correções, se houver (`fix(acessos): ...`). Sem correções, nada a commitar.
