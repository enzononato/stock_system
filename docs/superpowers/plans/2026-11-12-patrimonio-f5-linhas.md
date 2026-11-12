# Fase 5 — Linhas telefônicas — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cadastro de linhas telefônicas com número validado, vínculo opcional a um aparelho (Celular/Tablet) e a um colaborador, eventos gravados automaticamente a cada mudança, relatório de custo mensal por unidade com exportação Excel/PDF, e as abas "Linha" (equipamento) e "Linhas" (colaborador).

**Architecture:** Módulo isolado: `backend/app/routers/linhas.py` + `backend/app/db/linha_db.py` (`LinhaDBManager`, SQL cru com `get_cursor`) + `backend/app/schemas/linhas.py`; tabelas criadas pela migração `0500`. Cada escrita roda num único `get_cursor()`: UPDATE/INSERT da linha, eventos em `linha_eventos` e `audit.record(..., cur=cur)` na mesma transação. No frontend, `api/linhas.ts`, helpers em `features/linhas/`, componentes em `components/linhas/`, páginas `LinhasPage` e `LinhasRelatorioPage`, e uma aba por arquivo nos registros da F0.

**Tech Stack:** FastAPI 0.115 + PyMySQL/DBUtils + Pydantic 2.9 + MySQL 8; `app/services/exports.py` (openpyxl/fpdf2) da F0; React 18 + TS + Tailwind + TanStack Query 5 + Vitest/MSW.

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md` — seção 3 (linha "13.1–13.4 Linhas telefônicas → F5"), seção 5 "F5", seção 6 "F5", seções 7–9. Contexto comum: `C:/Users/enzo.jz/dev/patrimonio/_tools/contexto-planejamento.md`. Interfaces da F0: `docs/superpowers/plans/2026-11-12-patrimonio-f0-base.md`.

**Ambiente:** worktree `C:/Users/enzo.jz/dev/patrimonio/f5-linhas` (branch `fase/f5-linhas`, criada de `feat/patrimonio-melhorias` com a F0 integrada). Backend: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh C:/Users/enzo.jz/dev/patrimonio/f5-linhas f5 <args do pytest>` (abaixo abreviado `RBT <args>`). Frontend: em `C:/Users/enzo.jz/dev/patrimonio/f5-linhas/frontend`, `npx vitest run <arquivo>`, `npx tsc --noEmit`, `npm run build`.

## Global Constraints

- Spec 6.F5, item 1: "CRUD de linhas (`/linhas`), com número validado (DDD + 8/9 dígitos), operadora, plano, custo mensal, status, unidade e vínculos opcionais com aparelho (Celular/Tablet) e colaborador."
- Spec 6.F5, item 2: "Eventos registrados automaticamente a cada troca de aparelho ou de usuário, portabilidade, suspensão, reativação e cancelamento, com histórico na tela da linha e aba "Linha" no detalhe do aparelho."
- Spec 6.F5, item 3: "Relatório de custo mensal por unidade (só linhas ativas), com exportação Excel/PDF."
- Aceite da F5: "trocar o aparelho de uma linha gera evento; relatório soma corretamente por unidade; número duplicado é recusado."
- Modelo (spec 5): `linhas (id, numero UNIQUE, operadora, plano, custo_mensal DECIMAL(10,2), status ENUM('ativa','suspensa','cancelada'), unidade_id FK, item_id FK NULL, colaborador_id FK NULL, iccid NULL, observacoes, criada_em)`; `linha_eventos (id, linha_id FK, tipo ENUM('ativacao','troca_aparelho','troca_usuario','portabilidade','suspensao','reativacao','cancelamento','edicao'), de_valor, para_valor, ocorrido_em, por_usuario, observacao)`.
- Migrações só na faixa `0500–0599`, par `NNNN_descricao.sql` + `NNNN_descricao.down.sql`; o executor divide comandos por `;` no fim da linha (nada de `; -- comentário` nem `;` dentro de literais). Só mudanças aditivas.
- Toda escrita grava `audit.record(ctx, "linha", linha_id, acao, antes, depois, cur=cur)` com `ctx: AuditContext = Depends(get_audit_context)`.
- 403 com as mensagens padrão: `Acesso restrito a Gestor ou Técnico.` (dependência `gestor_or_tecnico`).
- Custo sempre `Decimal` no backend (nunca `float`); JSON sai como texto (`"49.90"`).
- Listagens novas usam `Pagination` + `usePageSize(tela)` (25/50/100), `StatusPill` (com `tone` explícito) e `Button loading`; unidade exibida com `upper()`; telefone com `formatTelefone`, dinheiro com `formatMoeda`. Texto azul só com `text-brand-text` (há teste-guarda contra `text-primary`).
- Arquivos compartilhados (`backend/app/main.py`, `backend/app/dependencies.py`, `frontend/src/App.tsx`, `frontend/src/components/layout/navigation.ts`): só ACRESCENTAR linhas, sem reordenar nem reformatar. Não editar componentes compartilhados da F0 (StatusPill, Pagination, Button, `lib/status.ts`, registros de abas, `schemas/validators.py`).
- Nenhuma dependência nova (backend ou frontend).
- Textos de interface, mensagens de erro e commits em português; commits `tipo(escopo): descrição` + linha em branco + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Decisões desta fase (o executor não reabre)

1. **Um aparelho, no máximo uma linha não cancelada** (ativa ou suspensa). Sem dual-SIM. Um colaborador pode ter várias linhas. A checagem trava a linha do item com `SELECT ... FOR UPDATE` antes de procurar outra linha não cancelada com o mesmo `item_id`.
2. **Status só muda por `POST /api/linhas/{id}/status`**: `ativa→suspensa` (suspensão), `suspensa|cancelada→ativa` (reativação), `ativa|suspensa→cancelada` (cancelamento). **Cancelar libera o aparelho** (`item_id` vira NULL; o evento guarda `item_id_anterior`) e mantém o colaborador. Linha cancelada não aceita PUT (reativar antes). Como `numero` é único também entre canceladas, a reativação a partir de `cancelada` é o caminho para "cancelei por engano". **Não há exclusão** (o "D" do CRUD é o cancelamento): linha e eventos ficam para o histórico.
3. Cadastro sempre nasce `ativa`, com evento `ativacao`.
4. PUT gera, na mesma transação: `troca_aparelho` (mudou `item_id`), `troca_usuario` (mudou `colaborador_id`), `portabilidade` (mudou `operadora`) e **um `edicao` por campo** mudado entre `numero`, `plano`, `custo_mensal`, `unidade_id` (gravado com `campo='unidade'` e nomes das unidades), `iccid`, `observacoes`. O `motivo` opcional do PUT vai para `observacao` de todos os eventos daquela alteração. PUT sem diferença real não gera evento nem auditoria.
5. Além do spec: `linhas.atualizada_em`; `linha_eventos.campo`, `linha_eventos.item_id` (aparelho da linha depois do evento) e `linha_eventos.item_id_anterior` (só em `troca_aparelho` e `cancelamento`). Servem à aba do aparelho (mostra linhas que já passaram por ele) e à F7 (linha do tempo). `de_valor`/`para_valor` guardam texto legível (aparelho `"Celular Samsung A15 · CEL-A"`, nome do colaborador, operadora, status, valor do campo; custo como `"49.90"`).
6. Operadora: lista fixa `("Vivo", "Claro", "TIM", "Algar", "Outra")` (coluna `VARCHAR(50)`, validação no schema; acrescentar uma operadora é mudança só de código).
7. **Permissões:** leitura (lista, detalhe, eventos, abas) para todos os papéis — o mesmo padrão do Estoque, em que o Jovem Aprendiz só consulta; escrita (criar, editar, mudar status) e a lista de aparelhos do seletor para Gestor e Técnico; relatório de custo e exportação para Gestor e Técnico (mesmo padrão de `/report`). Rota `/linhas` para todos (botões de escrita escondidos para o Jovem Aprendiz); `/linhas/relatorio` só Gestor e Técnico.
8. Datas da fase (`criada_em`, `atualizada_em`, `ocorrido_em`) em hora local do servidor (`datetime.now()`), como `history.data_operacao` (o `audit_log` continua em UTC).
9. Colaborador `desligado` não recebe linha (checado só quando o vínculo muda); `afastado` pode. A unidade da linha é independente da revenda do aparelho (sem checagem).
10. Relatório lista só unidades com pelo menos uma linha ativa (inclui unidade inativa que ainda tenha linha ativa), ordenado por nome; exportação traz uma linha final "TOTAL". A exportação da **listagem** de linhas é da F7 (spec 6.F7 item 4), não desta fase.

## Review Focus

- Número digitado de jeitos diferentes (`+55 (74) 99999-0001`, `74 99999 0001`, `74999990001`) é o mesmo número: normalizado antes de gravar e recusado como duplicado → tabela na Task 1 e `test_numero_duplicado_recusado` (máscara × dígitos) na Task 2.
- Custo digitado no padrão brasileiro (`49,90`, `1.234,56`) chega certo ao backend e as somas não perdem centavos → `parseCusto` na Task 5 e relatório com `0.05 + 49.90 + 30.10` na Task 4.
- Linha cancelada por engano / número reaproveitado → reativar funciona e cadastrar de novo o mesmo número responde "reative-a" em vez de um "duplicado" sem saída → Tasks 2 e 3.
- Aparelho baixado ou colaborador desligado **depois** do vínculo, e linhas sem aparelho/colaborador/plano → lista, detalhe e abas continuam abrindo e mostram "(baixado)"/"(desligado)" → Task 2 (backend) e Task 8 (frontend).
- Filtro trocado na página 3, tamanho de página trocado, total que encolhe → volta à página 1 (ou à última válida) sem requisição com página fora do intervalo → Task 8.

---

### Task 1: Migração 0500 e schemas de linhas

**Files:**
- Create: `backend/migrations/0500_linhas.sql`, `backend/migrations/0500_linhas.down.sql`, `backend/app/schemas/linhas.py`, `backend/tests/test_linhas_schemas.py`
- Modify: `backend/migrations/README.md` (acrescentar a linha `| 0500-0599 | Fase 5 (linhas telefônicas) |` na tabela de faixas)

**Interfaces:**
- Consumes: `blank_to_none` de `app/schemas/validators.py` (só import).
- Produces:
  - Tabelas (exatamente estas colunas; `CHARACTER SET utf8mb4`; FKs nomeadas):
    ```sql
    CREATE TABLE IF NOT EXISTS linhas (
        id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        numero VARCHAR(11) NOT NULL,
        operadora VARCHAR(50) NOT NULL,
        plano VARCHAR(150) NULL,
        custo_mensal DECIMAL(10,2) NOT NULL,
        status ENUM('ativa','suspensa','cancelada') NOT NULL DEFAULT 'ativa',
        unidade_id INT NOT NULL,
        item_id INT NULL,
        colaborador_id INT NULL,
        iccid VARCHAR(22) NULL,
        observacoes VARCHAR(500) NULL,
        criada_em DATETIME NOT NULL,
        atualizada_em DATETIME NOT NULL,
        UNIQUE KEY uq_linhas_numero (numero),
        KEY idx_linhas_status_unidade (status, unidade_id),
        CONSTRAINT fk_linhas_unidade FOREIGN KEY (unidade_id) REFERENCES unidades(id),
        CONSTRAINT fk_linhas_item FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE SET NULL,
        CONSTRAINT fk_linhas_colaborador FOREIGN KEY (colaborador_id) REFERENCES colaboradores(id) ON DELETE SET NULL
    ) CHARACTER SET utf8mb4;
    ```
    `linha_eventos (id INT PK AI, linha_id INT NOT NULL FK linhas(id) ON DELETE CASCADE, tipo ENUM('ativacao','troca_aparelho','troca_usuario','portabilidade','suspensao','reativacao','cancelamento','edicao') NOT NULL, campo VARCHAR(30) NULL, de_valor VARCHAR(500) NULL, para_valor VARCHAR(500) NULL, item_id INT NULL, item_id_anterior INT NULL, ocorrido_em DATETIME NOT NULL, por_usuario VARCHAR(100) NULL, observacao VARCHAR(255) NULL)` com índices `idx_linha_eventos_linha (linha_id, ocorrido_em)`, `idx_linha_eventos_item (item_id)`, `idx_linha_eventos_item_anterior (item_id_anterior)` (sem FK nos dois `item_id*`: são registro histórico). Down: `DROP TABLE IF EXISTS linha_eventos;` e depois `DROP TABLE IF EXISTS linhas;`.
  - `app/schemas/linhas.py`:
    - `OPERADORAS: tuple[str, ...] = ("Vivo", "Claro", "TIM", "Algar", "Outra")`; `STATUS_LINHA = ("ativa", "suspensa", "cancelada")`; `TIPOS_APARELHO = ("Celular", "Tablet")`
    - `DDDS_VALIDOS: frozenset[str]` = 11–19, 21, 22, 24, 27, 28, 31–35, 37, 38, 41–49, 51, 53, 54, 55, 61–69, 71, 73, 74, 75, 77, 79, 81–89, 91–99 (67 DDDs).
    - `normalizar_numero_linha(valor: str | None) -> str` — regras, nesta ordem: só dígitos; se sobrarem 12 ou 13 dígitos começando com `55`, remove o `55`; tamanho ≠ 10 e ≠ 11 → `ValueError("informe DDD + 8 ou 9 dígitos.")`; dois primeiros dígitos fora de `DDDS_VALIDOS` → `ValueError("DDD inexistente.")`; 11 dígitos com o terceiro ≠ `9` → `ValueError("celular com 9 dígitos deve começar com 9.")`. Vazio/None cai na primeira mensagem.
    - `formatar_numero_linha(numero: str | None) -> str` — `"74999990001"` → `"(74) 99999-0001"`; `"7436110000"` → `"(74) 3611-0000"`; outro tamanho volta como veio; None → `""`.
    - `normalizar_iccid(valor: str | None) -> str | None` — vazio → None; remove espaços; só dígitos com 18 a 22 posições, senão `ValueError("informe de 18 a 22 dígitos.")`.
    - `validar_custo(valor: object) -> Decimal` — None/"" → `ValueError("informe o valor.")`; `Decimal(str(valor).strip())` inválido ou não finito → `"valor inválido."`; `< 0` → `"não pode ser negativo."`; mais de 2 casas → `"use no máximo 2 casas decimais."`; `>= 100000000` → `"valor acima do permitido."`; devolve `quantize(Decimal("0.01"))`. Não aceita vírgula (o frontend converte).
    - `validar_operadora(valor: str | None) -> str` — compara sem caixa com `OPERADORAS` e devolve a grafia canônica; senão `ValueError("escolha Vivo, Claro, TIM, Algar ou Outra.")`.
    - `_texto(valor: str | None, maximo: int) -> str | None` — `blank_to_none` + `ValueError(f"use no máximo {maximo} caracteres.")`.
    - `LinhaCreate(numero: str, operadora: str, plano: str | None = None, custo_mensal: Decimal, unidade_id: int, item_id: int | None = None, colaborador_id: int | None = None, iccid: str | None = None, observacoes: str | None = None)` — validators `mode="before"`: numero, operadora, custo_mensal, iccid; `plano` (150), `observacoes` (500) via `_texto`.
    - `LinhaUpdate` — os mesmos campos, todos `Optional[...] = None`, mais `motivo: str | None = None` (255). Chave ausente = não muda; `null` limpa `plano`, `item_id`, `colaborador_id`, `iccid`, `observacoes`; `null` explícito em `numero`, `operadora`, `custo_mensal` ou `unidade_id` → `ValueError("não pode ficar vazio.")` (validators `mode="before"` só rodam para chaves enviadas). O router usa `model_dump(exclude_unset=True)`.
    - `LinhaStatusChange(status: str, observacao: str | None = None)` — status fora de `STATUS_LINHA` → `ValueError("escolha ativa, suspensa ou cancelada.")`; observacao até 255.
    - `LinhaResponse`: `id, numero, operadora, plano, custo_mensal: Decimal, status, unidade_id, unidade_nome: str, item_id, item_tipo, item_brand, item_model, item_identificador, item_ativo: bool | None, colaborador_id, colaborador_nome, colaborador_situacao, iccid, observacoes, criada_em: datetime, atualizada_em: datetime` (opcionais com `= None`).
    - `LinhaEventoResponse`: `id, linha_id, linha_numero: str, tipo, campo, de_valor, para_valor, item_id, item_id_anterior, ocorrido_em: datetime, por_usuario, observacao`.
    - `AparelhoOpcao`: `id, tipo, brand, model, identificador, codigo_patrimonial, revenda, status, responsavel_nome` (todos texto opcional exceto `id`, `tipo`).
    - `CustoUnidade(unidade_id: int, unidade_nome: str, quantidade: int, custo_total: Decimal)`; `RelatorioCusto(unidades: list[CustoUnidade], total_linhas: int, custo_total: Decimal, gerado_em: datetime)`.

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_linhas_schemas.py`:
  - `unit` `test_normalizar_numero_linha` (parametrizado): `"(74) 99999-0001"`, `"+55 74 99999-0001"`, `"5574999990001"`, `"74 99999 0001"` → `"74999990001"`; `"(74) 3611-0000"` → `"7436110000"`; erros: `"7499999"` → `"informe DDD + 8 ou 9 dígitos."`, `""` → mesma, `"20999990001"` e `"23999990001"` → `"DDD inexistente."`, `"74899990001"` → `"celular com 9 dígitos deve começar com 9."`.
  - `unit` `test_formatar_numero_linha` (os três casos da interface + None).
  - `unit` `test_validar_custo`: `"49.9"` → `Decimal("49.90")`, `0` → `Decimal("0.00")`, `49.9` (float) → `Decimal("49.90")`; erros `"-1"`, `"49,90"` (valor inválido), `"1.234"` (2 casas), `"100000000"`, `None`.
  - `unit` `test_validar_operadora`: `"vivo"` → `"Vivo"`, `"tim"` → `"TIM"`, `"Oi"` → erro.
  - `unit` `test_normalizar_iccid`: `"8955 0000 0000 0000 0001"` (20 dígitos) ok; `"123"` erro; `""` → None.
  - `unit` `test_linha_update_null_em_campo_obrigatorio`: `LinhaUpdate.model_validate({"numero": None})` levanta com `"não pode ficar vazio."`; `LinhaUpdate.model_validate({"item_id": None}).model_dump(exclude_unset=True) == {"item_id": None}`; `LinhaUpdate().model_dump(exclude_unset=True) == {}`.
  - `integration` `test_migracao_0500_cria_tabelas` (usa `_test_schema`): `information_schema.columns` de `linhas` contém `numero, operadora, plano, custo_mensal, status, unidade_id, item_id, colaborador_id, iccid, observacoes, criada_em, atualizada_em`; de `linha_eventos` contém `linha_id, tipo, campo, de_valor, para_valor, item_id, item_id_anterior, ocorrido_em, por_usuario, observacao`; `COLUMN_TYPE` de `custo_mensal` é `decimal(10,2)`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_linhas_schemas.py -q` → FAIL (`ModuleNotFoundError: app.schemas.linhas`).
- [ ] **Step 3: Implementar** a migração (+ down) e `app/schemas/linhas.py` com as assinaturas e mensagens acima.
- [ ] **Step 4: Rodar** `RBT tests/test_linhas_schemas.py tests/test_migrations.py -q` → PASS (o `test_reverter_ultima_e_reaplicar` da F0 agora reverte e reaplica a 0500, exercitando o `.down.sql`).
- [ ] **Step 5: Commit**
  ```bash
  git add backend/migrations/0500_linhas.sql backend/migrations/0500_linhas.down.sql backend/migrations/README.md backend/app/schemas/linhas.py backend/tests/test_linhas_schemas.py
  git commit -m "feat(linhas): tabelas de linhas e eventos e validação do número"
  ```

### Task 2: Cadastro e consulta de linhas (manager, router, seletor de aparelhos)

**Files:**
- Create: `backend/app/db/linha_db.py`, `backend/app/routers/linhas.py`, `backend/tests/fixtures_linhas.py`, `backend/tests/test_linhas.py`
- Modify (só acréscimos): `backend/app/dependencies.py` (import de `LinhaDBManager` e `get_linha_db`), `backend/app/main.py` (linha `from app.routers import linhas` logo abaixo do import de routers existente; `app.include_router(linhas.router)` depois do último `include_router`; no fim de `_ROTULOS_DE_CAMPO`, sob o comentário `# Campos de Linhas telefônicas (F5)`, as chaves `"numero": "Número"`, `"operadora": "Operadora"`, `"plano": "Plano"`, `"custo_mensal": "Custo mensal"`, `"iccid": "ICCID"`, `"observacoes": "Observações"`, `"observacao": "Observação"`, `"motivo": "Motivo"` — pule qualquer chave que já exista)

**Interfaces:**
- Consumes: schemas da Task 1; `get_cursor` (`app/db/database_mysql.py`); `audit.record`, `AuditContext`; `get_current_user`, `gestor_or_tecnico`, `get_audit_context`, `CurrentUser` (`app/dependencies.py`); `Paginated` (`app/schemas/common.py`); tabelas `items`, `unidades`, `colaboradores` (F0).
- Produces:
  - Exceções em `linha_db.py`: `LinhaNaoEncontrada` (→ 404 `"Linha não encontrada."`), `LinhaInvalida` (→ 400, `str(e)`), `LinhaConflito` (→ 409, `str(e)`).
  - Funções de módulo (as Tasks 3 e 4 reusam):
    - `_SELECT_LINHA` — `SELECT l.*, u.nome AS unidade_nome, i.tipo AS item_tipo, i.brand AS item_brand, i.model AS item_model, i.identificador AS item_identificador, i.is_active AS item_ativo, c.nome AS colaborador_nome, c.situacao AS colaborador_situacao FROM linhas l JOIN unidades u ON u.id = l.unidade_id LEFT JOIN items i ON i.id = l.item_id LEFT JOIN colaboradores c ON c.id = l.colaborador_id`
    - `_registrar_evento(cur, linha_id: int, tipo: str, *, por_usuario: str | None, de_valor: str | None = None, para_valor: str | None = None, campo: str | None = None, item_id: int | None = None, item_id_anterior: int | None = None, observacao: str | None = None) -> None` (INSERT em `linha_eventos` com `ocorrido_em = datetime.now()`)
    - `_descrever_item(item: dict | None) -> str | None` — `"Celular Samsung A15 · CEL-A"` (tipo, brand, model separados por espaço, `" · identificador"` quando houver); None → None.
    - `_validar_unidade(cur, unidade_id: int) -> dict` — inexistente ou `is_active = 0` → `LinhaInvalida("Unidade não encontrada ou inativa.")`.
    - `_validar_aparelho(cur, item_id: int, linha_id: int | None) -> dict` — `SELECT id, tipo, brand, model, identificador, is_active FROM items WHERE id=%s FOR UPDATE`; inexistente → `LinhaInvalida("Aparelho não encontrado.")`; `is_active = 0` → `LinhaInvalida("Aparelho baixado não pode receber linha.")`; tipo fora de `TIPOS_APARELHO` → `LinhaInvalida("Só Celular ou Tablet podem receber uma linha.")`; outra linha `status <> 'cancelada'` com o mesmo `item_id` e `id <> linha_id` (consulta `FOR UPDATE`) → `LinhaConflito(f"O aparelho já está vinculado à linha {formatar_numero_linha(numero)}.")`.
    - `_validar_colaborador(cur, colaborador_id: int) -> dict` — inexistente → `LinhaInvalida("Colaborador não encontrado.")`; `situacao='desligado'` → `LinhaInvalida("Colaborador desligado não pode receber linha.")`.
    - `_checar_numero_livre(cur, numero: str, linha_id: int | None) -> None` — existente ativa/suspensa → `LinhaConflito("Já existe uma linha com este número.")`; existente cancelada → `LinhaConflito("Já existe uma linha cancelada com este número: reative-a em vez de cadastrar de novo.")`. Um `pymysql` 1062 no INSERT/UPDATE (corrida) vira `LinhaConflito("Já existe uma linha com este número.")`.
  - `class LinhaDBManager` (sem `__init__` com DDL; as tabelas vêm da migração):
    - `listar(self, *, busca: str | None = None, status: str | None = None, unidade_id: int | None = None, operadora: str | None = None, item_id: int | None = None, colaborador_id: int | None = None, limit: int = 25, offset: int = 0) -> tuple[list[dict], int]` — `busca`: `l.plano`, `l.operadora`, `c.nome`, `i.identificador` com `LIKE %busca%`, e, se a busca tiver dígitos, `l.numero`/`l.iccid` com `LIKE %digitos%`; ordem `FIELD(l.status,'ativa','suspensa','cancelada'), l.numero`.
    - `obter(self, linha_id: int) -> dict | None`
    - `criar(self, dados: dict, ctx: AuditContext) -> int` — num único `get_cursor()`: `_checar_numero_livre`, `_validar_unidade`, `_validar_aparelho` (se `item_id`), `_validar_colaborador` (se `colaborador_id`), INSERT com `status='ativa'`, `criada_em = atualizada_em = datetime.now()`, `_registrar_evento(cur, id, "ativacao", por_usuario=ctx.username, item_id=dados.get("item_id"))`, `audit.record(ctx, "linha", id, "criar", depois=<linha lida de linhas>, cur=cur)`.
    - `aparelhos(self, *, busca: str | None = None, linha_id: int | None = None, limit: int = 20) -> list[dict]` — itens `is_active = 1` com `tipo IN ('Celular','Tablet')` e **sem** outra linha não cancelada (`NOT EXISTS ... AND (l.id <> %s)` quando `linha_id` vem, para o aparelho da própria linha continuar na lista); `busca` em `brand`, `model`, `identificador`, `codigo_patrimonial`, nome do responsável; `responsavel_nome = COALESCE(c.nome, i.assigned_to)` via `LEFT JOIN colaboradores c ON c.id = i.colaborador_id`; ordem `tipo, brand, model`.
  - `app.dependencies.get_linha_db() -> LinhaDBManager` (`@lru_cache`).
  - Router `APIRouter(prefix="/api/linhas", tags=["linhas"])` — **rotas estáticas sempre declaradas acima de `/{linha_id}`** (senão `/aparelhos` cai no 422 do parâmetro inteiro):
    - `GET ""` (`get_current_user`) → `Paginated[LinhaResponse]`; query `busca`, `status`, `unidade_id`, `operadora`, `item_id`, `colaborador_id`, `limit = Query(25, ge=1, le=100)`, `offset = Query(0, ge=0)`.
    - `GET "/aparelhos"` (`gestor_or_tecnico`) → `list[AparelhoOpcao]`; query `busca`, `linha_id`, `limit = Query(20, ge=1, le=50)`.
    - `GET "/{linha_id}"` (`get_current_user`) → `LinhaResponse` ou 404.
    - `POST ""` (`gestor_or_tecnico`, `status_code=201`) → `LinhaResponse` (relida com `obter`).
  - Fixtures em `tests/fixtures_linhas.py` (importadas pelos testes com `from tests.fixtures_linhas import *  # noqa: F401,F403`; o `conftest.py` não muda): `linha_db` (`LinhaDBManager()`, depende de `limpar_banco`); `ctx_teste` (`AuditContext(usuario_id=None, username="teste", papel="Gestor", ip="127.0.0.1")`); `criar_colaborador(nome="Maria Souza", situacao="ativo") -> int` (INSERT direto via `get_cursor` em `colaboradores (cpf, nome, situacao, origem, criado_em, atualizado_em)` com CPF de 11 dígitos vindo de um contador, `origem='manual'`, devolve `lastrowid`); `criar_aparelho(tipo="Celular", brand="Samsung", model="A15", identificador=None) -> dict` (sobre `criar_item`; identificador único quando None); `nova_linha(**overrides) -> dict` (POST com `client_gestor`; padrão `{"numero": f"7499998{n:04d}", "operadora": "Vivo", "plano": "Controle 20GB", "custo_mensal": "49.90", "unidade_id": unidade["id"]}`, afirma 201); helpers `eventos_no_banco(linha_id) -> list[dict]` (ordem `id`) e `auditoria(entidade, acao) -> list[dict]` (JSON de `antes`/`depois` decodificado).

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_linhas.py` (`integration`):
  - `test_criar_linha_normaliza_numero_e_grava_ativacao`: POST `numero="+55 (74) 99999-0001"` → 201; `numero == "74999990001"`, `status == "ativa"`, `custo_mensal == "49.90"`, `unidade_nome == unidade["nome"]`; `[e["tipo"] for e in eventos_no_banco(id)] == ["ativacao"]`, com `por_usuario == usuario_gestor["username"]`.
  - `test_criar_grava_auditoria`: uma linha em `auditoria("linha", "criar")` com `entidade_id == str(id)` e `depois["numero"] == "74999990001"`.
  - `test_numero_duplicado_recusado` (aceite + Review Focus): cria `"74999990001"`; POST `"(74) 99999-0001"` → 409 `{"detail": "Já existe uma linha com este número."}`.
  - `test_numero_de_linha_cancelada_sugere_reativar`: `executar_sql("UPDATE linhas SET status='cancelada' WHERE id=%s", ...)`; POST do mesmo número → 409 `"Já existe uma linha cancelada com este número: reative-a em vez de cadastrar de novo."`.
  - `test_numero_invalido_422`: `"7499999"` → 422 `detail == "Número: informe DDD + 8 ou 9 dígitos."`; `"20999990001"` → `"Número: DDD inexistente."`.
  - `test_regras_do_aparelho`: Notebook → 400 `"Só Celular ou Tablet podem receber uma linha."`; item com `is_active=0` → 400 `"Aparelho baixado não pode receber linha."`; id 999999 → 400 `"Aparelho não encontrado."`; Tablet → 201; Celular já com linha `(74) 99999-0001` → 409 `"O aparelho já está vinculado à linha (74) 99999-0001."`.
  - `test_regras_do_colaborador_e_unidade`: colaborador `desligado` → 400 `"Colaborador desligado não pode receber linha."`; `afastado` → 201; id inexistente → 400 `"Colaborador não encontrado."`; unidade inativa (`is_active=0` via SQL) → 400 `"Unidade não encontrada ou inativa."`.
  - `test_aprendiz_le_mas_nao_escreve`: `client_aprendiz` GET lista e GET `/{id}` → 200; POST → 403 `"Acesso restrito a Gestor ou Técnico."`; GET `/aparelhos` → 403. `client_tecnico` POST → 201.
  - `test_listar_filtros_e_paginacao`: 3 linhas (`…0001` Vivo com Celular e colaborador "Maria Souza", `…0002` Vivo, `…0003` Claro suspensa via SQL) → `?status=suspensa` total 1; `?operadora=Claro` total 1; `?item_id=` total 1; `?colaborador_id=` total 1; `?busca=99999-0002` total 1; `?busca=maria` total 1; `?limit=1` → `total == 3` e 1 item; sem filtro, a suspensa vem por último.
  - `test_linha_continua_legivel_com_aparelho_baixado_e_colaborador_desligado` (Review Focus): vincula Celular e colaborador; depois `UPDATE items SET is_active=0` e `UPDATE colaboradores SET situacao='desligado'` via SQL; GET `/{id}` → 200, `item_ativo is False`, `colaborador_situacao == "desligado"`; linha sem aparelho/colaborador/plano → GET 200 com esses campos `None`.
  - `test_aparelhos_so_celular_e_tablet_livres`: Celular livre, Tablet livre, Notebook, Celular baixado e Celular com linha L → `GET /aparelhos` devolve exatamente os ids {Celular livre, Tablet livre}; com `?linha_id=L` inclui também o Celular de L; `?busca=<model do Tablet>` devolve só o Tablet; `responsavel_nome` preenchido quando o item tem `colaborador_id`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_linhas.py -q` → FAIL (404 em `/api/linhas`).
- [ ] **Step 3: Implementar** `linha_db.py`, `routers/linhas.py`, `get_linha_db`, registro em `main.py` e rótulos. O router converte as três exceções no status indicado; outros erros propagam (500).
- [ ] **Step 4: Rodar** `RBT tests/test_linhas.py tests/test_linhas_schemas.py -q` → PASS.
- [ ] **Step 5: Commit**
  ```bash
  git add backend/app/db/linha_db.py backend/app/routers/linhas.py backend/app/dependencies.py backend/app/main.py backend/tests/fixtures_linhas.py backend/tests/test_linhas.py
  git commit -m "feat(linhas): cadastro e consulta de linhas com vínculo a aparelho e colaborador"
  ```

### Task 3: Alterações com eventos automáticos (PUT, status, histórico)

**Files:**
- Modify: `backend/app/db/linha_db.py`, `backend/app/routers/linhas.py`
- Create: `backend/tests/test_linhas_eventos.py`

**Interfaces:**
- Consumes: tudo da Task 2 (`_registrar_evento`, `_validar_*`, `_checar_numero_livre`, `_descrever_item`, `_SELECT_LINHA`, exceções, fixtures).
- Produces:
  - `LinhaDBManager.atualizar(self, linha_id: int, dados: dict, ctx: AuditContext, motivo: str | None = None) -> dict` (devolve a linha relida). Algoritmo, num único `get_cursor()`:
    ```text
    atual = SELECT * FROM linhas WHERE id=%s FOR UPDATE      -> None: LinhaNaoEncontrada
    atual.status == 'cancelada'                              -> LinhaConflito("Linha cancelada: reative-a antes de alterar.")
    mudancas = {k: v for k, v in dados.items() if v != atual[k]}   (Decimal compara por valor)
    mudancas vazio                                           -> devolve obter(linha_id); sem evento, sem auditoria
    valida só o que mudou: numero (_checar_numero_livre com linha_id), unidade_id, item_id não nulo (_validar_aparelho com linha_id), colaborador_id não nulo
    UPDATE linhas SET <mudancas>, atualizada_em=now()
    item_depois = mudancas.get("item_id", atual.item_id)
    eventos, nesta ordem, todos com por_usuario=ctx.username, observacao=motivo, item_id=item_depois:
      item_id        -> troca_aparelho  de/para = _descrever_item(antigo/novo), item_id_anterior = atual.item_id
      colaborador_id -> troca_usuario   de/para = nome do colaborador antigo/novo (None quando vazio)
      operadora      -> portabilidade   de/para = operadora antiga/nova
      numero, plano, custo_mensal, unidade_id, iccid, observacoes (nesta ordem)
                     -> edicao, campo = nome do campo ("unidade" para unidade_id), de/para = str do valor
                        (custo "49.90"; unidade = nome da unidade; None continua None)
    audit.record(ctx, "linha", linha_id, "editar", *audit.diff(atual_sem_atualizada_em, nova_sem_atualizada_em), cur=cur)
    ```
  - `LinhaDBManager.alterar_status(self, linha_id: int, status: str, ctx: AuditContext, observacao: str | None = None) -> dict` — `FOR UPDATE` na linha; transições e mensagens 409: para `suspensa` exige `ativa` (`"Só linhas ativas podem ser suspensas."`); para `ativa` exige `suspensa` ou `cancelada` (`"A linha já está ativa."`); para `cancelada` exige `ativa` ou `suspensa` (`"A linha já está cancelada."`). UPDATE `status` (e `item_id = NULL` no cancelamento) + `atualizada_em`; evento `suspensao` / `reativacao` / `cancelamento` com `de_valor` = status antigo, `para_valor` = novo, `item_id` = aparelho depois da mudança, `item_id_anterior` = aparelho antigo só no cancelamento, `observacao`; auditoria com acao `suspender` / `reativar` / `cancelar`, `antes={"status": antigo}`, `depois={"status": novo}` (no cancelamento também `item_id` antes/depois).
  - `LinhaDBManager.eventos(self, linha_id: int) -> list[dict]` e `eventos_do_aparelho(self, item_id: int) -> list[dict]` — `SELECT e.*, l.numero AS linha_numero FROM linha_eventos e JOIN linhas l ON l.id = e.linha_id WHERE ...` (`e.linha_id = %s`, ou `e.item_id = %s OR e.item_id_anterior = %s`), `ORDER BY e.ocorrido_em DESC, e.id DESC LIMIT 200`. **A F7 usa `eventos_do_aparelho` como fonte da linha do tempo do equipamento.**
  - Rotas (as estáticas acima de `/{linha_id}`):
    - `GET "/eventos"` (`get_current_user`, `item_id: int = Query(...)`) → `list[LinhaEventoResponse]`.
    - `PUT "/{linha_id}"` (`gestor_or_tecnico`) → `LinhaResponse`; corpo `LinhaUpdate`; `dados = body.model_dump(exclude_unset=True)`, `motivo = dados.pop("motivo", None)`; `dados` vazio → 400 `"Nenhum campo para atualizar."`.
    - `POST "/{linha_id}/status"` (`gestor_or_tecnico`) → `LinhaResponse`; corpo `LinhaStatusChange`.
    - `GET "/{linha_id}/eventos"` (`get_current_user`) → `list[LinhaEventoResponse]`; linha inexistente → 404.

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_linhas_eventos.py` (`integration`):
  - `test_trocar_aparelho_gera_evento` (aceite): linha com Celular `Samsung A15 · CEL-A`; PUT `{"item_id": B, "motivo": "Tela quebrada"}` (B = `Motorola G54 · CEL-B`) → 200 com `item_id == B`; `GET /{id}/eventos` → `[0]` tem `tipo == "troca_aparelho"`, `de_valor == "Celular Samsung A15 · CEL-A"`, `para_valor == "Celular Motorola G54 · CEL-B"`, `item_id == B`, `item_id_anterior == A`, `observacao == "Tela quebrada"`, `por_usuario == usuario_gestor["username"]`, `linha_numero == linha["numero"]`.
  - `test_trocar_usuario_gera_evento`: "Maria Souza" → "João Lima" gera `troca_usuario` com esses nomes; depois PUT `{"colaborador_id": null}` gera `de_valor == "João Lima"`, `para_valor is None`.
  - `test_mudar_operadora_gera_portabilidade`: Vivo → Claro → evento `portabilidade` `de_valor == "Vivo"`, `para_valor == "Claro"`.
  - `test_edicao_gera_um_evento_por_campo`: PUT `{"plano": "Controle 30GB", "custo_mensal": "59.90", "unidade_id": <Unidade B>}` → três `edicao` com `campo` em `{"plano", "custo_mensal", "unidade"}`; o de custo tem `de_valor == "49.90"`, `para_valor == "59.90"`; o de unidade tem os nomes das duas unidades.
  - `test_put_sem_mudanca_nao_gera_evento_nem_auditoria`: PUT com os mesmos valores → 200; `len(eventos_no_banco)` igual a antes e `auditoria("linha", "editar") == []`. PUT `{}` → 400 `"Nenhum campo para atualizar."`.
  - `test_auditoria_editar_so_campos_alterados`: depois da troca de aparelho, `auditoria("linha", "editar")[0]` tem `antes == {"item_id": A}` e `depois == {"item_id": B}`.
  - `test_falha_ao_gravar_evento_desfaz_a_alteracao` (manager): `monkeypatch.setattr("app.db.linha_db._registrar_evento", <função que levanta RuntimeError>)`; `pytest.raises(RuntimeError)` em `linha_db.atualizar(id, {"operadora": "Claro"}, ctx_teste)`; depois `linha_db.obter(id)["operadora"] == "Vivo"` e `auditoria("linha", "editar") == []`.
  - `test_trocar_para_aparelho_ocupado_409`: outra linha criada com `numero="74999990009"` e o Celular X; PUT na primeira linha com `{"item_id": X}` → 409 `"O aparelho já está vinculado à linha (74) 99999-0009."`; a primeira linha continua sem mudança e sem evento novo.
  - `test_suspender_reativar_cancelar`: `POST /status {"status": "suspensa", "observacao": "Férias"}` → `status == "suspensa"`, evento `suspensao` `de_valor == "ativa"`, `para_valor == "suspensa"`, `observacao == "Férias"`; `"ativa"` → evento `reativacao`; `"cancelada"` → evento `cancelamento`; `auditoria("linha", "suspender")`, `("reativar")`, `("cancelar")` com uma linha cada.
  - `test_transicoes_invalidas_409`: suspender suspensa → `"Só linhas ativas podem ser suspensas."`; reativar ativa → `"A linha já está ativa."`; cancelar cancelada → `"A linha já está cancelada."`; status `"bloqueada"` → 422 com `detail.endswith("escolha ativa, suspensa ou cancelada.")` (o prefixo depende de `status` estar ou não em `_ROTULOS_DE_CAMPO`, que esta fase não altera para essa chave).
  - `test_cancelar_libera_o_aparelho_e_reativar_funciona` (Review Focus): linha L com Celular A; cancelar → `item_id is None`, evento `cancelamento` com `item_id_anterior == A`; nova linha com A → 201; PUT em L → 409 `"Linha cancelada: reative-a antes de alterar."`; reativar L → 200, `status == "ativa"`, `item_id is None`.
  - `test_eventos_do_aparelho_inclui_linha_que_saiu`: L1 criada com A e depois movida para B; `GET /eventos?item_id=A` → tipos `["troca_aparelho", "ativacao"]`, todos com `linha_numero == L1["numero"]`; `GET /eventos?item_id=B` → `["troca_aparelho"]`.
  - `test_permissoes_de_alteracao`: `client_aprendiz` PUT e POST status → 403 `"Acesso restrito a Gestor ou Técnico."`; GET `/{id}/eventos` → 200; `GET /999999/eventos` → 404 `"Linha não encontrada."`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_linhas_eventos.py -q` → FAIL (405 no PUT).
- [ ] **Step 3: Implementar** `atualizar`, `alterar_status`, `eventos`, `eventos_do_aparelho` e as quatro rotas.
- [ ] **Step 4: Rodar** `RBT tests/test_linhas_eventos.py tests/test_linhas.py -q` → PASS.
- [ ] **Step 5: Commit**
  ```bash
  git add backend/app/db/linha_db.py backend/app/routers/linhas.py backend/tests/test_linhas_eventos.py
  git commit -m "feat(linhas): eventos automáticos em troca de aparelho, usuário, portabilidade e status"
  ```

### Task 4: Relatório de custo mensal por unidade e exportação

**Files:**
- Modify: `backend/app/db/linha_db.py`, `backend/app/routers/linhas.py`
- Create: `backend/tests/test_linhas_relatorio.py`

**Interfaces:**
- Consumes: `Coluna`, `gerar_excel`, `gerar_pdf`, `resposta_arquivo` (`app/services/exports.py`, F0 Task 8); `RelatorioCusto`, `CustoUnidade` (Task 1).
- Produces:
  - `LinhaDBManager.relatorio_custo(self) -> dict` — `SELECT u.id AS unidade_id, u.nome AS unidade_nome, COUNT(*) AS quantidade, SUM(l.custo_mensal) AS custo_total FROM linhas l JOIN unidades u ON u.id = l.unidade_id WHERE l.status = 'ativa' GROUP BY u.id, u.nome ORDER BY u.nome`; devolve `{"unidades": [...], "total_linhas": soma das quantidades, "custo_total": soma em Decimal começando de Decimal("0.00"), "gerado_em": datetime.now()}`. **A F7 usa este método para "linhas ativas e custo mensal" no painel.**
  - `GET "/relatorio/custo"` (`gestor_or_tecnico`) → `RelatorioCusto`.
  - `GET "/relatorio/custo/export"` (`gestor_or_tecnico`, `formato: str = "xlsx"`) — formato fora de `{"xlsx", "pdf"}` → 422 `"Formato inválido: use xlsx ou pdf."`; colunas `[Coluna("unidade_nome", "Unidade", largura=30, maiusculas=True), Coluna("quantidade", "Linhas ativas", largura=14, tipo="numero"), Coluna("custo_total", "Custo mensal", largura=18, tipo="moeda")]`; linhas = as unidades + `{"unidade_nome": "Total", "quantidade": total_linhas, "custo_total": custo_total}`; título `"Custo mensal de linhas por unidade"`; filtros `{"Status": "Somente linhas ativas"}`; PDF com `paisagem=False`; `resposta_arquivo(conteudo, "custo_linhas", formato)`. Rotas acima de `/{linha_id}`.

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_linhas_relatorio.py` (`integration`):
  - `test_relatorio_soma_por_unidade_so_ativas` (aceite + Review Focus): `criar_unidade(nome="Unidade A")`, `"Unidade B"`, `"Unidade C"`; A: ativas `49.90`, `30.10` e `0.05`, suspensa `100.00`, cancelada `200.00`; B: ativa `15.55`; C: só uma suspensa `10.00` (status via `POST /status`). `GET /relatorio/custo` → `unidades == [{"unidade_id": A, "unidade_nome": "Unidade A", "quantidade": 3, "custo_total": "80.05"}, {"unidade_id": B, "unidade_nome": "Unidade B", "quantidade": 1, "custo_total": "15.55"}]`, `total_linhas == 4`, `custo_total == "95.60"`.
  - `test_relatorio_sem_linhas`: `unidades == []`, `total_linhas == 0`, `custo_total == "0.00"`.
  - `test_exportar_excel`: mesmos dados; `?formato=xlsx` → 200, `content-type` `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, `content-disposition` contém `custo_linhas_`; `openpyxl.load_workbook(BytesIO(resp.content)).active`: a linha cuja 1ª célula é `"UNIDADE A"` tem `2ª == 3` e `float(3ª) == pytest.approx(80.05)`; existe linha com 1ª célula `"TOTAL"` e `float(3ª) == pytest.approx(95.60)` (procurar as linhas varrendo `ws.iter_rows(values_only=True)`, sem fixar o número da linha).
  - `test_exportar_pdf`: `?formato=pdf` → `resp.content[:4] == b"%PDF"`, `content-type` `application/pdf`.
  - `test_formato_invalido_422`: `?formato=csv` → 422 `"Formato inválido: use xlsx ou pdf."`.
  - `test_relatorio_exige_gestor_ou_tecnico`: `client_aprendiz` nas duas rotas → 403; `client_tecnico` → 200.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_linhas_relatorio.py -q` → FAIL.
- [ ] **Step 3: Implementar** `relatorio_custo` e as duas rotas.
- [ ] **Step 4: Rodar** `RBT tests/test_linhas_relatorio.py tests/test_linhas_eventos.py tests/test_linhas.py tests/test_linhas_schemas.py -q` → PASS.
- [ ] **Step 5: Commit**
  ```bash
  git add backend/app/db/linha_db.py backend/app/routers/linhas.py backend/tests/test_linhas_relatorio.py
  git commit -m "feat(linhas): relatório de custo mensal por unidade com exportação Excel e PDF"
  ```

### Task 5: Cliente da API e helpers de linhas (frontend)

**Files:**
- Create: `frontend/src/api/linhas.ts`, `frontend/src/api/linhas.test.ts`, `frontend/src/features/linhas/status.ts`, `frontend/src/features/linhas/eventos.ts`, `frontend/src/features/linhas/format.ts`, `frontend/src/features/linhas/linhas.test.ts`

**Interfaces:**
- Consumes: `api`, `Paginated`, `downloadAuthenticated` (`api/client.ts`); `StatusTone` (`lib/status.ts`); `formatMoeda`, `formatTelefone`, `upper` (`lib/format.ts`); `SituacaoColaborador` (`api/colaboradores.ts`); contratos HTTP das Tasks 2–4.
- Produces (`api/linhas.ts`):
  - `type StatusLinha = 'ativa' | 'suspensa' | 'cancelada'`; `type TipoEventoLinha = 'ativacao' | 'troca_aparelho' | 'troca_usuario' | 'portabilidade' | 'suspensao' | 'reativacao' | 'cancelamento' | 'edicao'`; `const OPERADORAS = ['Vivo', 'Claro', 'TIM', 'Algar', 'Outra'] as const`.
  - `interface Linha` — espelho de `LinhaResponse` (`custo_mensal: string`, opcionais como `T | null`, `item_ativo: boolean | null`, `colaborador_situacao: SituacaoColaborador | null`); `interface LinhaEvento` — espelho de `LinhaEventoResponse`; `interface AparelhoOpcao` — espelho de `AparelhoOpcao`; `interface CustoUnidade { unidade_id: number; unidade_nome: string; quantidade: number; custo_total: string }`; `interface RelatorioCusto { unidades: CustoUnidade[]; total_linhas: number; custo_total: string; gerado_em: string }`.
  - `interface LinhaInput { numero?: string; operadora?: string; plano?: string | null; custo_mensal?: string; unidade_id?: number; item_id?: number | null; colaborador_id?: number | null; iccid?: string | null; observacoes?: string | null; motivo?: string | null }`; `interface ListLinhasParams { busca?: string; status?: StatusLinha; unidade_id?: number; operadora?: string; item_id?: number; colaborador_id?: number; limit?: number; offset?: number }`.
  - `listLinhas(params?: ListLinhasParams): Promise<Paginated<Linha>>`; `getLinha(id: number): Promise<Linha>`; `createLinha(data: LinhaInput): Promise<Linha>`; `updateLinha(id: number, data: LinhaInput): Promise<Linha>`; `alterarStatusLinha(id: number, status: StatusLinha, observacao?: string): Promise<Linha>`; `listEventosLinha(id: number): Promise<LinhaEvento[]>`; `listEventosDoAparelho(itemId: number): Promise<LinhaEvento[]>` (`GET /linhas/eventos?item_id=`); `listAparelhos(params: { busca?: string; linha_id?: number; limit?: number }): Promise<AparelhoOpcao[]>`; `getRelatorioCusto(): Promise<RelatorioCusto>`; `exportarRelatorioCusto(formato: 'xlsx' | 'pdf'): Promise<void>` → `downloadAuthenticated('/linhas/relatorio/custo/export?formato=<f>', 'custo_linhas_<AAAAMMDD>.<f>')` (data local).
- Produces (`features/linhas/`):
  - `status.ts`: `type AcaoStatusLinha = 'suspender' | 'reativar' | 'cancelar'`; `STATUS_LINHA_OPCOES: { value: StatusLinha; label: string; tone: StatusTone }[]` = ativa "Ativa" `available`, suspensa "Suspensa" `pending`, cancelada "Cancelada" `retired`; `statusLinhaInfo(s: StatusLinha): { label: string; tone: StatusTone }`; `acoesDisponiveis(s: StatusLinha): AcaoStatusLinha[]` (ativa → `['suspender', 'cancelar']`, suspensa → `['reativar', 'cancelar']`, cancelada → `['reativar']`); `STATUS_DA_ACAO: Record<AcaoStatusLinha, StatusLinha>` (`suspensa`, `ativa`, `cancelada`).
  - `eventos.ts`: `ROTULO_EVENTO: Record<TipoEventoLinha, string>` = Ativação, Troca de aparelho, Troca de usuário, Portabilidade, Suspensão, Reativação, Cancelamento, Edição; `ROTULO_CAMPO: Record<string, string>` = `numero` Número, `plano` Plano, `custo_mensal` Custo mensal, `unidade` Unidade, `iccid` ICCID, `observacoes` Observações; `descreverEvento(e: LinhaEvento): string` — ativacao `"Linha ativada"`; troca_aparelho `"<de ?? 'sem aparelho'> → <para ?? 'sem aparelho'>"`; troca_usuario idem com `'sem usuário'`; portabilidade `"<de> → <para>"`; suspensao `"Linha suspensa"`; reativacao `"Linha reativada"`; cancelamento `"Linha cancelada"`; edicao `"<ROTULO_CAMPO[campo] ?? campo>: <de> → <para>"` com `formatMoeda` para `custo_mensal`, `formatTelefone` para `numero`, `upper` para `unidade`, e `'—'` para vazio.
  - `format.ts`: `parseCusto(texto: string): string | null` — remove espaços e `R$`; com vírgula: tira os pontos e troca a vírgula por ponto; o resultado precisa casar `^\d+(\.\d{1,2})?$`; devolve com 2 casas; senão `null`. `custoParaInput(v: string | number | null | undefined): string` (`"49.90"` → `"49,90"`, vazio → `''`). `descreverAparelho(a: { tipo?: string | null; brand?: string | null; model?: string | null; identificador?: string | null } | null | undefined): string` — mesmo formato do backend (`"Celular Samsung A15 · CEL-A"`), vazio → `''`.

- [ ] **Step 1: Escrever os testes:**
  - `features/linhas/linhas.test.ts`: `parseCusto` em tabela — `'49,90'`→`'49.90'`, `'1.234,56'`→`'1234.56'`, `'R$ 49,9'`→`'49.90'`, `'49.9'`→`'49.90'`, `'49'`→`'49.00'`, `'0'`→`'0.00'`; `null` para `'49,999'`, `'-1'`, `'abc'`, `''`, `'1.234'`. `custoParaInput('49.90') === '49,90'`. `descreverAparelho({ tipo: 'Celular', brand: 'Samsung', model: 'A15', identificador: 'CEL-A' }) === 'Celular Samsung A15 · CEL-A'` e sem identificador `'Tablet Apple iPad'`. `acoesDisponiveis` nos três status. `statusLinhaInfo('suspensa')` → `{ label: 'Suspensa', tone: 'pending' }`. `descreverEvento`: troca_aparelho com `de_valor: null` → `'sem aparelho → Celular Motorola G54 · CEL-B'`; edicao `custo_mensal` `'49.90'`→`'59.90'` igual a `` `Custo mensal: ${formatMoeda('49.90')} → ${formatMoeda('59.90')}` `` (montar o esperado com `formatMoeda`: o Intl usa espaço não separável); edicao `numero` → `'Número: (74) 99999-0001 → (74) 99999-0002'`; edicao `plano` com `de_valor: null` → `'Plano: — → Controle 30GB'`.
  - `api/linhas.test.ts` (MSW): `listLinhas({ status: 'ativa', unidade_id: 2, limit: 25, offset: 0 })` envia exatamente esses parâmetros; `alterarStatusLinha(5, 'cancelada', 'Chip perdido')` faz POST `/api/linhas/5/status` com `{ status: 'cancelada', observacao: 'Chip perdido' }`; `exportarRelatorioCusto('pdf')` chama `GET /api/linhas/relatorio/custo/export?formato=pdf` (stub de `URL.createObjectURL`/`revokeObjectURL` e de `HTMLAnchorElement.prototype.click`, como em `api/client.test.ts`) e o `download` do link termina em `.pdf`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/linhas src/api/linhas.test.ts` → FAIL (módulos inexistentes).
- [ ] **Step 3: Implementar** os quatro módulos.
- [ ] **Step 4: Rodar** os mesmos testes e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit**
  ```bash
  git add frontend/src/api/linhas.ts frontend/src/api/linhas.test.ts frontend/src/features/linhas
  git commit -m "feat(linhas): cliente da API e helpers de status, eventos e custo"
  ```

### Task 6: Seletor de aparelho e formulário da linha

**Files:**
- Create: `frontend/src/components/linhas/AparelhoSelect.tsx`, `frontend/src/components/linhas/AparelhoSelect.test.tsx`, `frontend/src/components/linhas/LinhaFormDialog.tsx`, `frontend/src/components/linhas/LinhaFormDialog.test.tsx`

**Interfaces:**
- Consumes: `listAparelhos`, `createLinha`, `updateLinha`, `OPERADORAS`, `Linha`, `LinhaInput`, `AparelhoOpcao` (Task 5); `parseCusto`, `custoParaInput`, `descreverAparelho` (Task 5); `getItem` (`api/items.ts`); `listUnidades` (`api/unidades.ts`); `ColaboradorSelect` (F0); `SearchableSelect` (com `search`/`onSearchChange`/`emptyText`); `useDebouncedValue`; `Dialog*`, `Button`, `Input`, `Label`, `toast`; `getErrorMessage`; `formatTelefone`, `upper`.
- Produces:
  - `AparelhoSelect({ value, onChange, linhaId, disabled, placeholder }: { value: number | null; onChange: (a: AparelhoOpcao | null) => void; linhaId?: number; disabled?: boolean; placeholder?: string })` — mesmo desenho do `ColaboradorSelect`: busca no servidor com debounce de 300 ms (`listAparelhos({ busca, linha_id: linhaId, limit: 20 })`), rótulo `descreverAparelho(a)`, subtítulo `[upper(a.revenda), upper(a.responsavel_nome), a.codigo_patrimonial]` não vazios unidos por `' · '`, `emptyText` `"Nenhum celular ou tablet livre encontrado."`; quando `value` não está nos resultados, busca `getItem(value)` para o rótulo; botão "Limpar aparelho" (`aria-label`).
  - `LinhaFormDialog({ open, onOpenChange, linha }: { open: boolean; onOpenChange: (open: boolean) => void; linha: Linha | null })` — `linha === null` cria, senão edita. Campos: Número (`inputMode="tel"`, pré-preenchido com `formatTelefone`), Operadora (select de `OPERADORAS`), Plano, Custo mensal (`inputMode="decimal"`, pré-preenchido com `custoParaInput`), Unidade (unidades ativas; na edição inclui a unidade atual da linha mesmo se inativa), Aparelho (`AparelhoSelect` com `linhaId`), Colaborador (`ColaboradorSelect`), ICCID, Observações (textarea) e, só na edição, "Motivo da alteração (opcional)". Validação local só do custo: `parseCusto(...) === null` → mensagem `"Informe o custo no formato 0,00."` sem requisição. Criação envia todos os campos (vazios como `null`, `custo_mensal` já convertido); edição envia só as chaves que mudaram (número comparado pelos dígitos; `null` para limpar aparelho/colaborador/textos) mais `motivo` se preenchido; nada mudou → mensagem `"Nenhuma alteração para salvar."` sem requisição. Erro da API aparece dentro do diálogo (`getErrorMessage`). Sucesso: invalida `['linhas']`, toast `"Linha cadastrada."` / `"Linha atualizada."` e fecha. Botão salvar com `loading`.

- [ ] **Step 1: Escrever os testes** (MSW):
  - `AparelhoSelect.test.tsx`: digitar `A15` → requisição `GET /api/linhas/aparelhos` com `busca=A15` e `limit=20` (e `linha_id=9` quando `linhaId={9}`); escolher a opção `"Celular Samsung A15 · CEL-A"` chama `onChange` com o objeto; `value` fora da lista busca `GET /api/items/7` e mostra o rótulo do item.
  - `LinhaFormDialog.test.tsx`: criar preenchendo número `(74) 99999-0001`, operadora `Vivo`, custo `49,90`, unidade id 1 → POST cujo corpo contém `numero: '(74) 99999-0001'`, `operadora: 'Vivo'`, `custo_mensal: '49.90'`, `unidade_id: 1`, `item_id: null`, `colaborador_id: null`; resposta 409 `{ detail: 'Já existe uma linha com este número.' }` mostra esse texto no diálogo; custo `12,345` mostra `"Informe o custo no formato 0,00."` e nenhuma requisição sai (o MSW com `onUnhandledRequest: 'error'` garante); editar mudando só o plano e o motivo → PUT `/api/linhas/5` com corpo exatamente `{ plano: 'Controle 30GB', motivo: 'Upgrade' }`; limpar o colaborador → corpo `{ colaborador_id: null }`; salvar sem mudar nada mostra `"Nenhuma alteração para salvar."`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/components/linhas` → FAIL.
- [ ] **Step 3: Implementar** os dois componentes.
- [ ] **Step 4: Rodar** `npx vitest run src/components/linhas` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit**
  ```bash
  git add frontend/src/components/linhas
  git commit -m "feat(linhas): formulário da linha com seletor de celular ou tablet"
  ```

### Task 7: Histórico de eventos e mudança de status (componentes)

**Files:**
- Create: `frontend/src/components/linhas/LinhaEventosList.tsx`, `frontend/src/components/linhas/LinhaHistoricoDialog.tsx`, `frontend/src/components/linhas/LinhaStatusDialog.tsx`, `frontend/src/components/linhas/LinhaEventos.test.tsx`

**Interfaces:**
- Consumes: `listEventosLinha`, `alterarStatusLinha`, `Linha`, `LinhaEvento` (Task 5); `ROTULO_EVENTO`, `descreverEvento`, `STATUS_DA_ACAO`, `AcaoStatusLinha` (Task 5); `formatDateTime` (`lib/utils.ts`), `formatTelefone`; `Dialog*`, `Button`, `toast`, `getErrorMessage`.
- Produces:
  - `LinhaEventosList({ eventos, mostrarNumero = false }: { eventos: LinhaEvento[]; mostrarNumero?: boolean })` — lista: título `ROTULO_EVENTO[tipo]` (com `formatTelefone(linha_numero)` antes quando `mostrarNumero`), linha `descreverEvento(e)`, `por_usuario · formatDateTime(ocorrido_em)`, e `observacao` quando houver; vazia → `"Nenhum evento registrado."`.
  - `LinhaHistoricoDialog({ linha, open, onOpenChange }: { linha: Linha | null; open: boolean; onOpenChange: (open: boolean) => void })` — título `"Histórico da linha <formatTelefone(numero)>"`; query `['linhas', 'eventos', linha.id]` só com `open`; estados de carregando/erro.
  - `LinhaStatusDialog({ linha, acao, open, onOpenChange }: { linha: Linha; acao: AcaoStatusLinha; open: boolean; onOpenChange: (open: boolean) => void })` — títulos `"Suspender linha"`, `"Reativar linha"`, `"Cancelar linha"`; avisos: suspender `"A linha deixa de contar no custo mensal até ser reativada."`, cancelar `"A linha deixa de contar no custo mensal e o aparelho vinculado é liberado."`; campo "Observação (opcional)"; botões `"Voltar"` e confirmar (`"Suspender"`, `"Reativar"`, `"Cancelar linha"`, com `loading`); chama `alterarStatusLinha(linha.id, STATUS_DA_ACAO[acao], observacao || undefined)`; sucesso invalida `['linhas']`, toast `"Linha suspensa."` / `"Linha reativada."` / `"Linha cancelada."` e fecha; erro da API no diálogo.

- [ ] **Step 1: Escrever os testes** em `LinhaEventos.test.tsx` (MSW):
  - `LinhaEventosList` com eventos de troca de aparelho e de edição de custo mostra `"Troca de aparelho"`, `"Celular Samsung A15 · CEL-A → Celular Motorola G54 · CEL-B"`, a observação `"Tela quebrada"` e o texto de custo montado com `formatMoeda`; com `mostrarNumero` mostra `"(74) 99999-0001"`; lista vazia mostra `"Nenhum evento registrado."`.
  - `LinhaHistoricoDialog` aberto para a linha 5 chama `GET /api/linhas/5/eventos` e mostra o título `"Histórico da linha (74) 99999-0001"`; fechado não faz requisição.
  - `LinhaStatusDialog` com `acao="cancelar"` mostra o aviso de liberação do aparelho; digitar `"Chip perdido"` e confirmar faz POST `/api/linhas/5/status` com `{ status: 'cancelada', observacao: 'Chip perdido' }` e chama `onOpenChange(false)`; resposta 409 `{ detail: 'A linha já está cancelada.' }` mostra o texto e mantém o diálogo aberto.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/components/linhas/LinhaEventos.test.tsx` → FAIL.
- [ ] **Step 3: Implementar** os três componentes.
- [ ] **Step 4: Rodar** `npx vitest run src/components/linhas` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit**
  ```bash
  git add frontend/src/components/linhas
  git commit -m "feat(linhas): histórico de eventos e diálogo de suspensão, reativação e cancelamento"
  ```

### Task 8: Tela `/linhas`, rota e menu

**Files:**
- Create: `frontend/src/pages/LinhasPage.tsx`, `frontend/src/pages/LinhasPage.test.tsx`
- Modify (só acréscimos): `frontend/src/App.tsx` (import `LinhasPage` depois do último import de página; `<Route path="linhas" element={<LinhasPage />} />` sem `RequireRole`, logo depois do bloco da rota `locais`), `frontend/src/components/layout/navigation.ts` (nova linha `import { Smartphone } from 'lucide-react'` abaixo do import do lucide existente; no grupo `'Inventário'`, depois do item `/locais`: `{ to: '/linhas', label: 'Linhas', icon: Smartphone },` sem `roles`)

**Interfaces:**
- Consumes: `listLinhas`, `OPERADORAS`, `Linha` (Task 5); `STATUS_LINHA_OPCOES`, `statusLinhaInfo`, `acoesDisponiveis`, `descreverAparelho` (Task 5); `LinhaFormDialog` (Task 6); `LinhaHistoricoDialog`, `LinhaStatusDialog` (Task 7); `listUnidades({ include_inactive: true })`; `useAuth`; `PageHeader`, `Pagination`, `usePageSize('linhas')`, `StatusPill`, `Button`, `Input`, `EmptyState`/`ErrorState`/`LoadingState`, `useDebouncedValue`; `formatTelefone`, `formatMoeda`, `upper`.
- Produces: `LinhasPage` (default export) em `/linhas`:
  - `PageHeader` eyebrow `"Telefonia"`, título `"Linhas telefônicas"`; ações só para Gestor/Técnico: link-botão `"Relatório de custo"` → `/linhas/relatorio` e `"Nova linha"`.
  - Filtros: busca (`aria-label="Buscar linha"`, placeholder `"Buscar por número, plano, ICCID ou colaborador..."`, debounce 300 ms, valor inicial de `?busca=` da URL), Status, Unidade, Operadora (`<select>` com `aria-label` igual ao rótulo). Mudar qualquer filtro ou o tamanho da página volta à página 1 no mesmo evento; se o total encolher e a página passar da última, corrige para a última (mesmo padrão da `ColaboradoresPage`).
  - Tabela (`overflow-x-auto`): Número (`formatTelefone`, `tabular-nums`), Operadora, Plano, Custo mensal (`formatMoeda`), Status (`StatusPill` com `status={label}` e `tone` de `statusLinhaInfo`), Unidade (`upper`), Aparelho (link `text-brand-text` para `/items/:id` com `descreverAparelho`; `" (baixado)"` quando `item_ativo === false`; `—` sem aparelho), Colaborador (`upper`; link para `/colaboradores/:id` só para Gestor/Técnico; `" (desligado)"` quando `colaborador_situacao === 'desligado'`; `—` sem colaborador), Ações: `"Histórico"` (todos), `"Editar"` (Gestor/Técnico, não em cancelada) e um botão por `acoesDisponiveis(status)` (Gestor/Técnico) com rótulos `"Suspender"`, `"Reativar"`, `"Cancelar"`; cada botão com `aria-label` incluindo o número formatado (ex.: `"Editar (74) 99999-0001"`).
  - `Pagination` com `usePageSize('linhas')`.

- [ ] **Step 1: Escrever os testes** em `LinhasPage.test.tsx` (MSW; papel via `let role = 'Gestor'` + `vi.mock('@/contexts/AuthContext', ...)` como em `ColaboradoresPage.test.tsx`; `MemoryRouter`):
  - Lista: mostra `"(74) 99999-0001"`, o custo com `formatMoeda('49.90')`, `"Ativa"`, `"UNIDADE A"`, link `"Celular Samsung A15 · CEL-A"` com `href="/items/7"`.
  - `test` Review Focus: linha com `item_ativo: false` mostra `"(baixado)"`; com `colaborador_situacao: 'desligado'` mostra `"(desligado)"`; linha sem aparelho/colaborador/plano renderiza `—` sem erro.
  - Filtros: escolher Status `Suspensa` envia `status=suspensa` e `offset=0`; Unidade e Operadora enviam `unidade_id` e `operadora`; estando na página 3 (`offset=50` com 25), trocar a operadora envia `offset=0`; trocar o tamanho para 50 envia `limit=50&offset=0`; total que encolhe de 120 para 10 estando na página 3 leva a uma requisição com `offset=0`.
  - `/linhas?busca=99999` faz a primeira requisição com `busca=99999`.
  - Jovem Aprendiz: não vê `"Nova linha"`, `"Relatório de custo"`, `"Editar …"`, `"Suspender …"`; vê `"Histórico …"`; o colaborador aparece como texto, sem link.
  - Gestor: linha `ativa` tem `"Suspender …"` e `"Cancelar …"`; `cancelada` tem só `"Reativar …"` e não tem `"Editar …"`; clicar `"Histórico …"` abre o diálogo e chama `/api/linhas/<id>/eventos`.
  - `navigation`: `getVisibleNavItems('Jovem Aprendiz')` contém `to: '/linhas'`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/LinhasPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar** a página, a rota e o item de menu.
- [ ] **Step 4: Rodar** `npx vitest run src/pages/LinhasPage.test.tsx src/components/layout` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit**
  ```bash
  git add frontend/src/pages/LinhasPage.tsx frontend/src/pages/LinhasPage.test.tsx frontend/src/App.tsx frontend/src/components/layout/navigation.ts
  git commit -m "feat(linhas): tela de linhas telefônicas com filtros, paginação e ações por status"
  ```

### Task 9: Tela do relatório de custo com exportação

**Files:**
- Create: `frontend/src/pages/LinhasRelatorioPage.tsx`, `frontend/src/pages/LinhasRelatorioPage.test.tsx`
- Modify (só acréscimos): `frontend/src/App.tsx` (import `LinhasRelatorioPage`; rota `linhas/relatorio` com `<RequireRole roles={['Gestor', 'Técnico']}>`, logo depois da rota `linhas`)

**Interfaces:**
- Consumes: `getRelatorioCusto`, `exportarRelatorioCusto`, `RelatorioCusto` (Task 5); `PageHeader`, `Button loading`, `EmptyState`/`ErrorState`/`LoadingState`, `toast`, `getErrorMessage`, `formatMoeda`, `upper`, `formatDateTime`.
- Produces: `LinhasRelatorioPage` (default export) em `/linhas/relatorio`: link `"Voltar para linhas"` (`/linhas`); `PageHeader` título `"Custo mensal de linhas"`, descrição `"Somente linhas ativas, somadas por unidade."`; tabela Unidade (`upper`) · Linhas ativas · Custo mensal (`formatMoeda`), rodapé `"Total"` com `total_linhas` e `custo_total`; texto `"Gerado em <formatDateTime(gerado_em)>"`; botões `"Exportar Excel"` e `"Exportar PDF"`, cada um com seu próprio `loading`, chamando `exportarRelatorioCusto('xlsx' | 'pdf')`; falha → toast de erro com `getErrorMessage`; `unidades` vazio → `EmptyState` `"Nenhuma linha ativa"`.

- [ ] **Step 1: Escrever os testes** (MSW, stubs de download como na Task 5): com o relatório da Task 4 (`Unidade A` 3 / `"80.05"`, `Unidade B` 1 / `"15.55"`, total 4 / `"95.60"`) mostra `"UNIDADE A"`, `formatMoeda('80.05')`, `"Total"`, `formatMoeda('95.60')`; clicar `"Exportar PDF"` faz `GET /api/linhas/relatorio/custo/export?formato=pdf` e o botão fica `aria-busy="true"` enquanto a resposta não chega; `"Exportar Excel"` usa `formato=xlsx`; export com 500 mostra o toast de erro; relatório vazio mostra `"Nenhuma linha ativa"`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/LinhasRelatorioPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar** a página e a rota.
- [ ] **Step 4: Rodar** o teste e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit**
  ```bash
  git add frontend/src/pages/LinhasRelatorioPage.tsx frontend/src/pages/LinhasRelatorioPage.test.tsx frontend/src/App.tsx
  git commit -m "feat(linhas): relatório de custo mensal por unidade com exportação Excel e PDF"
  ```

### Task 10: Abas "Linha" no equipamento e "Linhas" no colaborador

**Files:**
- Create: `frontend/src/features/item-detail/tabs/linha.tsx`, `frontend/src/features/item-detail/tabs/linha.test.tsx`, `frontend/src/features/colaborador-detail/tabs/linhas.tsx`, `frontend/src/features/colaborador-detail/tabs/linhas.test.tsx`
- Modify: `frontend/src/pages/ItemDetailPage.test.tsx` e `frontend/src/pages/ColaboradorDetailPage.test.tsx` — só a asserção da lista exata de abas (ver Step 1), porque qualquer aba nova a quebra.

**Interfaces:**
- Consumes: `ItemDetailTab` / `ColaboradorDetailTab` (F0, registros por arquivo — não editar o registro); `listLinhas`, `listEventosDoAparelho` (Task 5); `statusLinhaInfo`, `descreverAparelho` (Task 5); `LinhaEventosList`, `LinhaHistoricoDialog` (Task 7); `Pagination`, `usePageSize('colaborador-linhas')`, `StatusPill`, `formatTelefone`, `formatMoeda`, `upper`.
- Produces:
  - `features/item-detail/tabs/linha.tsx`: `export default { id: 'linha', label: 'Linha', order: 50, Component }` (sem `roles`). Item com `tipo` fora de `Celular`/`Tablet` → `"Linhas telefônicas só se aplicam a celulares e tablets."` **sem requisição**. Celular/Tablet: `listLinhas({ item_id, limit: 25 })` (query `['linhas', 'aparelho', item.id]`) → cartão da linha atual (número, operadora, plano, custo, `StatusPill`, colaborador, unidade em `upper`, link `"Abrir em Linhas"` para `/linhas?busca=<numero>`) ou `"Nenhuma linha vinculada a este aparelho."`; abaixo, `"Histórico de linhas neste aparelho"` com `listEventosDoAparelho(item.id)` (query `['linhas', 'eventos-aparelho', item.id]`) em `LinhaEventosList` com `mostrarNumero`.
  - `features/colaborador-detail/tabs/linhas.tsx`: `export default { id: 'linhas', label: 'Linhas', order: 40, Component }`. `listLinhas({ colaborador_id, limit, offset })` (query `['linhas', 'colaborador', id, page, pageSize]`) paginada como a aba `equipamentos` (volta à página 1 ao trocar de colaborador e corrige página além da última); cada linha: número, `operadora · plano`, custo, `StatusPill`, aparelho com link para `/items/:id`, botão `"Histórico (74) 99999-0001"` que abre `LinhaHistoricoDialog`; vazio → `"Nenhuma linha com este colaborador."`.

- [ ] **Step 1: Ajustar as asserções da F0 e escrever os testes:**
  - Em `ItemDetailPage.test.tsx`, trocar `expect(tabs.map((t) => t.textContent)).toEqual(['Dados', 'Histórico'])` por:
    ```ts
    const rotulos = tabs.map((t) => t.textContent)
    expect(rotulos[0]).toBe('Dados')
    expect(rotulos[rotulos.length - 1]).toBe('Histórico')
    ```
    Em `ColaboradorDetailPage.test.tsx`, trocar `expect(tabs.map((t) => t.textContent)).toEqual(['Dados', 'Equipamentos'])` por `expect(tabs.map((t) => t.textContent).slice(0, 2)).toEqual(['Dados', 'Equipamentos'])`. Se a F0 já tiver relaxado essas asserções, não mexer. (Usar exatamente este texto: as outras fases que acrescentam abas devem fazer a mesma troca, e mudanças idênticas não conflitam no merge.)
  - `linha.test.tsx` (MSW, renderizando o `Component` da aba diretamente dentro de `MemoryRouter` + `renderWithClient`): Notebook mostra a mensagem e não faz requisição; Celular com linha mostra `"(74) 99999-0001"`, `"Ativa"`, link `"Abrir em Linhas"` com `href="/linhas?busca=74999990001"` e o histórico com `"Troca de aparelho"`; Celular sem linha mostra `"Nenhuma linha vinculada a este aparelho."`. Um teste com `ItemDetailPage` em `/items/7` afirma que o rótulo `'Linha'` aparece depois de `'Dados'` e antes de `'Histórico'` (por índice, sem fixar a lista inteira).
  - `linhas.test.tsx`: lista as linhas do colaborador 3 chamando `/api/linhas` com `colaborador_id=3`; mostra `"Cancelada"` numa linha cancelada; `"Histórico (74) 99999-0001"` abre o diálogo e chama `/api/linhas/5/eventos`; vazio mostra `"Nenhuma linha com este colaborador."`; troca de tamanho para 50 envia `limit=50&offset=0`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features src/pages/ItemDetailPage.test.tsx src/pages/ColaboradorDetailPage.test.tsx` → os testes novos FAIL; os dois de página passam.
- [ ] **Step 3: Implementar** as duas abas.
- [ ] **Step 4: Rodar** o mesmo comando e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit**
  ```bash
  git add frontend/src/features/item-detail/tabs/linha.tsx frontend/src/features/item-detail/tabs/linha.test.tsx frontend/src/features/colaborador-detail/tabs/linhas.tsx frontend/src/features/colaborador-detail/tabs/linhas.test.tsx frontend/src/pages/ItemDetailPage.test.tsx frontend/src/pages/ColaboradorDetailPage.test.tsx
  git commit -m "feat(linhas): abas de linha no detalhe do equipamento e do colaborador"
  ```

### Task 11: Verificação final da fase

**Files:** nenhum novo (só correções que as verificações apontarem, cada uma com seu teste).

- [ ] **Step 1: Backend inteiro:** `RBT -q -m "not mudanca_esperada"` → todos passam (linha de base da F0 + os testes novos desta fase; nenhum teste antigo alterado no backend).
- [ ] **Step 2: Migração ida e volta:** `RBT tests/test_migrations.py -q` → PASS (a 0500 é a última versão e é revertida e reaplicada).
- [ ] **Step 3: Frontend inteiro:** em `frontend/`, `npx vitest run`, `npx tsc --noEmit` e `npm run build` → verdes, incluindo o teste-guarda de `text-primary` em `lib/tokens.test.ts`.
- [ ] **Step 4: Conferir o spec item a item** (registrar o resultado no relatório da fase):
  - [ ] 13.1–13.4 / 6.F5.1 — CRUD em `/api/linhas` e tela `/linhas`; número com DDD + 8/9 dígitos (Task 1), operadora, plano, custo mensal (`DECIMAL(10,2)`), status `ativa/suspensa/cancelada`, unidade, aparelho só Celular/Tablet, colaborador opcional.
  - [ ] 6.F5.2 — eventos automáticos em troca de aparelho, troca de usuário, portabilidade, suspensão, reativação e cancelamento (e `ativacao`/`edicao`), na mesma transação; histórico na tela da linha (diálogo) e aba "Linha" no detalhe do aparelho.
  - [ ] 6.F5.3 — relatório por unidade só com ativas; exportação Excel e PDF por `app/services/exports.py`.
  - [ ] Aceite — `test_trocar_aparelho_gera_evento`, `test_relatorio_soma_por_unidade_so_ativas`, `test_numero_duplicado_recusado` passando.
  - [ ] Seção 5 — tabelas `linhas` e `linha_eventos` com as colunas do spec (mais `atualizada_em`, `campo`, `item_id`, `item_id_anterior`).
  - [ ] Auditoria `linha` em `criar`, `editar`, `suspender`, `reativar`, `cancelar`.
  - [ ] Permissões (Decisão 7): Jovem Aprendiz lê lista/abas, recebe 403 em escrita, aparelhos e relatório.
  - [ ] UI: `Pagination` 25/50/100, `StatusPill` com tom, `Button loading`, unidade em CAIXA ALTA, item de menu "Linhas", tela usável a 390 px (tabela rola na horizontal, filtros quebram linha).
  - [ ] Arquivos compartilhados só com linhas acrescentadas: `git diff feat/patrimonio-melhorias -- backend/app/main.py backend/app/dependencies.py frontend/src/App.tsx frontend/src/components/layout/navigation.ts` sem linhas removidas.
- [ ] **Step 5: Commit** (só se algum passo exigiu correção) `fix(linhas): ajustes da verificação final da fase`.

---

## Notas para a integração e outras fases

- **Asserções de abas da F0:** `ItemDetailPage.test.tsx` fixa `['Dados', 'Histórico']` e `ColaboradorDetailPage.test.tsx` fixa `['Dados', 'Equipamentos']`; toda fase que acrescenta aba (F3, F4, F5, F8) quebra esses testes. A Task 10 traz o texto exato da troca; as outras fases devem usar o mesmo texto (ou a F0 relaxa antes de as fases ramificarem).
- **F7:** `LinhaDBManager.relatorio_custo()` (linhas ativas e custo mensal do painel) e `LinhaDBManager.eventos_do_aparelho(item_id)` (fonte "linha telefônica" da linha do tempo do equipamento). A exportação da listagem de linhas fica com a F7.
- **F4 (baixa):** baixar um aparelho com linha não desvincula a linha (a F5 não toca a baixa); a tela mostra "(baixado)". Candidato a alerta da F7.
- `_ROTULOS_DE_CAMPO`: `observacoes`/`observacao`/`motivo` também podem vir da F4; na integração, manter uma só entrada (rótulos iguais).
