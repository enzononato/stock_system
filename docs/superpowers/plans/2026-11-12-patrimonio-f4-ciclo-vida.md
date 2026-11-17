# Fase 4 — Ciclo de vida e manutenção — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Dar ao equipamento um ciclo de vida completo: dados de compra com anexo da NF e local no cadastro/edição, transferência entre unidades, ordens de manutenção com custo, baixa com registro de apagamento de dados (LGPD) e uma linha do tempo única por equipamento.

**Architecture:** Cinco migrações (0401–0405) criam as colunas de compra, o status `Em Manutenção` e as tabelas `manutencoes`, `transferencias` e `baixas_apagamento`. Manutenção e transferência ganham managers próprios (`manutencao_db.py`, `transferencia_db.py`) e a manutenção um router próprio; cadastro, edição, NF, transferência, baixa e linha do tempo ficam em `routers/items.py` (dono: F4). A linha do tempo é um registro de fontes (`app/services/linha_do_tempo.py`) em que cada fonte devolve eventos de uma tabela; a F4 implementa histórico, manutenções, transferências e baixa, e a F7 liga termos, linha telefônica e inventário. No frontend, tudo novo fica em `features/ciclo-vida/` e `features/manutencoes/`, as abas são arquivos novos em `features/item-detail/tabs/` e a tela `/manutencoes` é nova.

**Tech Stack:** FastAPI 0.115 + PyMySQL/DBUtils (SQL cru via `get_cursor`) + MySQL 8.0.40 + Pydantic 2.9; React 18 + TS 5.6 + Tailwind 3.4 + TanStack Query 5 + Vitest 4 + MSW 2 + react-dropzone.

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md` — seção 6 "F4", matriz da seção 3 (4.1, 4.2, 8.4, 10.1–10.5), seção 5 "F4", seções 7–9. Contexto comum: `C:/Users/enzo.jz/dev/patrimonio/_tools/contexto-planejamento.md`. Interfaces da F0: `docs/superpowers/plans/2026-11-12-patrimonio-f0-base.md`.

**Ambiente:** worktree `WT=C:/Users/enzo.jz/dev/patrimonio/f4-ciclo-vida`, branch `fase/f4-ciclo-vida` (a partir de `feat/patrimonio-melhorias` com a F0 integrada). Backend: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh $WT f4 <args do pytest>`. Frontend: em `$WT/frontend`, `npx vitest run <arquivo>`, `npx tsc --noEmit`, `npm run build`.

## Global Constraints

- Migrações em `backend/migrations/NNNN_descricao.sql` + `NNNN_descricao.down.sql`, faixa da F4 `0400–0499`. O executor divide comandos por `;` no fim da linha: nunca `; -- comentário` na mesma linha nem `;` dentro de literais. Mudanças só aditivas.
- `items` ganha `valor_compra DECIMAL(12,2) NULL`, `data_compra DATE NULL`, `garantia_ate DATE NULL`, `nf_anexo_key VARCHAR NULL`; `items.status` ENUM ganha `'Em Manutenção'`.
- Tabelas da seção 5: `manutencoes (id, item_id FK, status ENUM('aberta','enviada','em_reparo','retornada','sem_conserto','cancelada'), defeito, assistencia, aberta_por, aberta_em, enviada_em NULL, previsao_retorno NULL, retornada_em NULL, custo DECIMAL(12,2) NULL, anexo_key NULL, observacoes)`; `transferencias (id, item_id FK, de_unidade, para_unidade, de_local_id NULL, para_local_id NULL, motivo, transferido_por, transferido_em)`; `baixas_apagamento (id, item_id FK, history_id, dados_apagados ENUM('sim','nao_aplica'), metodo ENUM('formatacao_segura','reset_fabrica','destruicao_fisica','nao_aplica'), apagado_por, apagado_em, comprovante_key NULL)`. Colunas extras decididas neste plano estão na Task 1.
- Anexo da NF: "PDF/imagem, até 10 MB, storage `notas_fiscais`".
- Transferência: `POST /items/{id}/transferir`, "Gestor e Técnico": "destino, local opcional, motivo; atualiza `revenda` e `local_id`; registra em `transferencias`, `history` (operação `Edição` com detalhe "Transferência X → Y") e auditoria".
- Manutenção: "só para equipamento sem colaborador; status do item vai para `Em Manutenção`"; "Ao retornar, o item volta a `Disponível`"; "sem conserto" sugere a baixa; "item em manutenção não aparece como disponível para empréstimo"; listagem `/manutencoes` "com filtros e total de custo por equipamento".
- Baixa: "para Notebook, Desktop, Celular, Tablet e HD, a baixa exige registrar se os dados foram apagados, o método, quem fez e quando (comprovante opcional); para os demais tipos, `nao_aplica`". "baixa de notebook sem registro de apagamento é recusada".
- Linha do tempo: `GET /items/{id}/linha-do-tempo`, em ordem cronológica.
- Toda escrita nova grava `audit.record(ctx, entidade, entidade_id, acao, antes, depois)` com `ctx: AuditContext = Depends(get_audit_context)`, depois do sucesso.
- 403 com as mensagens padrão: `Acesso restrito ao Gestor.` / `Acesso restrito a Gestor ou Técnico.`
- Horários gravados pela F4 usam `datetime.now()` (hora local do servidor, igual a `history.data_operacao`), para a linha do tempo ordenar fontes diferentes na mesma escala.
- Arquivos compartilhados (`backend/app/main.py`, `backend/app/dependencies.py`, `frontend/src/App.tsx`, `frontend/src/components/layout/navigation.ts`, `frontend/src/test/handlers.ts`): só acrescentar linhas, sem reordenar nem reformatar. Não alterar componentes da F0 (`StatusPill`, `Pagination`, `Button`, `lib/status.ts`, registros de abas).
- Telas: `Pagination` + `usePageSize` (25/50/100), `StatusPill`, `Button loading`, `upper()` para unidade/revenda; texto azul só com `text-brand-text` (nunca `text-primary`); nenhuma cor fora dos tokens.
- Textos de interface, mensagens e commits em português; commits `tipo(escopo): descrição` + linha em branco + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`. Nenhuma dependência nova.

## Review Focus

- Datas só-dia (`data_compra`, `garantia_ate`, `previsao_retorno`) vêm como `AAAA-MM-DD`; passar por `new Date()` mostra um dia a menos no fuso de Brasília → a tela converte por texto (`isoParaBR`). Testes na Task 8 (`isoParaBR('2026-01-10') === '10/01/2026'`) e na Task 9 (edição hidrata `10/01/2026`).
- Valor digitado no padrão brasileiro (`4.599,90`, `4599,9`, `4599.90`) → sempre 4599.9; negativo ou com 3 casas é recusado com mensagem clara. Testes na Task 8 (`parseMoedaBR`) e na Task 3 (422 do backend).
- Notebook roubado ou perdido não tem como ter os dados apagados → "não se aplica" vale só para os motivos `Roubo` e `Perda`; para os demais motivos continua exigindo o apagamento. Testes na Task 6.
- Equipamento "sem conserto" não pode ficar preso em `Em Manutenção` → a baixa aceita esse estado e uma nova manutenção pode ser aberta (segunda assistência). Testes nas Tasks 5 e 6.
- Upload da NF falha depois que o cadastro já foi gravado → o usuário não recadastra em dobro: vai para a edição do item com o aviso. Teste na Task 9.

---

### Task 1: Migrações da F4 e o status `Em Manutenção` nos indicadores e constantes

**Files:**
- Create: `backend/migrations/0401_items_compra.sql` (+ `.down.sql`), `backend/migrations/0402_items_status_manutencao.sql` (+ `.down.sql`), `backend/migrations/0403_manutencoes.sql` (+ `.down.sql`), `backend/migrations/0404_transferencias.sql` (+ `.down.sql`), `backend/migrations/0405_baixas_apagamento.sql` (+ `.down.sql`), `backend/tests/test_migracoes_f4.py`
- Modify: `backend/app/core/config.py` (constante nova, só acréscimo), `backend/app/db/unidade_db.py:26-28`, `backend/app/routers/constants.py`, `backend/migrations/README.md` (linha `0400-0499 | Fase 4` na tabela de faixas), `backend/tests/test_unidades.py:437-441,454-456`

**Interfaces:**
- Produces:
  - `app.core.config.ITEM_STATUSES = ["Disponível", "Indisponível", "Pendente", "Pendente Devolução", "Em Manutenção"]`
  - `app.db.unidade_db._STATUS_ITEM = tuple(ITEM_STATUSES)` (o comentário passa a dizer "as 5 chaves").
  - `GET /api/constants` ganha `"item_statuses": ITEM_STATUSES`.
  - Schema (colunas além do spec decididas aqui):
    - `items`: `valor_compra DECIMAL(12,2) NULL`, `data_compra DATE NULL`, `garantia_ate DATE NULL`, `nf_anexo_key VARCHAR(255) NULL`.
    - `manutencoes (id INT AI PK, item_id INT NOT NULL FK items(id), status ENUM(...6 valores...) NOT NULL DEFAULT 'aberta', defeito VARCHAR(500) NOT NULL, assistencia VARCHAR(150) NULL, aberta_por VARCHAR(100) NOT NULL, aberta_em DATETIME NOT NULL, enviada_em DATETIME NULL, em_reparo_em DATETIME NULL, previsao_retorno DATE NULL, retornada_em DATETIME NULL, cancelada_em DATETIME NULL, encerrada_por VARCHAR(100) NULL, custo DECIMAL(12,2) NULL, anexo_key VARCHAR(255) NULL, observacoes VARCHAR(1000) NULL)`, índices `(item_id)`, `(status)`, `(aberta_em)`. `em_reparo_em`, `cancelada_em` e `encerrada_por` existem para a linha do tempo ter data e autor de cada passo.
    - `transferencias (id INT AI PK, item_id INT NOT NULL FK items(id), de_unidade VARCHAR(100) NULL, para_unidade VARCHAR(100) NOT NULL, de_local_id INT NULL FK locais(id) ON DELETE SET NULL, para_local_id INT NULL FK locais(id) ON DELETE SET NULL, motivo VARCHAR(255) NOT NULL, transferido_por VARCHAR(100) NOT NULL, transferido_em DATETIME NOT NULL)`, índice `(item_id)`. As unidades são gravadas por nome (retrato do momento, como `history.revenda`); renomear uma unidade depois não reescreve transferências antigas.
    - `baixas_apagamento (id INT AI PK, item_id INT NOT NULL FK items(id), history_id INT NOT NULL UNIQUE FK history(id), dados_apagados ENUM('sim','nao_aplica') NOT NULL, metodo ENUM('formatacao_segura','reset_fabrica','destruicao_fisica','nao_aplica') NOT NULL, apagado_por VARCHAR(100) NULL, apagado_em DATE NULL, comprovante_key VARCHAR(255) NULL, registrado_por VARCHAR(100) NOT NULL, registrado_em DATETIME NOT NULL)`.
  - Todas as tabelas com `CHARACTER SET utf8mb4`, `CREATE TABLE IF NOT EXISTS`; downs com `DROP TABLE IF EXISTS`.
  - `0401` é idempotente coluna a coluna (MySQL 8.0 não tem `ADD COLUMN IF NOT EXISTS`). Para cada coluna, quatro comandos:
    ```sql
    SET @ddl = (SELECT IF(COUNT(*) = 0, 'ALTER TABLE items ADD COLUMN valor_compra DECIMAL(12,2) NULL', 'SELECT 1') FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'valor_compra');
    PREPARE stmt FROM @ddl;
    EXECUTE stmt;
    DEALLOCATE PREPARE stmt;
    ```
    O down usa o mesmo padrão com `COUNT(*) = 1` e `ALTER TABLE items DROP COLUMN <coluna>` (apaga os dados de compra; comentário no topo dizendo isso).
  - `0402` up: `ALTER TABLE items MODIFY status ENUM('Disponível','Indisponível','Pendente','Pendente Devolução','Em Manutenção') DEFAULT 'Disponível';` (o `CREATE TABLE` legado em `InventoryDBManager._create_tables` não muda).
  - `0402` down — **regra: a reversão recusa (falha sem alterar nada) enquanto existir item com status `Em Manutenção`**; o operador encerra/cancela as manutenções antes. Fora do modo estrito o MySQL trocaria o valor por `''` em silêncio, por isso a trava não depende do `sql_mode`:
    ```sql
    -- Trava: se houver item 'Em Manutenção', a subconsulta devolve 2 linhas e o MySQL aborta (erro 1242).
    SELECT (SELECT x FROM (SELECT 1 AS x UNION ALL SELECT 2) d WHERE EXISTS (SELECT 1 FROM items WHERE status = 'Em Manutenção')) AS trava_em_manutencao;
    ALTER TABLE items MODIFY status ENUM('Disponível','Indisponível','Pendente','Pendente Devolução') DEFAULT 'Disponível';
    ```

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_migracoes_f4.py` (`pytestmark = pytest.mark.integration`; executar arquivos com `from app.db import migrations as mig` + `with get_cursor() as cur: mig.executar_sql(cur, caminho.read_text(encoding="utf-8"))` — o nome `executar_sql` sozinho é a fixture do conftest):
  - `test_colunas_de_compra_existem`: `information_schema.columns` de `items` traz `valor_compra` `decimal(12,2)`, `data_compra` `date`, `garantia_ate` `date`, `nf_anexo_key` `varchar(255)`.
  - `test_status_aceita_em_manutencao`: `COLUMN_TYPE` de `items.status` == `"enum('Disponível','Indisponível','Pendente','Pendente Devolução','Em Manutenção')"`.
  - `test_down_do_status_recusa_com_item_em_manutencao`: `criar_item()`, `executar_sql("UPDATE items SET status='Em Manutenção' WHERE id=%s", (id,))`; rodar `0402_items_status_manutencao.down.sql` → `pytest.raises(pymysql.MySQLError)` com `args[0] == 1242`; depois o `COLUMN_TYPE` ainda contém `'Em Manutenção'` e o item continua `Em Manutenção`.
  - `test_down_e_up_do_status_sem_item_em_manutencao`: down → `COLUMN_TYPE == "enum('Disponível','Indisponível','Pendente','Pendente Devolução')"`; `finally`: up → volta a ter `'Em Manutenção'`.
  - `test_ida_e_volta_das_migracoes_f4`: com um item já cadastrado (`criar_item()`, simula o schema atual com dados), downs na ordem 0405, 0404, 0403, 0401 → as três tabelas e as quatro colunas somem; `finally`: ups na ordem 0401, 0403, 0404, 0405 rodados **duas vezes** (idempotência) → tudo existe de novo e o item continua lá, com `valor_compra` nulo.
  - Em `tests/test_unidades.py`: nos dois `assert body["itens_por_status"] == {...}` existentes, acrescentar `"Em Manutenção": 0` (comentário "MUDANÇA INTENCIONAL (F4): novo status"); novo `test_indicadores_contam_em_manutencao` (item da unidade com status forçado via `executar_sql` → `itens_por_status["Em Manutenção"] == 1`, `itens_total == 1`); novo `test_constants_traz_status_de_item` (`/api/constants` → `item_statuses == ["Disponível", "Indisponível", "Pendente", "Pendente Devolução", "Em Manutenção"]`).
- [ ] **Step 2: Rodar e ver falhar:** `run-backend-tests.sh $WT f4 tests/test_migracoes_f4.py tests/test_unidades.py -q` → FAIL (arquivos de migração inexistentes; chave `Em Manutenção` ausente).
- [ ] **Step 3: Implementar** as 10 migrações, `ITEM_STATUSES`, `_STATUS_ITEM`, a chave em `/api/constants` e a linha da faixa no README.
- [ ] **Step 4: Rodar a suíte inteira:** `run-backend-tests.sh $WT f4 -q` → tudo verde.
- [ ] **Step 5: Commit** `feat(ciclo-vida): migrações da F4 e status Em Manutenção nos indicadores`.

### Task 2: Serviço de anexos (NF, manutenção, comprovante de apagamento)

**Files:**
- Create: `backend/app/services/anexos.py`, `backend/tests/test_anexos.py`
- Modify: `backend/app/core/storage.py` (uma linha em `LocalStorage.save`)

**Interfaces:**
- Produces (`app/services/anexos.py`):
  - `MAX_ANEXO_BYTES = 10 * 1024 * 1024`
  - `CATEGORIA_NF = "notas_fiscais"`, `CATEGORIA_MANUTENCAO = "manutencoes"`, `CATEGORIA_APAGAMENTO = "apagamentos"`
  - `tipo_do_anexo(conteudo: bytes) -> str | None` — pelo conteúdo (nunca pelo nome/Content-Type): `b"%PDF"` → `"pdf"`, `b"\x89PNG\r\n\x1a\n"` → `"png"`, `b"\xff\xd8\xff"` → `"jpg"`, senão `None`.
  - `validar_anexo(conteudo: bytes) -> str` — devolve a extensão; vazio → `HTTPException(400, "O arquivo enviado está vazio.")`; `len > MAX_ANEXO_BYTES` → `HTTPException(413, "Arquivo excede o tamanho máximo de 10MB.")`; tipo `None` → `HTTPException(400, "Envie um PDF ou uma imagem PNG/JPG.")`.
  - `async def salvar_anexo(upload: UploadFile, categoria: str, prefixo: str) -> str` — lê, valida e grava como `f"{prefixo}_{AAAAMMDDHHMMSS}_{uuid4().hex[:8]}.{ext}"` (o nome do cliente é descartado); devolve a chave do storage (`"<categoria>/<arquivo>"`).
  - `resposta_anexo(key: str, nome_base: str) -> Response` — `storage.exists(key)` falso → `HTTPException(404, "Arquivo não encontrado.")`; media type pela extensão (`application/pdf`, `image/png`, `image/jpeg`); `Content-Disposition: attachment; filename="<nome_base>.<ext>"`.
  - `LocalStorage.save` passa a criar o diretório da categoria (`os.makedirs(os.path.dirname(path), exist_ok=True)`), para categorias novas não dependerem da lista do `__init__`. A F4 não toca em `routers/documents.py` (dono: F3): os downloads da F4 têm endpoints próprios.

- [ ] **Step 1: Testes** (`pytest.mark.unit`, `tests/test_anexos.py`): `test_tipo_do_anexo_pelo_conteudo` (pdf/png/jpg; `b"MZ..."` → `None`); `test_validar_anexo_vazio_400`; `test_validar_anexo_acima_de_10mb_413` (`b"%PDF" + b"0" * MAX_ANEXO_BYTES` → 413 com a mensagem exata); `test_validar_anexo_no_limite_aceito` (exatamente `MAX_ANEXO_BYTES` com cabeçalho PDF); `test_validar_anexo_executavel_400`; `test_salvar_anexo_ignora_nome_do_cliente` (`UploadFile(filename="../../x.exe", file=BytesIO(b"%PDF-1.4"))` → chave começa com `"notas_fiscais/nf_1_"`, termina com `".pdf"` e o arquivo existe em `storage.base_path`); `test_resposta_anexo_media_type` (png → `image/png`) e `test_resposta_anexo_inexistente_404`.
- [ ] **Step 2: Rodar e ver falhar:** `run-backend-tests.sh $WT f4 tests/test_anexos.py -q` → FAIL (módulo inexistente).
- [ ] **Step 3: Implementar** `anexos.py` e a linha em `storage.py`.
- [ ] **Step 4: Rodar** `tests/test_anexos.py tests/test_uploads.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(ciclo-vida): serviço de anexos com limite de 10MB e checagem do tipo pelo conteúdo`.

### Task 3: Dados de compra, local e anexo da NF no cadastro e na edição

**Files:**
- Create: `backend/app/db/transferencia_db.py` (só `erro_local` nesta task), `backend/tests/test_items_compra.py`
- Modify: `backend/app/schemas/validators.py`, `backend/app/schemas/items.py`, `backend/app/routers/items.py`, `backend/app/main.py` (`_ROTULOS_DE_CAMPO`, só chaves novas)

**Interfaces:**
- Consumes: `salvar_anexo`, `resposta_anexo`, `CATEGORIA_NF` (Task 2); `audit.record`, `audit.diff` (F0).
- Produces:
  - `validators.validar_valor_monetario(valor: Decimal | None, campo: str) -> Decimal | None` — `None` → `None`; `< 0` → `ValueError(f"Informe um {campo} maior ou igual a zero.")`; mais de 2 casas → `ValueError(f"Use no máximo duas casas decimais no {campo}.")`; `> Decimal("9999999999.99")` → `ValueError(f"O {campo} excede o limite permitido.")`; devolve `quantize(Decimal("0.01"))`.
  - `validators.validar_data_nao_futura(data: str | None, campo: str) -> str | None` — usa `validar_data_br(data, campo=campo)`; data depois de hoje → `ValueError(f"A {campo} não pode ser no futuro.")`.
  - `validators.data_br_para_date(data: str | None) -> date | None`.
  - `ItemCreate`/`ItemUpdate` ganham `valor_compra: Optional[Decimal] = None` (validador `validar_valor_monetario(v, "valor de compra")`), `data_compra: Optional[str] = None` (`validar_data_nao_futura(v, "data da compra")`), `garantia_ate: Optional[str] = None` (`validar_data_br(v, campo="fim da garantia")`), `local_id: Optional[int] = None`.
  - `ItemResponse` ganha `valor_compra: Optional[Decimal] = None`, `data_compra: Optional[date] = None`, `garantia_ate: Optional[date] = None`, `nf_anexo_key: Optional[str] = None` (JSON: `"4599.90"`, `"2026-01-10"`).
  - `transferencia_db.erro_local(local_id: int, unidade_nome: str | None, *, rotulo_unidade: str) -> str | None` — `SELECT l.ativo, u.nome FROM locais l JOIN unidades u ON u.id = l.unidade_id WHERE l.id=%s`; inexistente → `"Local não encontrado."`; inativo → `"Este local está inativo."`; outra unidade → `f"O local informado não pertence à unidade {rotulo_unidade}."`.
  - `POST /api/items`: converte as datas para `date`; `local_id` informado → `erro_local(local_id, body.revenda, rotulo_unidade="do equipamento")` → 400; `garantia_ate < data_compra` → 400 `"A garantia não pode terminar antes da data da compra."`.
  - `PUT /api/items/{id}`:
    - `revenda` presente e diferente da atual → 400 `"Para mudar a unidade do equipamento, use a transferência."` (igual à atual passa).
    - Semântica de limpeza só para `CAMPOS_ANULAVEIS = ("valor_compra", "data_compra", "garantia_ate", "local_id")`: chave ausente = não muda; `null` explícito (está em `body.model_fields_set` com valor `None`) = grava `NULL`. Os campos antigos continuam com `exclude_none=True`.
    - `local_id` só é validado quando difere do `local_id` atual (manter um local que foi desativado depois não trava a edição).
    - Garantia conferida com os valores efetivos (o enviado ou, na falta, o gravado).
  - `POST /api/items/{id}/nota-fiscal` (`gestor_or_tecnico`, multipart `arquivo: UploadFile`): item ativo senão 404 `"Item não encontrado."`; `key = await salvar_anexo(arquivo, CATEGORIA_NF, f"nf_{id}")`; `inv.update_item(id, {"nf_anexo_key": key}, username)`; auditoria `item`/`editar` com `antes={"nf_anexo_key": antigo}`, `depois={"nf_anexo_key": key}`; resposta `{"detail": "Nota fiscal anexada.", "nf_anexo_key": key}`. O arquivo antigo fica no storage.
  - `GET /api/items/{id}/nota-fiscal` (`gestor_or_tecnico`): sem anexo → 404 `"Nenhuma nota fiscal anexada a este equipamento."`; senão `resposta_anexo(key, f"nota_fiscal_item_{id}")`.
  - `_ROTULOS_DE_CAMPO` ganha: `"valor_compra": "Valor de compra"`, `"data_compra": "Data da compra"`, `"garantia_ate": "Fim da garantia"`, `"local_id": "Local"`.

- [ ] **Step 1: Testes** (`tests/test_items_compra.py`, integration; unidades via `criar_unidade(nome=...)`, locais via `client_gestor.post("/api/locais", json={"unidade_id": u["id"], "nome": "Sala TI"}).json()["id"]`):
  - `test_criar_item_com_dados_de_compra`: POST com `valor_compra: 4599.9, data_compra: "10/01/2026", garantia_ate: "10/01/2027"` → GET traz `valor_compra == "4599.90"`, `data_compra == "2026-01-10"`, `garantia_ate == "2027-01-10"`.
  - `test_valor_negativo_422`: `detail == "Valor de compra: Informe um valor de compra maior ou igual a zero."`.
  - `test_valor_com_tres_casas_422`: `4599.999` → `"Valor de compra: Use no máximo duas casas decimais no valor de compra."`.
  - `test_data_compra_futura_422`: `"Data da compra: A data da compra não pode ser no futuro."`.
  - `test_garantia_antes_da_compra_400`: `"A garantia não pode terminar antes da data da compra."`.
  - `test_criar_com_local_da_unidade` (200, GET `local_id` igual) e `test_local_de_outra_unidade_400` (`"O local informado não pertence à unidade do equipamento."`).
  - `test_editar_limpa_valor_com_null`: PUT `{"valor_compra": None}` → GET `valor_compra is None`; PUT `{"brand": "X"}` não mexe em `data_compra`.
  - `test_editar_mudando_revenda_recusado`: PUT `{"revenda": "Outra Unidade"}` → 400 com a mensagem; PUT `{"revenda": <a atual>, "brand": "Y"}` → 200.
  - `test_editar_mantendo_local_inativo_aceito`: item com local; desativar o local (`PUT /api/locais/{id}` `{"ativo": false}`); PUT `{"local_id": <o mesmo>, "brand": "Z"}` → 200.
  - `test_anexar_e_baixar_nf_pdf`: POST `files={"arquivo": ("nf.pdf", b"%PDF-1.4 teste", "application/pdf")}` → 200; GET item → `nf_anexo_key.startswith(f"notas_fiscais/nf_{id}_")`; GET `/nota-fiscal` → 200, `content == b"%PDF-1.4 teste"`, `headers["content-type"] == "application/pdf"`; linha de auditoria `item`/`editar` com `depois["nf_anexo_key"]`.
  - `test_nf_acima_de_10mb_413`, `test_nf_tipo_invalido_400` (`b"MZ\x90"`), `test_nf_sem_anexo_404`, `test_nf_exige_gestor_ou_tecnico` (aprendiz → 403 `"Acesso restrito a Gestor ou Técnico."`).
- [ ] **Step 2: Rodar e ver falhar:** `run-backend-tests.sh $WT f4 tests/test_items_compra.py -q` → FAIL.
- [ ] **Step 3: Implementar** validadores, schemas, `erro_local`, as mudanças nos dois endpoints, os dois endpoints da NF e os rótulos.
- [ ] **Step 4: Rodar** `tests/test_items_compra.py tests/test_items.py tests/test_audit_fluxos.py tests/test_rbac.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(ciclo-vida): valor, data da compra, garantia, local e anexo da NF no equipamento`.

### Task 4: Transferência entre unidades

**Files:**
- Create: `backend/tests/test_transferencias.py`
- Modify: `backend/app/db/transferencia_db.py`, `backend/app/schemas/items.py`, `backend/app/routers/items.py`, `backend/app/dependencies.py` (acréscimo), `backend/app/main.py` (rótulos `motivo`, `para_unidade_id`, `para_local_id`)

**Interfaces:**
- Consumes: `erro_local` (Task 3).
- Produces:
  - `class TransferenciaErro(Exception)` com `status_code: int` e `detail: str`.
  - `TransferenciaDBManager.transferir(item_id: int, para_unidade_id: int, para_local_id: int | None, motivo: str, usuario: str) -> dict` — numa transação: `SELECT ... FROM items WHERE id=%s AND is_active=1 FOR UPDATE`; devolve `{"id", "de_unidade", "para_unidade", "de_local_id", "para_local_id"}`. Validações em ordem, cada uma `TransferenciaErro`:
    1. item inexistente/inativo → 404 `"Item não encontrado."`
    2. status `Pendente` ou `Pendente Devolução` → 400 `"Conclua ou estorne o empréstimo ou a devolução pendente antes de transferir."`
    3. unidade inexistente ou `is_active=0` → 400 `"Unidade de destino não encontrada ou inativa."`
    4. nome da unidade == `items.revenda` → 400 `"O equipamento já está nesta unidade."`
    5. `para_local_id` → `erro_local(..., rotulo_unidade="de destino")` → 400
    Efeitos: `UPDATE items SET revenda=<destino>, local_id=<para_local_id ou NULL>`; `INSERT transferencias`; `INSERT history (item_id, operador, data_operacao, operation, details, revenda)` com `operation='Edição'`, `details=f"Transferência {de_unidade or 'sem unidade'} → {para_unidade}"`, `revenda=<destino>`.
  - `TransferenciaDBManager.listar_por_item(item_id: int) -> list[dict]` (com `de_local_nome`, `para_local_nome`; ordem `transferido_em, id`).
  - `get_transferencia_db() -> TransferenciaDBManager` (`@lru_cache`, em `dependencies.py`).
  - `TransferenciaCreate(BaseModel)`: `para_unidade_id: int`, `para_local_id: Optional[int] = None`, `motivo: str` (`blank_to_none`; vazio → `"Informe o motivo da transferência."`; > 255 → `"O motivo deve ter no máximo 255 caracteres."`).
  - `POST /api/items/{id}/transferir` (`gestor_or_tecnico`) → `{"detail": f"Equipamento transferido de {de} para {para}.", "id": <transferencia_id>}`; auditoria `item`/`transferir` com `antes={"revenda": de, "local_id": de_local_id}`, `depois={"revenda": para, "local_id": para_local_id, "motivo": motivo}`.
  - Rótulos: `"motivo": "Motivo"`, `"para_unidade_id": "Unidade de destino"`, `"para_local_id": "Local de destino"`.

- [ ] **Step 1: Testes** (`tests/test_transferencias.py`; unidades `criar_unidade(nome="Revalle Juazeiro")` — a revenda padrão de `criar_item` — e `criar_unidade(nome="Revalle Petrolina")`):
  - `test_transferir_atualiza_item_e_registra`: com local de Petrolina → 200, `detail == "Equipamento transferido de Revalle Juazeiro para Revalle Petrolina."`; GET item `revenda == "Revalle Petrolina"`, `local_id == <local>`; uma linha em `transferencias` com `de_unidade`, `para_unidade`, `motivo`, `transferido_por == "gestor.teste"`; linha de `history` `operation == "Edição"`, `details == "Transferência Revalle Juazeiro → Revalle Petrolina"`; auditoria `item`/`transferir`.
  - `test_transferir_sem_local_limpa_o_local_antigo` → `local_id is None`.
  - `test_transferir_para_a_mesma_unidade_400`, `test_transferir_local_de_outra_unidade_400` (`"O local informado não pertence à unidade de destino."`), `test_transferir_item_pendente_400` (fixture `item_pendente`), `test_transferir_unidade_inativa_400`, `test_transferir_item_inexistente_404`.
  - `test_transferir_sem_motivo_422`: `detail == "Motivo: Informe o motivo da transferência."`.
  - `test_transferir_exige_gestor_ou_tecnico`: aprendiz 403; técnico 200.
- [ ] **Step 2: Rodar e ver falhar:** `run-backend-tests.sh $WT f4 tests/test_transferencias.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `tests/test_transferencias.py tests/test_items.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(ciclo-vida): transferência de equipamento entre unidades com histórico e auditoria`.

### Task 5: Manutenções (backend)

**Files:**
- Create: `backend/app/db/manutencao_db.py`, `backend/app/schemas/manutencoes.py`, `backend/app/routers/manutencoes.py`, `backend/tests/test_manutencoes.py`
- Modify: `backend/app/dependencies.py` (acréscimo), `backend/app/main.py` (nova linha `from app.routers import manutencoes`, `app.include_router(manutencoes.router)` depois do último `include_router`, rótulos `"defeito": "Defeito"`, `"assistencia": "Assistência técnica"`, `"observacoes": "Observações"`, `"previsao_retorno": "Previsão de retorno"`, `"item_id": "Equipamento"`)

**Interfaces:**
- Consumes: `salvar_anexo`, `resposta_anexo`, `CATEGORIA_MANUTENCAO` (Task 2); `validar_valor_monetario`, `validar_data_br`, `data_br_para_date` (Task 3).
- Produces (`app/db/manutencao_db.py`):
  - `STATUS_MANUTENCAO = ("aberta", "enviada", "em_reparo", "retornada", "sem_conserto", "cancelada")`; `STATUS_EM_ANDAMENTO = ("aberta", "enviada", "em_reparo")`; `ROTULOS_STATUS = {"aberta": "Aberta", "enviada": "Enviada", "em_reparo": "Em reparo", "retornada": "Retornada", "sem_conserto": "Sem conserto", "cancelada": "Cancelada"}`.
  - `TRANSICOES = {"aberta": {"enviada", "em_reparo", "retornada", "sem_conserto", "cancelada"}, "enviada": {"em_reparo", "retornada", "sem_conserto", "cancelada"}, "em_reparo": {"retornada", "sem_conserto", "cancelada"}, "retornada": set(), "sem_conserto": set(), "cancelada": set()}`. Transição fora da tabela → 400 `f"Ação não permitida para manutenção com status {ROTULOS_STATUS[atual]}."`.
  - `class ManutencaoErro(Exception)` com `status_code`, `detail`. Manutenção inexistente → 404 `"Manutenção não encontrada."`.
  - `ManutencaoDBManager` (todas as escritas numa transação com `SELECT ... FOR UPDATE` na linha do item e na da manutenção):
    - `abrir(item_id: int, defeito: str, assistencia: str | None, observacoes: str | None, usuario: str) -> int` — item ativo senão 404 `"Item não encontrado."`; item com `colaborador_id` ou `assigned_to` preenchido, ou status fora de `Disponível`/`Em Manutenção` → 400 `"Só é possível abrir manutenção para equipamento sem colaborador. Faça a devolução antes."`; já existe manutenção em `STATUS_EM_ANDAMENTO` → 400 `"Este equipamento já tem uma manutenção em andamento."`. Grava `aberta_em=now`, `aberta_por`; item → `Em Manutenção`. (Status `Em Manutenção` sem ordem em andamento = depois de "sem conserto": reabrir é permitido.)
    - `enviar(manutencao_id: int, assistencia: str, previsao_retorno: date | None, usuario: str) -> None` — status → `enviada`, `enviada_em=now`; previsão antes da data de abertura → 400 `"A previsão de retorno não pode ser anterior à abertura."`.
    - `marcar_em_reparo(manutencao_id: int, previsao_retorno: date | None, usuario: str) -> None` — `em_reparo_em=now`; previsão informada substitui a anterior.
    - `registrar_retorno(manutencao_id: int, resultado: Literal["retornada", "sem_conserto"], custo: Decimal | None, observacoes: str | None, anexo_key: str | None, usuario: str) -> None` — `retornada_em=now`, `encerrada_por`, custo e anexo; `retornada` → item `Disponível`; `sem_conserto` → item continua `Em Manutenção`.
    - `cancelar(manutencao_id: int, motivo: str, usuario: str) -> None` — `cancelada_em=now`, `encerrada_por`, item → `Disponível`.
    - Observações acrescentadas com `LEFT(CONCAT_WS('\n', observacoes, %s), 1000)`; o cancelamento acrescenta `f"Cancelada: {motivo}"`.
    - `obter(manutencao_id: int) -> dict | None` e `listar(*, status: str | None = None, item_id: int | None = None, revenda: str | None = None, tipo: str | None = None, de: date | None = None, ate: date | None = None, busca: str | None = None, limit: int | None = 25, offset: int = 0) -> tuple[list[dict], int, Decimal]` — `status="em_andamento"` = os três de `STATUS_EM_ANDAMENTO`; `de`/`ate` inclusivos sobre `aberta_em` (`>= de` e `< ate + 1 dia`); `busca` com `LIKE` em `i.brand`, `i.model`, `i.identificador`, `m.assistencia`, `m.defeito`; JOIN com `items` **sem** filtrar `is_active` (itens baixados mantêm o histórico de custo); cada linha traz `item_tipo`, `item_brand`, `item_model`, `item_identificador`, `item_revenda`, `item_status` e `custo_acumulado_item` (subconsulta `COALESCE(SUM(custo), 0)` de todas as manutenções do item); o terceiro valor é `SUM(custo)` de todas as linhas filtradas (não só da página); ordem `aberta_em DESC, id DESC`; `limit=None` devolve tudo (para exportação da F7).
    - `tem_em_andamento(item_id: int) -> bool`.
  - `get_manutencao_db() -> ManutencaoDBManager` (`@lru_cache`).
  - Schemas: `ManutencaoCreate(item_id: int, defeito: str (obrigatório, ≤ 500, vazio → "Descreva o defeito."), assistencia: Optional[str] (≤ 150), observacoes: Optional[str] (≤ 500))`; `ManutencaoEnvio(assistencia: str (vazio → "Informe a assistência técnica."), previsao_retorno: Optional[str] dd/mm/aaaa)`; `ManutencaoEmReparo(previsao_retorno: Optional[str])`; `ManutencaoCancelamento(motivo: str (vazio → "Informe o motivo do cancelamento.", ≤ 255))`; `ManutencaoResponse` (todas as colunas da tabela + campos `item_*` + `custo_acumulado_item: Decimal`); `ManutencaoLista(items: list[ManutencaoResponse], total: int, custo_total: Decimal)` — valores monetários quantizados em `0.01` (JSON `"0.00"` sem linhas).
  - Router `prefix="/api/manutencoes"`, todos `gestor_or_tecnico`:
    - `GET ""` (`status: Optional[Literal[...6..., "em_andamento"]]`, `item_id`, `revenda`, `tipo`, `de: Optional[date]`, `ate: Optional[date]`, `busca`, `limit: int = Query(25, ge=1, le=settings.MAX_PAGE_SIZE)`, `offset: int = Query(0, ge=0)`) → `ManutencaoLista`. O "por equipamento" é `GET /api/manutencoes?item_id=<id>`: `custo_total` vira o custo acumulado do item.
    - `GET "/{id}"` → `ManutencaoResponse`; `GET "/{id}/anexo"` → `resposta_anexo(key, f"manutencao_{id}")`, sem anexo 404 `"Nenhum anexo nesta manutenção."`.
    - `POST ""` → `{"detail": "Manutenção aberta.", "id": id}`.
    - `POST "/{id}/enviar"` → `{"detail": "Envio à assistência registrado."}`; `POST "/{id}/em-reparo"` → `{"detail": "Manutenção marcada como em reparo."}`.
    - `POST "/{id}/retorno"` multipart: `resultado: Literal["retornada", "sem_conserto"] = Form(...)`, `custo: Optional[str] = Form(None)` (decimal com ponto; inválido → 400 `"Custo inválido: use números com ponto para os centavos (ex.: 150.00)."`; validado com `validar_valor_monetario(v, "custo")` → 400 com a mensagem), `observacoes: Optional[str] = Form(None)`, `anexo: Optional[UploadFile] = File(None)`. A checagem de status (`obter` + `STATUS_EM_ANDAMENTO`) vem **antes** de gravar o anexo. Respostas: `{"detail": "Retorno registrado. O equipamento voltou a Disponível.", "sugerir_baixa": False}` / `{"detail": "Manutenção encerrada sem conserto. Avalie a baixa do equipamento.", "sugerir_baixa": True}`.
    - `POST "/{id}/cancelar"` → `{"detail": "Manutenção cancelada. O equipamento voltou a Disponível."}`.
    - Auditoria `entidade="manutencao"`, `acao` = `criar` | `enviar` | `em_reparo` | `retornar` | `sem_conserto` | `cancelar`; `antes`/`depois` = `audit.diff(obter antes, obter depois)` (criação: `depois=obter(id)`), com `item_id` e o status do item em `depois`.

- [ ] **Step 1: Testes** (`tests/test_manutencoes.py`; helper `_abrir(client, item_id)`):
  - `test_abrir_coloca_item_em_manutencao`: 200 `{"detail": "Manutenção aberta.", ...}`; GET item `status == "Em Manutenção"`.
  - `test_item_em_manutencao_nao_fica_disponivel_nem_emprestavel`: `GET /api/items?status=Disponível` não traz o item; `inv_manager.issue(...)` (mesmos argumentos da fixture `item_pendente`) → `(False, "Este item não está disponível para empréstimo.")`.
  - `test_abrir_para_item_com_colaborador_recusa` (`item_indisponivel`) → 400 com a mensagem.
  - `test_abrir_segunda_manutencao_recusa` → 400 `"Este equipamento já tem uma manutenção em andamento."`.
  - `test_fluxo_envio_reparo_retorno_com_custo_e_anexo`: enviar (`assistencia="TecInfo"`, previsão hoje + 7 dias) → em-reparo → retorno `resultado=retornada`, `custo="150.00"`, anexo PDF → item `Disponível`; GET manutenção `status == "retornada"`, `custo == "150.00"`, `anexo_key.startswith("manutencoes/")`, `enviada_em`, `em_reparo_em`, `retornada_em` preenchidos, `encerrada_por == "gestor.teste"`; GET `/anexo` → 200.
  - `test_previsao_antes_da_abertura_400`.
  - `test_sem_conserto_mantem_item_e_permite_nova_manutencao`: retorno `sem_conserto` → `sugerir_baixa is True`, item segue `Em Manutenção`; `POST /api/manutencoes` de novo para o mesmo item → 200.
  - `test_cancelar_volta_disponivel` e `test_cancelar_sem_motivo_422`.
  - `test_transicao_invalida`: retorno numa manutenção já `retornada` → 400 `"Ação não permitida para manutenção com status Retornada."`.
  - `test_retorno_com_anexo_invalido_nao_altera`: anexo `b"MZ"` → 400 e a manutenção continua `aberta`.
  - `test_custo_negativo_400`: `custo="-1"` → 400 `"Informe um custo maior ou igual a zero."`.
  - `test_custo_acumulado_por_equipamento`: item A com dois retornos (`150.00`, `89.90`) e uma cancelada; item B com `1000.00` → `GET ?item_id=A` → `total == 3`, `custo_total == "239.90"`, toda linha com `custo_acumulado_item == "239.90"`; sem filtro → `custo_total == "1239.90"`; `GET ?item_id=A&limit=1` → `custo_total` continua `"239.90"`.
  - `test_filtros`: `status=em_andamento`, `revenda`, `tipo`, `de`/`ate`, `busca="TecInfo"` devolvem só as linhas esperadas.
  - `test_exige_gestor_ou_tecnico`: aprendiz → 403 em GET e POST.
  - `test_auditoria_de_manutencao`: linhas `manutencao`/`criar` e `manutencao`/`retornar`, com `entidade_id` da manutenção.
- [ ] **Step 2: Rodar e ver falhar:** `run-backend-tests.sh $WT f4 tests/test_manutencoes.py -q` → FAIL.
- [ ] **Step 3: Implementar** manager, schemas, router e as linhas em `dependencies.py`/`main.py`.
- [ ] **Step 4: Rodar a suíte inteira:** `run-backend-tests.sh $WT f4 -q` → verde.
- [ ] **Step 5: Commit** `feat(manutencoes): ordens de manutenção com custo, anexo e efeito no status do equipamento`.

### Task 6: Baixa com registro de apagamento de dados

**Files:**
- Create: `backend/app/services/baixa.py`, `backend/tests/test_baixa_apagamento.py`, `backend/tests/dados_baixa.py`
- Modify: `backend/app/db/inventory_manager_db.py` (só `remove`), `backend/app/routers/items.py` (`remove_item`), `backend/app/routers/constants.py`, `backend/tests/test_items.py:205-212`, `backend/tests/test_audit_fluxos.py:87-95`, `backend/tests/test_rbac.py:83-102`, `backend/tests/test_uploads.py` (as 4 chamadas `DELETE` que esperam 200/413)

**Interfaces:**
- Consumes: `salvar_anexo`, `CATEGORIA_APAGAMENTO` (Task 2); `ManutencaoDBManager.tem_em_andamento` (Task 5).
- Produces (`app/services/baixa.py`):
  - `TIPOS_EXIGEM_APAGAMENTO = ("Notebook", "Desktop", "Celular", "Tablet", "HD")`
  - `METODOS_APAGAMENTO = {"formatacao_segura": "Formatação segura", "reset_fabrica": "Reset de fábrica", "destruicao_fisica": "Destruição física"}`
  - `MOTIVOS_SEM_POSSE = ("Roubo", "Perda")` (chaves de `REMOVAL_REASONS`)
  - `@dataclass(frozen=True) class Apagamento: dados_apagados: str; metodo: str; apagado_por: str | None; apagado_em: date | None`; `NAO_APLICA = Apagamento("nao_aplica", "nao_aplica", None, None)`
  - `exige_apagamento(tipo: str | None) -> bool`
  - `validar_apagamento(*, tipo: str | None, motivo: str, dados_apagados: str | None, metodo: str | None, apagado_por: str | None, apagado_em: str | None, hoje: date | None = None) -> Apagamento` — levanta `ValueError(msg)`. Regras, nesta ordem:
    1. tipo fora de `TIPOS_EXIGEM_APAGAMENTO` → `NAO_APLICA` (ignora o que vier).
    2. `dados_apagados` vazio → `"Informe se os dados do equipamento foram apagados."`; fora de `sim`/`nao_aplica` → `"Valor inválido para dados apagados."`.
    3. `nao_aplica` com motivo em `MOTIVOS_SEM_POSSE` → `NAO_APLICA`; com outro motivo → `f"Para {tipo}, a baixa exige o apagamento dos dados (LGPD). \"Não se aplica\" só vale para Roubo ou Perda."`.
    4. `sim`: método fora de `METODOS_APAGAMENTO` → `"Informe o método de apagamento dos dados."`; `apagado_por` vazio → `"Informe quem apagou os dados."`; > 100 caracteres → `"Quem apagou: máximo de 100 caracteres."`; `apagado_em` inválido → `"Informe a data do apagamento no formato dd/mm/aaaa."`; futuro → `"A data do apagamento não pode ser no futuro."`.
  - `motivo_bloqueio_baixa(item: dict, tem_manutencao_em_andamento: bool) -> str | None` — `Disponível` → `None`; `Em Manutenção` com ordem em andamento → `"Conclua ou cancele a manutenção em andamento antes da baixa."`; `Em Manutenção` sem ordem em andamento (sem conserto) → `None`; outros → `"Não é possível remover produto emprestado."` (mensagem atual mantida).
  - `InventoryDBManager.remove(item_id: int, logged_user: str, reason: str, storage_key: str = None, apagamento: Apagamento | None = None, comprovante_key: str | None = None) -> tuple[bool, str]` — a checagem de status passa a usar `motivo_bloqueio_baixa` (conta as manutenções em andamento com SQL próprio na mesma conexão); com `apagamento`, insere `baixas_apagamento` na mesma transação com `history_id = cur.lastrowid` do `INSERT` da `Exclusão`, `registrado_por=logged_user`, `registrado_em=now`. Chamadas antigas sem `apagamento` continuam funcionando.
  - `DELETE /api/items/{id}` ganha `dados_apagados`, `metodo_apagamento`, `apagado_por`, `apagado_em` (`Optional[str] = Form(None)`) e `comprovante_apagamento: Optional[UploadFile] = File(None)`. Ordem: 404 → `motivo_bloqueio_baixa` (400) → `validar_apagamento` (400 com a mensagem) → anexo atual (20MB, como hoje) → comprovante (`salvar_anexo(..., CATEGORIA_APAGAMENTO, f"apagamento_{id}")`, só quando `dados_apagados == "sim"`) → `inv.remove(...)`. Nada é gravado no storage antes das duas validações. Auditoria `item`/`excluir` passa a ter `depois["apagamento"] = {"dados_apagados", "metodo", "apagado_por", "apagado_em"}`.
  - `/api/constants` ganha `"tipos_exigem_apagamento": list(TIPOS_EXIGEM_APAGAMENTO)`, `"motivos_sem_posse": list(MOTIVOS_SEM_POSSE)`, `"metodos_apagamento": METODOS_APAGAMENTO`.
  - `tests/dados_baixa.py`: `APAGAMENTO_OK = {"dados_apagados": "sim", "metodo_apagamento": "formatacao_segura", "apagado_por": "Técnico TI", "apagado_em": datetime.now().strftime("%d/%m/%Y")}`.
  - Os campos do apagamento são `Form`, não passam pelo handler de 422: os erros de `validar_apagamento` saem como 400 com a mensagem pura (sem prefixo de rótulo).

- [ ] **Step 1: Testes.**
  - Unit (`test_baixa_apagamento.py`, `hoje=date(2026, 11, 12)`): `test_tipo_sem_dados_vira_nao_aplica` (Impressora com qualquer entrada → `NAO_APLICA`); `test_notebook_sem_resposta_recusa`; `test_nao_aplica_para_doacao_recusado` (mensagem exata com `Notebook`); `test_roubo_aceita_nao_aplica` (`Notebook`/`Roubo` → `NAO_APLICA`); `test_celular_com_reset_valido` (`Venda`, `sim`, `reset_fabrica`, `"Ana"`, `"01/11/2026"` → `Apagamento("sim", "reset_fabrica", "Ana", date(2026, 11, 1))`); `test_data_futura_recusa` (`"13/11/2026"`); `test_metodo_invalido_recusa`; `test_motivo_bloqueio_baixa` (tabela: Disponível, Em Manutenção com/sem ordem, Pendente).
  - Integration: `test_baixa_notebook_sem_apagamento_recusada` (`data={"reason": "Doação"}` → 400 `"Informe se os dados do equipamento foram apagados."`; item continua no GET); `test_baixa_notebook_com_apagamento_grava_registro` (`APAGAMENTO_OK` + comprovante PDF → 200; linha em `baixas_apagamento` com `history_id` == id da `history` `Exclusão` do item, `metodo == "formatacao_segura"`, `apagado_por == "Técnico TI"`, `comprovante_key.startswith("apagamentos/")`); `test_baixa_impressora_grava_nao_aplica` (`criar_item(tipo="Impressora")`, só `reason`); `test_baixa_notebook_roubo_nao_aplica`; `test_baixa_item_sem_conserto_permitida` (abrir manutenção → retorno `sem_conserto` → DELETE com `APAGAMENTO_OK` → 200); `test_baixa_com_manutencao_em_andamento_recusada`; `test_comprovante_invalido_nao_baixa` (400, item ativo, nenhuma linha em `baixas_apagamento`); `test_auditoria_da_baixa_inclui_apagamento`; `test_constants_traz_regras_de_apagamento`.
  - Testes existentes (comentário "MUDANÇA INTENCIONAL (F4): Notebook exige registro de apagamento"): as chamadas `DELETE` que esperam 200 ou 413 passam a enviar `data={"reason": ..., **APAGAMENTO_OK}`; as que esperam 400 por item emprestado ficam iguais (a checagem de status vem antes); `test_monthly_report.py` (chama `inv_manager.remove` direto) fica igual.
- [ ] **Step 2: Rodar e ver falhar:** `run-backend-tests.sh $WT f4 tests/test_baixa_apagamento.py -q` → FAIL.
- [ ] **Step 3: Implementar** `baixa.py`, `remove`, o endpoint, as constantes, os rótulos e o ajuste dos testes existentes.
- [ ] **Step 4: Rodar a suíte inteira:** `run-backend-tests.sh $WT f4 -q` → verde.
- [ ] **Step 5: Commit** `feat(baixa): registro de apagamento de dados (LGPD) na baixa de equipamentos com armazenamento`.

### Task 7: Linha do tempo do equipamento (registro de fontes + endpoint)

**Files:**
- Create: `backend/app/services/linha_do_tempo.py`, `backend/app/services/linha_do_tempo_fontes.py`, `backend/app/schemas/linha_do_tempo.py`, `backend/tests/test_linha_do_tempo.py`
- Modify: `backend/app/routers/items.py`

**Interfaces:**
- Consumes: tabelas da Task 1; `METODOS_APAGAMENTO` (Task 6).
- Produces (`app/services/linha_do_tempo.py`):
  - `@dataclass class Evento: ocorrido_em: datetime; origem: str; tipo: str; titulo: str; descricao: str | None = None; usuario: str | None = None; referencia_id: int | None = None`
  - `FonteLinhaDoTempo = Callable[[int], list[Evento]]`
  - `_FONTES: dict[str, FonteLinhaDoTempo]` (ordem de registro preservada)
  - `registrar_fonte(func: FonteLinhaDoTempo) -> FonteLinhaDoTempo` — usável como decorador; chave = `func.__name__`; registrar de novo o mesmo nome substitui sem mudar a posição.
  - `fontes_registradas() -> list[str]`
  - `MODULOS_DE_FONTES: tuple[str, ...] = ("app.services.linha_do_tempo_fontes",)` — **a F7 acrescenta aqui o módulo das fontes de termos, linha telefônica e inventário.**
  - `carregar_fontes() -> None` — `importlib.import_module` de cada módulo (idempotente).
  - `@dataclass class LinhaDoTempo: eventos: list[Evento]; fontes_com_erro: list[str]`
  - `montar_linha_do_tempo(item_id: int) -> LinhaDoTempo` — chama `carregar_fontes()`, depois cada fonte; exceção numa fonte → `logger.exception` e o nome vai para `fontes_com_erro` (as outras continuam); ordena por `(ocorrido_em, posição da fonte no registro, referencia_id or 0)`, crescente.
  - `item_existe(item_id: int) -> bool` (qualquer `is_active`: equipamento baixado continua com linha do tempo).
- Produces (`linha_do_tempo_fontes.py`, cada função com `@registrar_fonte`; títulos exatos):
  - `fonte_historico` (`origem="historico"`, `referencia_id=history.id`, `usuario=operador`, `descricao=details`): `history` do item com `is_reversed = 0`, **sem** `Exclusão` (vem da fonte de baixa) e **sem** `Edição` cujo `details LIKE %s` com parâmetro `"Transferência %"` (vem da fonte de transferência; com parâmetros o `%` vai no valor, nunca literal no SQL). Mapa `operation → (tipo, titulo)`: `Cadastro → ("cadastro", "Cadastro do equipamento")`; `Edição → ("edicao", "Edição dos dados")`; `Empréstimo → ("emprestimo", f"Empréstimo para {usuario}" ou "Empréstimo")`; `Confirmação Empréstimo → ("emprestimo_confirmado", "Empréstimo confirmado")`; `Devolução → ("devolucao", "Devolução iniciada")`; `Confirmação Devolução → ("devolucao_confirmada", "Devolução confirmada")`; `Estorno → ("estorno", "Estorno de operação")`; `Vínculo Periférico → ("periferico", "Periférico vinculado")`; `Desvínculo Periférico → ("periferico", "Periférico desvinculado")`; `Substituição Periférico → ("periferico", "Periférico substituído")`; qualquer outra → `("outro", operation)`.
  - `fonte_manutencoes` (`origem="manutencao"`, `referencia_id=manutencao.id`): por manutenção, um evento por data preenchida — `aberta_em → ("manutencao_aberta", "Manutenção aberta", descricao=defeito, usuario=aberta_por)`; `enviada_em → ("manutencao_enviada", f"Enviado à assistência {assistencia}", descricao=f"Previsão de retorno: {dd/mm/aaaa}" ou None)`; `em_reparo_em → ("manutencao_em_reparo", "Em reparo na assistência")`; `retornada_em` com status `retornada` → `("manutencao_retornada", "Retorno da manutenção", descricao=f"Custo: {moeda_br(custo)}" ou None, usuario=encerrada_por)`; com status `sem_conserto` → `("manutencao_sem_conserto", "Manutenção encerrada sem conserto", descricao="Baixa sugerida." + (f" Custo: {moeda_br(custo)}" se houver), usuario=encerrada_por)`; `cancelada_em → ("manutencao_cancelada", "Manutenção cancelada", usuario=encerrada_por)`.
  - `fonte_transferencias` (`origem="transferencia"`, `tipo="transferencia"`): `titulo=f"Transferência {de_unidade or '—'} → {para_unidade}"`, `descricao=motivo` + `f" · Local: {de_local_nome or '—'} → {para_local_nome or '—'}"` quando algum local existir, `usuario=transferido_por`.
  - `fonte_baixa` (`origem="baixa"`, `tipo="baixa"`): `history` `Exclusão` do item (`is_reversed=0`) `LEFT JOIN baixas_apagamento ON history_id`; `titulo=f"Baixa do equipamento — {details}"`; `descricao`: `sim` → `f"Dados apagados ({METODOS_APAGAMENTO[metodo]}) por {apagado_por} em {dd/mm/aaaa}"`; `nao_aplica` → `"Apagamento de dados: não se aplica"`; baixa antiga sem registro → `None`.
  - `moeda_br(valor: Decimal) -> str` (`Decimal("1234.5")` → `"R$ 1.234,50"`).
- Produces (endpoint e schema): `EventoResponse` (campos de `Evento`), `LinhaDoTempoResponse(item_id: int, eventos: list[EventoResponse], fontes_com_erro: list[str])`; `GET /api/items/{item_id}/linha-do-tempo` (`gestor_or_tecnico`; `item_existe` falso → 404 `"Item não encontrado."`).

- [ ] **Step 1: Testes** (`tests/test_linha_do_tempo.py`):
  - Unit: `test_registrar_fonte_substitui_pelo_nome` (`monkeypatch.setattr(linha_do_tempo, "_FONTES", {})`); `test_empate_ordena_por_fonte_e_referencia` (duas fontes falsas com o mesmo `ocorrido_em`); `test_moeda_br`.
  - `test_fonte_com_erro_nao_derruba`: `monkeypatch.setattr(linha_do_tempo, "_FONTES", dict(linha_do_tempo._FONTES))` + fonte `fonte_quebrada` que levanta → 200, `fontes_com_erro == ["fonte_quebrada"]` e o evento `cadastro` presente.
  - `test_linha_do_tempo_completa_em_ordem`: via API, cadastrar item (Notebook, revenda Juazeiro), editar (`brand`), transferir para Petrolina, abrir manutenção, retorno com `custo="150.00"`, baixa com `APAGAMENTO_OK`. Como `DATETIME` tem precisão de segundos, regravar os horários com `executar_sql` em `2026-01-01 10:00` + n minutos na ordem das ações (`history` por `operation`, `transferencias.transferido_em`, `manutencoes.aberta_em`/`retornada_em`). Afirmar `[e["tipo"] for e in eventos] == ["cadastro", "edicao", "transferencia", "manutencao_aberta", "manutencao_retornada", "baixa"]` (sem `edicao` duplicada da transferência), `titulo` da transferência `"Transferência Revalle Juazeiro → Revalle Petrolina"`, `descricao` do retorno `"Custo: R$ 150,00"`, `descricao` da baixa começa com `"Dados apagados (Formatação segura) por Técnico TI em "`.
  - `test_baixa_antiga_sem_registro_de_apagamento` (`inv_manager.remove` sem `apagamento` → evento `baixa` com `descricao is None`).
  - `test_item_baixado_tem_linha_do_tempo` (200), `test_item_inexistente_404`, `test_exige_gestor_ou_tecnico` (aprendiz 403).
- [ ] **Step 2: Rodar e ver falhar:** `run-backend-tests.sh $WT f4 tests/test_linha_do_tempo.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar a suíte inteira:** `run-backend-tests.sh $WT f4 -q` → verde.
- [ ] **Step 5: Commit** `feat(ciclo-vida): linha do tempo do equipamento com registro de fontes`.

### Task 8: Frontend — clientes de API, constantes e helpers da F4

**Files:**
- Create: `frontend/src/api/manutencoes.ts`, `frontend/src/api/manutencoes.test.ts`, `frontend/src/features/ciclo-vida/format.ts`, `frontend/src/features/ciclo-vida/format.test.ts`, `frontend/src/features/manutencoes/status.ts`
- Modify: `frontend/src/api/items.ts` (só acréscimos), `frontend/src/api/items.test.ts`, `frontend/src/api/constants.ts`, `frontend/src/hooks/useConstants.ts`, `frontend/src/pages/UnidadesPage.tsx:107-112` (só a constante `STATUS_ORDER` e o comentário), `frontend/src/pages/UnidadesPage.test.tsx`, `frontend/src/api/unidades.ts` (comentário "4 chaves" → "5 chaves")

**Interfaces:**
- Consumes: contratos das Tasks 3–7.
- Produces:
  - `api/items.ts`: `Item` ganha `valor_compra?: string | null; data_compra?: string | null; garantia_ate?: string | null; nf_anexo_key?: string | null`.
    - `interface ApagamentoInput { dados_apagados: 'sim' | 'nao_aplica'; metodo_apagamento?: string; apagado_por?: string; apagado_em?: string; comprovante?: File | null }`
    - `removeItem(id: number, reason: string, attachment?: File, apagamento?: ApagamentoInput)` — acrescenta ao `FormData` só os campos preenchidos (`comprovante` vai como `comprovante_apagamento`).
    - `uploadNotaFiscal(id: number, arquivo: File): Promise<{ detail: string; nf_anexo_key: string }>` (`POST /items/{id}/nota-fiscal`, campo `arquivo`).
    - `baixarNotaFiscal(item: Pick<Item, 'id' | 'nf_anexo_key'>): Promise<void>` — `downloadAuthenticated(`/items/${id}/nota-fiscal`, `nota_fiscal_item_${id}.<extensão da chave>`)`.
    - `interface TransferenciaInput { para_unidade_id: number; para_local_id: number | null; motivo: string }`; `transferirItem(id: number, data: TransferenciaInput): Promise<{ detail: string; id: number }>`.
    - `interface EventoLinhaDoTempo { ocorrido_em: string; origem: string; tipo: string; titulo: string; descricao?: string | null; usuario?: string | null; referencia_id?: number | null }`; `interface LinhaDoTempo { item_id: number; eventos: EventoLinhaDoTempo[]; fontes_com_erro: string[] }`; `getLinhaDoTempo(id: number): Promise<LinhaDoTempo>`.
  - `api/manutencoes.ts`: `type StatusManutencao = 'aberta' | 'enviada' | 'em_reparo' | 'retornada' | 'sem_conserto' | 'cancelada'`; `interface Manutencao` (espelho de `ManutencaoResponse`; decimais e datas como `string`); `interface ManutencaoLista { items: Manutencao[]; total: number; custo_total: string }`; `interface ListManutencoesParams { status?: StatusManutencao | 'em_andamento'; item_id?: number; revenda?: string; tipo?: string; de?: string; ate?: string; busca?: string; limit?: number; offset?: number }`; `listManutencoes(params)`, `getManutencao(id)`, `abrirManutencao(data: { item_id: number; defeito: string; assistencia?: string; observacoes?: string }): Promise<{ detail: string; id: number }>`, `enviarManutencao(id, data: { assistencia: string; previsao_retorno?: string })`, `marcarEmReparo(id, data: { previsao_retorno?: string })`, `registrarRetorno(id, data: { resultado: 'retornada' | 'sem_conserto'; custo?: number | null; observacoes?: string; anexo?: File | null }): Promise<{ detail: string; sugerir_baixa: boolean }>` (multipart; `custo` vai como `custo.toFixed(2)`), `cancelarManutencao(id, data: { motivo: string })`, `baixarAnexoManutencao(m: Pick<Manutencao, 'id' | 'anexo_key'>)`.
  - `api/constants.ts`: `Constants` ganha opcionais `item_statuses?: string[]; tipos_exigem_apagamento?: string[]; motivos_sem_posse?: string[]; metodos_apagamento?: Record<string, string>`. `useConstants()` devolve também `itemStatuses`, `tiposExigemApagamento`, `motivosSemPosse` (padrão `[]`) e `metodosApagamento` (padrão `{}`).
  - `features/ciclo-vida/format.ts`:
    - `parseMoedaBR(texto: string): number | null` — vazio, negativo ou não numérico → `null`; com vírgula: tira os pontos e troca a vírgula por ponto; sem vírgula: `^\d+\.\d{1,2}$` é decimal, senão os pontos são de milhar; mais de 2 casas → `null`.
    - `formatMoedaInput(v: string | number | null | undefined): string` (`"4599.90"` → `"4.599,90"`; vazio → `''`).
    - `maskMoedaInput(v: string): string` (aceita só dígitos, `.` e `,`).
    - `isoParaBR(iso: string | null | undefined): string` — só por texto (`'2026-01-10'` e `'2026-01-10T00:00:00'` → `'10/01/2026'`; vazio → `''`); nunca `new Date()`.
    - `isDataBR(texto: string): boolean` (data de calendário válida em dd/mm/aaaa); `maskDataBR(v: string): string` (dígitos → `dd/mm/aaaa`); `hojeBR(): string`.
    - `MAX_ANEXO_BYTES = 10 * 1024 * 1024`; `anexoGrandeDemais(file: File): boolean`.
  - `features/manutencoes/status.ts`: `ROTULO_STATUS_MANUTENCAO: Record<StatusManutencao, string>` (`Aberta`, `Enviada`, `Em reparo`, `Retornada`, `Sem conserto`, `Cancelada`); `TOM_STATUS_MANUTENCAO: Record<StatusManutencao, StatusTone>` (`aberta: 'pending'`, `enviada: 'maintenance'`, `em_reparo: 'maintenance'`, `retornada: 'available'`, `sem_conserto: 'danger'`, `cancelada: 'retired'`) — passado sempre como `tone` do `StatusPill`, sem mexer em `lib/status.ts`; `STATUS_EM_ANDAMENTO: StatusManutencao[] = ['aberta', 'enviada', 'em_reparo']`.
  - `UnidadesPage`: `STATUS_ORDER = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução', 'Em Manutenção']`.

- [ ] **Step 1: Testes:**
  - `format.test.ts`: `parseMoedaBR` — `'4.599,90'`, `'4599,9'`, `'4599.90'` → `4599.9`; `'1.234'` → `1234`; `''`, `'abc'`, `'-5'`, `'1,234'` (3 casas) → `null`; `formatMoedaInput('4599.90') === '4.599,90'`; `isoParaBR('2026-01-10') === '10/01/2026'`, `isoParaBR('2026-01-10T00:00:00') === '10/01/2026'`, `isoParaBR(null) === ''`; `isDataBR('31/02/2026') === false`, `isDataBR('29/02/2028') === true`; `maskDataBR('10012026') === '10/01/2026'`.
  - `items.test.ts`: `removeItem(77, 'Doação', undefined, { dados_apagados: 'sim', metodo_apagamento: 'formatacao_segura', apagado_por: 'Ana', apagado_em: '10/11/2026' })` → o `FormData` recebido pelo MSW tem `reason`, os quatro campos e nenhum `comprovante_apagamento`; `uploadNotaFiscal` envia o campo `arquivo`; `transferirItem` envia o JSON.
  - `manutencoes.test.ts`: `listManutencoes({ status: 'em_andamento', limit: 25, offset: 0 })` repassa os parâmetros; `registrarRetorno(5, { resultado: 'retornada', custo: 150 })` envia `custo === '150.00'`.
  - `UnidadesPage.test.tsx`: indicadores com `'Em Manutenção': 2` mostram `'Em Manutenção: 2'`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/ciclo-vida src/api src/pages/UnidadesPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(ciclo-vida): clientes de API, constantes e formatação de valores e datas da F4`.

### Task 9: Frontend — Cadastro/edição com compra, NF, local e transferência

**Files:**
- Create: `frontend/src/features/ciclo-vida/TransferirItemDialog.tsx`, `frontend/src/pages/RegisterItemPage.test.tsx`
- Modify: `frontend/src/pages/RegisterItemPage.tsx`

**Interfaces:**
- Consumes: `createItem`, `updateItem`, `getItem`, `uploadNotaFiscal`, `baixarNotaFiscal`, `transferirItem` (Task 8); `listUnidades` (`api/unidades.ts`); `LocalSelect`, `StatusPill`, `Button loading`, `upper`, `FileUpload`, `Dialog` (F0); helpers de `features/ciclo-vida/format.ts`.
- Produces:
  - `TransferirItemDialog({ item, open, onOpenChange, onTransferido }: { item: Pick<Item, 'id' | 'revenda' | 'local_id'>; open: boolean; onOpenChange: (open: boolean) => void; onTransferido?: (r: { para_unidade: string; para_local_id: number | null }) => void })` — `<select aria-label="Unidade de destino">` com as unidades ativas menos a atual (rótulo `upper(nome)`); `LocalSelect` da unidade escolhida; textarea `Motivo da transferência` (obrigatório); botão `Confirmar transferência` com `loading`. Sucesso: `toast(detail)`, invalida `['item']`, `['items']`, `['item-linha-do-tempo']`, chama `onTransferido`, fecha. Erro: `toast(getErrorMessage(err), 'error')`.
  - `RegisterItemPage` (dono: F4):
    - Bloco 2 ganha `LocalSelect` (unidade = `listUnidades()` cujo `nome === revenda`; sem correspondência → `null`, seletor desabilitado). Itens do select de revenda exibidos com `upper()` (o valor continua o nome gravado).
    - Bloco novo "3. Compra e garantia": `Valor de compra (R$)` (`id="valor_compra"`, `maskMoedaInput`), `Data da compra` (`id="data_compra"`), `Fim da garantia` (`id="garantia_ate"`) com `maskDataBR` e placeholder `DD/MM/AAAA`; `FileUpload` com `accept` PDF/PNG/JPG e rótulo `Arraste ou clique para anexar a nota fiscal (PDF ou imagem, até 10 MB)`; em edição com `nf_anexo_key`, texto "Nota fiscal anexada" + botão `Baixar nota fiscal`. Campos específicos do tipo viram o bloco 4.
    - Edição: hidrata com `formatMoedaInput(valor_compra)`, `isoParaBR(data_compra)`, `isoParaBR(garantia_ate)`, `local_id`; select de revenda `disabled`, com o botão `Transferir` (abre `TransferirItemDialog`; em `onTransferido`, atualiza `revenda` e `localId` do formulário) e a ajuda "Para mudar a unidade, use Transferir: a mudança fica registrada na linha do tempo."; `StatusPill` do status atual no cabeçalho.
    - Validação antes de enviar (toasts exatos): valor preenchido e `parseMoedaBR` nulo → `Valor de compra inválido.`; data preenchida e `!isDataBR` → `Data da compra inválida. Use dd/mm/aaaa.` / `Fim da garantia inválido. Use dd/mm/aaaa.`; arquivo com `anexoGrandeDemais` (também ao selecionar, descartando o arquivo) → `A nota fiscal deve ter no máximo 10 MB.`.
    - Payload: criação inclui `valor_compra` (número), `data_compra`, `garantia_ate`, `local_id` só quando preenchidos; edição envia os quatro sempre, com `null` quando vazios (limpa no backend).
    - `mutationFn` salva e, havendo arquivo, chama `uploadNotaFiscal(id, arquivo)`; devolve `{ id, erroNf?: string }`. Sem erro → toast de sucesso atual e `navigate('/stock')`. Com `erroNf` → `toast(`Equipamento salvo, mas a nota fiscal não foi anexada: ${erroNf}`, 'error')` e, na criação, `navigate(`/edit/${id}`)`; na edição, fica na tela.
    - Botão de envio com `loading={mutation.isPending}` (o rótulo fixo; sai o "Salvando…").

- [ ] **Step 1: Testes** (`RegisterItemPage.test.tsx`; stubs de Pointer Events como no `RemovePage.test.tsx`; `MemoryRouter` com rotas `/register`, `/edit/:id`, `/stock` → `<p>Estoque</p>`; MSW para `/api/constants` (revendas `['Revalle Juazeiro', 'Revalle Petrolina']`, tipos `['Notebook']`), `/api/unidades` (`[{ id: 1, nome: 'Revalle Juazeiro', ... }, { id: 2, nome: 'Revalle Petrolina', ... }]`), `/api/locais` (`[{ id: 5, unidade_id: 1, nome: 'Sala TI', ativo: true }]`)):
  - `cadastra com dados de compra e local e depois anexa a NF`: preencher; corpo do POST contém `valor_compra: 4599.9`, `data_compra: '10/01/2026'`, `garantia_ate: '10/01/2027'`, `local_id: 5`; `POST /api/items/42/nota-fiscal` chamado uma vez; aparece `Estoque`.
  - `recusa NF acima de 10 MB sem enviar` (`Object.defineProperty(file, 'size', { value: 10 * 1024 * 1024 + 1 })`) → toast `A nota fiscal deve ter no máximo 10 MB.`; nenhum POST de NF.
  - `falha no anexo leva para a edição do item` (upload → 413 `{detail: 'Arquivo excede o tamanho máximo de 10MB.'}`) → toast `Equipamento salvo, mas a nota fiscal não foi anexada: Arquivo excede o tamanho máximo de 10MB.` e rota `/edit/42`; um único `POST /api/items`.
  - `edição hidrata datas sem deslocar o dia`: item com `valor_compra: '4599.90'`, `data_compra: '2026-01-10'` → campos `4.599,90` e `10/01/2026`.
  - `edição bloqueia a unidade e transfere`: select de revenda desabilitado; `Transferir` → no diálogo escolher `REVALLE PETROLINA`, motivo `Mudança de setor`, `Confirmar transferência` → corpo `{ para_unidade_id: 2, para_local_id: null, motivo: 'Mudança de setor' }`.
  - `edição envia null ao limpar o valor de compra` → corpo do PUT com `valor_compra: null`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/RegisterItemPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar** `TransferirItemDialog` e as mudanças na página. `TypeSpecificFields` não muda (não tem status, botão nem unidade).
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(ciclo-vida): compra, garantia, NF, local e transferência no cadastro do equipamento`.

### Task 10: Frontend — Baixa com apagamento de dados

**Files:**
- Create: `frontend/src/features/ciclo-vida/ApagamentoDadosFields.tsx`
- Modify: `frontend/src/pages/RemovePage.tsx`, `frontend/src/pages/RemovePage.test.tsx` (os dois testes atuais ficam; acrescentar)

**Interfaces:**
- Consumes: `removeItem` + `ApagamentoInput` (Task 8); `useConstants().tiposExigemApagamento | motivosSemPosse | metodosApagamento`; `StatusPill`, `Button loading`, `upper`, `FileUpload` (F0); `isDataBR`, `maskDataBR`, `hojeBR`, `anexoGrandeDemais` (Task 8).
- Produces:
  - `interface ApagamentoState { dadosApagados: '' | 'sim' | 'nao_aplica'; metodo: string; apagadoPor: string; apagadoEm: string; comprovante: File | null }`; `apagamentoInicial(): ApagamentoState` (`apagadoEm = hojeBR()`); `apagamentoCompleto(s: ApagamentoState, permiteNaoAplica: boolean): boolean`; `paraApagamentoInput(s: ApagamentoState): ApagamentoInput`.
  - `ApagamentoDadosFields({ tipo, value, onChange, metodos, permiteNaoAplica }: { tipo: string; value: ApagamentoState; onChange: (s: ApagamentoState) => void; metodos: Record<string, string>; permiteNaoAplica: boolean })` — `fieldset` com legenda `Apagamento de dados (LGPD) — obrigatório para {tipo}`; rádios `Sim, os dados foram apagados` e (só com `permiteNaoAplica`) `Não se aplica — equipamento fora de posse (roubo ou perda)`; com "Sim": `<select aria-label="Método de apagamento">` (opções de `metodos`), `Quem apagou os dados`, `Data do apagamento` (`maskDataBR`), `FileUpload` `Comprovante do apagamento (opcional)` (PDF/imagem; acima de 10 MB descarta com toast `O comprovante deve ter no máximo 10 MB.`).
  - `RemovePage` (dono: F4):
    - `?item=<id>` (via `useSearchParams`) já abre com o equipamento selecionado (é o atalho "Dar baixa" da aba Manutenções para itens sem conserto, que não aparecem na busca de `Disponível`).
    - `exige = tiposExigemApagamento.includes(selectedItem?.tipo ?? '')`; `permiteNaoAplica = motivosSemPosse.includes(reason)`; trocar para um motivo que não permite mais "não se aplica" zera a escolha.
    - Item selecionado de tipo que não exige: texto `Este tipo de equipamento não armazena dados pessoais: o apagamento será registrado como "não se aplica".`
    - Botão de confirmar também fica travado com `exige && !apagamentoCompleto(...)`; usa `loading={mutation.isPending}` (sai o "Processando…").
    - Mutação: `removeItem(id, reason, attachment ?? undefined, exige ? paraApagamentoInput(apagamento) : undefined)`; sucesso também zera o apagamento.
    - Ficha do ativo: `Badge` de status vira `<StatusPill status={selectedItem.status} />`; revenda com `upper()`.

- [ ] **Step 1: Testes** (acrescentar em `RemovePage.test.tsx`; mock de constantes com `removal_reasons: ['Doação', 'Roubo']`, `tipos_exigem_apagamento: ['Notebook']`, `motivos_sem_posse: ['Roubo']`, `metodos_apagamento: { formatacao_segura: 'Formatação segura', reset_fabrica: 'Reset de fábrica', destruicao_fisica: 'Destruição física' }`):
  - `Notebook com Doação exige o apagamento antes de destravar`: item + `Doação` + ciência → botão desabilitado; marcar `Sim, os dados foram apagados`, escolher `Formatação segura`, digitar `Ana TI` → habilitado; confirmar → `FormData` do `DELETE /api/items/77` com `dados_apagados=sim`, `metodo_apagamento=formatacao_segura`, `apagado_por=Ana TI` e `apagado_em` igual à data de hoje em dd/mm/aaaa.
  - `"Não se aplica" só aparece para Roubo`: com `Doação` a opção não existe; com `Roubo` existe e, marcada, destrava o botão (com a ciência).
  - `Impressora não pede apagamento`: item `tipo: 'Impressora'` → texto de "não se aplica" visível, sem o `fieldset`; o `FormData` não tem `dados_apagados`.
  - `?item=77 abre com o equipamento selecionado`: `MemoryRouter initialEntries={['/remove?item=77']}` → `Ficha do Ativo Selecionado` e o `StatusPill` com `Disponível`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/RemovePage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(baixa): registro de apagamento de dados na tela de baixa`.

### Task 11: Frontend — aba "Linha do tempo" (ordem 20)

**Files:**
- Create: `frontend/src/features/item-detail/tabs/linha-do-tempo.tsx`, `frontend/src/features/item-detail/tabs/linha-do-tempo.test.tsx`, `frontend/src/features/ciclo-vida/CompraResumo.tsx`
- Modify: `frontend/src/pages/ItemDetailPage.test.tsx:60-61` (só a asserção da lista de abas)

**Interfaces:**
- Consumes: `getLinhaDoTempo`, `baixarNotaFiscal` (Task 8); `TransferirItemDialog` (Task 9); `isoParaBR` (Task 8); `formatMoeda` (F0 `lib/format.ts`); `formatDateTime` (`lib/utils.ts`).
- Produces:
  - `export default { id: 'linha-do-tempo', label: 'Linha do tempo', order: 20, roles: ['Gestor', 'Técnico'], Component }` (tipo `ItemDetailTab`).
  - `CompraResumo({ item }: { item: Item })` — mostra, quando houver: `Valor de compra` (`formatMoeda`), `Data da compra` e `Garantia até` (`isoParaBR`), botão `Baixar nota fiscal` (com `nf_anexo_key`); nada quando nenhum campo existe. Fica no topo desta aba porque a aba Dados é da F0/F1 e esta é a aba do ciclo de vida (compra → uso → manutenção → baixa).
  - Aba: query `['item-linha-do-tempo', item.id]`; botão `Transferir unidade` (abre `TransferirItemDialog`); lista `<ol aria-label="Linha do tempo">` na ordem da API (cronológica), cada evento com data/hora (`formatDateTime`), ícone por `origem` (`historico` → `History`, `manutencao` → `Wrench`, `transferencia` → `ArrowLeftRight`, `baixa` → `Trash2`, outra → `Circle`, para as fontes da F7 aparecerem sem mudar esta aba), `titulo` em negrito, `descricao` e `por {usuario}`; `fontes_com_erro` não vazio → aviso `Parte da linha do tempo não pôde ser carregada. Tente novamente mais tarde.`; sem eventos → `Nenhum evento registrado para este equipamento.`
  - `ItemDetailPage.test.tsx`: `expect(tabs.map((t) => t.textContent)).toEqual(['Dados', 'Histórico'])` vira
    ```ts
    const nomes = tabs.map((t) => t.textContent)
    expect(nomes[0]).toBe('Dados')
    expect(nomes[nomes.length - 1]).toBe('Histórico')
    ```
    (mesma troca que as outras fases que acrescentam abas devem fazer, para o merge não conflitar).

- [ ] **Step 1: Testes** (`linha-do-tempo.test.tsx`, MSW + `vi.mock('@/contexts/AuthContext')` como no `ItemDetailPage.test.tsx`):
  - `mostra os eventos na ordem da API`: três eventos (`cadastro`, `transferencia`, `baixa`) → itens da lista na mesma ordem, com títulos, `por gestor` e a descrição.
  - `avisa quando uma fonte falhou` (`fontes_com_erro: ['fonte_termos']`).
  - `evento de origem desconhecida aparece` (`origem: 'termo'`, título `Termo assinado`).
  - `mostra o resumo de compra com data sem deslocar o dia` (`garantia_ate: '2027-01-10'` → `10/01/2027`; `valor_compra: '4599.90'` → `R$ 4.599,90`).
  - `Transferir unidade abre o diálogo`.
  - Rodar também `ItemDetailPage.test.tsx` com a asserção ajustada.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/item-detail src/pages/ItemDetailPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(ciclo-vida): aba Linha do tempo no detalhe do equipamento`.

### Task 12: Frontend — componentes de manutenção e aba "Manutenções" (ordem 40)

**Files:**
- Create: `frontend/src/features/manutencoes/regras.ts`, `frontend/src/features/manutencoes/AbrirManutencaoDialog.tsx`, `frontend/src/features/manutencoes/ManutencaoAcoes.tsx`, `frontend/src/features/manutencoes/ManutencaoAcoes.test.tsx`, `frontend/src/features/manutencoes/ManutencoesLista.tsx`, `frontend/src/features/item-detail/tabs/manutencoes.tsx`, `frontend/src/features/item-detail/tabs/manutencoes.test.tsx`

**Interfaces:**
- Consumes: `api/manutencoes.ts`, `features/manutencoes/status.ts`, `parseMoedaBR`, `maskMoedaInput`, `maskDataBR`, `isDataBR`, `isoParaBR`, `anexoGrandeDemais` (Task 8); `listItemsPaginated` (`api/items.ts`); `StatusPill`, `Pagination`, `usePageSize`, `Button loading`, `SearchableSelect`, `FileUpload`, `Dialog`, `formatMoeda`, `upper` (F0); `useAuth().hasRole`.
- Produces:
  - `podeAbrirManutencao(item: Item, temEmAndamento: boolean): { pode: boolean; motivo?: string }` — com `colaborador_id` ou `assigned_to` → `Equipamento com colaborador: faça a devolução antes de abrir a manutenção.`; `temEmAndamento` → `Já existe uma manutenção em andamento.`; status fora de `Disponível`/`Em Manutenção` → `Só é possível abrir manutenção para equipamento disponível.`
  - `AbrirManutencaoDialog({ item, open, onOpenChange, onAberta }: { item?: Pick<Item, 'id' | 'tipo' | 'brand' | 'model' | 'identificador'> | null; open: boolean; onOpenChange: (open: boolean) => void; onAberta?: (id: number) => void })` — sem `item`, mostra `SearchableSelect` de equipamentos `Disponível` (busca no servidor com debounce de 300 ms, `limit: 20`, como na `RemovePage`); campos `Defeito` (obrigatório), `Assistência técnica`, `Observações`; botão `Abrir manutenção` com `loading`.
  - `ManutencaoAcoes({ manutencao }: { manutencao: Manutencao })` — botões por status: `aberta` → `Registrar envio`, `Marcar em reparo`, `Registrar retorno`, `Cancelar manutenção`; `enviada` → os três últimos; `em_reparo` → `Registrar retorno`, `Cancelar manutenção`; com `anexo_key` → `Baixar anexo`. Diálogos: envio (`Assistência técnica` pré-preenchida, `Previsão de retorno` dd/mm/aaaa opcional); retorno (rádios `Consertado — volta para Disponível` / `Sem conserto`, `Custo (R$)`, `Observações`, `FileUpload` de anexo até 10 MB); cancelamento (`Motivo do cancelamento` obrigatório). `Marcar em reparo` é ação direta com `loading`. Todo sucesso: `toast(detail)` e invalida `['manutencoes']`, `['item']`, `['items']`, `['item-linha-do-tempo']`.
  - `ManutencoesLista({ manutencoes, mostrarEquipamento }: { manutencoes: Manutencao[]; mostrarEquipamento: boolean })` — tabela (em telas estreitas, cartões): abertura (`formatDateTime`), equipamento (se `mostrarEquipamento`: tipo/marca/modelo/identificador com link para `/items/{item_id}?aba=manutencoes`, unidade com `upper`), defeito, assistência, previsão (`isoParaBR`), `StatusPill` com `ROTULO_`/`TOM_STATUS_MANUTENCAO`, custo (`formatMoeda`), custo acumulado do equipamento (só com `mostrarEquipamento`), `ManutencaoAcoes`.
  - Aba `export default { id: 'manutencoes', label: 'Manutenções', order: 40, roles: ['Gestor', 'Técnico'], Component }` — `listManutencoes({ item_id, limit: pageSize, offset })` com `Pagination` + `usePageSize('item-manutencoes')` (mesmo padrão de página da aba Histórico); cabeçalho `Custo acumulado: {formatMoeda(custo_total)}`; botão `Abrir manutenção` desabilitado quando `podeAbrirManutencao` nega, com o motivo em texto ao lado; havendo manutenção `sem_conserto`, item `Em Manutenção` e nenhuma em andamento: aviso `Equipamento sem conserto. Avalie a baixa.` e, só para Gestor, link `Dar baixa` para `/remove?item={id}`.

- [ ] **Step 1: Testes:**
  - `manutencoes.test.tsx` (aba): lista e `Custo acumulado: R$ 239,90`; item com `colaborador_id: 3` → `Abrir manutenção` desabilitado e o texto `Equipamento com colaborador: faça a devolução antes de abrir a manutenção.`; item `Disponível` → abrir com `Tela quebrada` envia `{ item_id: 7, defeito: 'Tela quebrada' }` e refaz a consulta; `sem_conserto` mostra `Dar baixa` com `href="/remove?item=7"` para Gestor e não mostra para Técnico.
  - `ManutencaoAcoes.test.tsx`: `aberta` mostra os quatro botões e `retornada` nenhum (só `Baixar anexo` com anexo); retorno com `Custo (R$)` `150,00` envia `resultado=retornada` e `custo=150.00`; cancelar sem motivo não chama a API; envio repassa `assistencia` e `previsao_retorno`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/manutencoes src/features/item-detail/tabs/manutencoes.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(manutencoes): aba Manutenções no detalhe do equipamento`.

### Task 13: Frontend — página `/manutencoes`

**Files:**
- Create: `frontend/src/pages/ManutencoesPage.tsx`, `frontend/src/pages/ManutencoesPage.test.tsx`
- Modify: `frontend/src/App.tsx` (uma linha de import e um bloco `<Route>` novo, Gestor e Técnico), `frontend/src/components/layout/navigation.ts` (linha nova `import { Wrench } from 'lucide-react'` e, no grupo "Operações", depois de "Vincular Periférico": `{ to: '/manutencoes', label: 'Manutenções', icon: Wrench, roles: ['Gestor', 'Técnico'] },`)

**Interfaces:**
- Consumes: `listManutencoes` (Task 8); `ManutencoesLista`, `AbrirManutencaoDialog` (Task 12); `useConstants().revendas | equipmentTypes`; `Pagination`, `usePageSize('manutencoes')`, `useDebouncedValue`, `upper`, `formatMoeda` (F0).
- Produces: `ManutencoesPage` em `/manutencoes`:
  - Filtros: `Buscar manutenções` (debounce 300 ms), `<select aria-label="Status">` (`Todas`, `Em andamento` → `em_andamento`, e os 6 rótulos), `<select aria-label="Unidade">` (revendas com `upper`), `<select aria-label="Tipo">`, período `Aberta de` / `até` (`type="date"`, enviados como `de`/`ate` `AAAA-MM-DD`, como na `AuditPage`); `?item_id=<id>` na URL filtra por equipamento e mostra o chip `Equipamento #<id>` com botão para limpar.
  - Resumo: `{total} manutenções · Custo total: {formatMoeda(custo_total)}` (soma de todas as linhas filtradas, não só da página).
  - `ManutencoesLista mostrarEquipamento`; `Pagination` com `page` 1-based, `limit: pageSize`, `offset: (page - 1) * pageSize`; trocar filtro ou tamanho volta à página 1; página além da última é corrigida.
  - Botão `Nova manutenção` → `AbrirManutencaoDialog` sem `item`.

- [ ] **Step 1: Testes** (`ManutencoesPage.test.tsx`, MSW):
  - `lista as manutenções e o custo total` (`custo_total: '1239.90'` → `Custo total: R$ 1.239,90`; link do equipamento `href="/items/7?aba=manutencoes"`; `REVALLE JUAZEIRO`).
  - `filtro de status volta à página 1`: ir para a página 2 (`offset=25`), escolher `Em andamento` → última chamada com `status=em_andamento` e `offset=0`.
  - `?item_id=7 filtra por equipamento` → chamada com `item_id=7` e chip `Equipamento #7`.
  - `Nova manutenção abre o diálogo com busca de equipamento` (busca `GET /api/items` com `status=Disponível`).
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/ManutencoesPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar** a página, a rota e o item de menu.
- [ ] **Step 4: Rodar** `npx vitest run`, `npx tsc --noEmit` e `npm run build` → verdes.
- [ ] **Step 5: Commit** `feat(manutencoes): tela de manutenções com filtros e custo total`.

### Task 14: Verificação final da F4

**Files:** nenhum novo (corrigir o que a verificação apontar, na tarefa dona).

- [ ] **Step 1: Suítes completas:** `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh $WT f4 -q` → tudo verde (nenhum skip novo além dos já existentes); em `$WT/frontend`: `npx vitest run`, `npx tsc --noEmit`, `npm run build` → verdes.
- [ ] **Step 2: Migrações do zero:** `run-backend-tests.sh $WT f4zero tests/test_migracoes_f4.py tests/test_migrations.py -q` (sufixo novo = database `stock_sys_test_f4zero` vazio: baseline dos managers + todas as migrações) → verde; depois `docker exec stock_sys_mysql_test mysql -uroot -pstock_test_root_pw -e "DROP DATABASE stock_sys_test_f4zero"`. A aplicação "sobre o schema atual" (tabelas legadas com linhas) é coberta por `test_ida_e_volta_das_migracoes_f4` e pelos dois testes do down de 0402, que rodam no database `f4` já usado pela suíte.
- [ ] **Step 3: Conferir o spec F4 item a item** e anotar no relatório onde cada um está:
  - [ ] 1 — cadastro/edição com NF, fornecedor, valor, data da compra, fim da garantia, anexo da NF (PDF/imagem, 10 MB, `notas_fiscais`) e local da unidade → Tasks 3, 9 (matriz 10.1, 10.2, 8.4).
  - [ ] 2 — transferência `POST /items/{id}/transferir` (Gestor e Técnico), destino, local opcional, motivo, `revenda`/`local_id`, `transferencias`, `history` `Edição` "Transferência X → Y", auditoria → Tasks 4, 9, 11 (matriz 10.3).
  - [ ] 3 — manutenção: só sem colaborador, `Em Manutenção`, envio, previsão, retorno com custo e anexo, sem conserto sugere a baixa, retorno volta a `Disponível`, `/manutencoes` com filtros e custo por equipamento, aba "Manutenções" → Tasks 5, 12, 13 (matriz 4.1, 4.2).
  - [ ] 4 — baixa com apagamento para Notebook, Desktop, Celular, Tablet e HD; `nao_aplica` para os demais → Tasks 6, 10 (matriz 10.5).
  - [ ] 5 — `GET /items/{id}/linha-do-tempo` em ordem cronológica com cadastro, empréstimos, devoluções, manutenções, transferências e baixa → Tasks 7, 11 (matriz 10.4). Termos, inventários e linha telefônica: fontes ligadas pela F7 via `registrar_fonte` + `MODULOS_DE_FONTES` (registrar no relatório como pendência da F7).
  - [ ] Aceite: item em manutenção não aparece como disponível para empréstimo (Task 5); custo acumulado correto (Task 5); baixa de notebook sem registro de apagamento é recusada (Task 6); linha do tempo com dados de todas as tabelas da F4 (Task 7).
  - [ ] Seção 5 F4: colunas e tabelas conferidas no `information_schema` (Task 1).
  - [ ] Toda escrita nova audita: NF, transferência, 6 ações de manutenção, baixa com apagamento.
  - [ ] Telas da F4 com `StatusPill`, `Button loading`, `Pagination` 25/50/100 e `upper()`; nenhum `text-primary` novo (`grep -rn "text-primary" frontend/src/features/ciclo-vida frontend/src/features/manutencoes frontend/src/pages/ManutencoesPage.tsx` vazio).
- [ ] **Step 4: Limpeza:** nenhum container, imagem ou database de teste deixado para trás (o script usa `docker run --rm`); listar no relatório o que foi removido.
- [ ] **Step 5: Commit** (só se houve correção) `fix(ciclo-vida): ajustes da verificação final da F4`.

---

## Notas de integração (para o relatório da fase)

- `StockPage` (F1) tem `STATUS_OPTIONS` fixo sem `Em Manutenção`: a F1 deve usar `useConstants().itemStatuses` (criado aqui) ou acrescentar o valor. A F4 não edita a `StockPage`.
- `ItemDetailPage.test.tsx` afirma a lista exata de abas; toda fase que acrescenta aba (F3, F4, F5, F8) quebra esse teste. A F4 troca a asserção pela forma da Task 11; as outras fases devem fazer a mesma troca, com o mesmo texto.
- O spec da F1 cita uma aba "Linha do tempo"; pelo contexto comum ela é da F4 (ordem 20). A F1 não deve criar outra.
- `issue()` (F3) grava `items.revenda` a partir do formulário de empréstimo, o que muda a unidade sem passar por `transferencias`. Também é a checagem `status != 'Disponível'` de `issue()` que impede emprestar item em manutenção: a F3 precisa mantê-la.
- `core/storage.py` ganhou uma linha em `LocalStorage.save`; se a F3 mexer na lista de subdiretórios do `__init__`, manter as duas mudanças.
- Para a F7: alertas de manutenção parada usam `STATUS_EM_ANDAMENTO` e `aberta_em`; a exportação de manutenções usa `ManutencaoDBManager.listar(..., limit=None)`; o painel conta `Em Manutenção` por `items.status`; as fontes novas da linha do tempo entram por `registrar_fonte` num módulo acrescentado a `MODULOS_DE_FONTES`; os dados de compra podem ir para a aba Dados (F0/F1) se o Enzo preferir, hoje aparecem no topo da aba Linha do tempo.
