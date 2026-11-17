# Fase 6 — Inventário físico — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Etiquetas com QR code impressas em folha A4, atalho `/e/{id}` para o detalhe do equipamento, inventário físico por unidade com leitura pela câmera do celular, por digitação ou por leitor USB, e relatório de cada inventário na tela e em Excel/PDF.

**Architecture:** É um módulo novo e isolado. A migração `0600` cria as tabelas, as regras puras de conferência ficam em `app/services/inventario.py` e o `InventarioDBManager` usa SQL cru com `get_cursor`. Leituras, encerramento e cancelamento travam a linha do inventário com `SELECT ... FOR UPDATE`, o que serializa os técnicos que leem ao mesmo tempo. Os routers são `/api/inventarios` e `/api/etiquetas`. No frontend só há telas novas (`/etiquetas`, `/inventarios`, `/inventarios/:id`, `/inventarios/:id/leitura`, `/e/:id`), e `qrcode` e `@zxing/browser` entram por `import()` dinâmico, só quando são usados.

**Tech Stack:** FastAPI + PyMySQL/DBUtils + MySQL 8 (coluna gerada para garantir "um inventário aberto por unidade"); `app/services/exports.py` da F0 (openpyxl/fpdf2); React 18 + TS + TanStack Query 5 + Vitest/MSW; `qrcode@1.5.4` (SVG); `@zxing/browser@0.2.1`, que exige a dependência par `@zxing/library@0.23.0`.

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md`: seção 6 "F6", matriz 11.1–11.4 e 12.4, seção 5 "F6" e seções 7–9. Plano mestre: `docs/superpowers/plans/2026-11-12-patrimonio-melhorias.md`. Interfaces da F0: `docs/superpowers/plans/2026-11-12-patrimonio-f0-base.md` (Tasks 3, 8, 10–13).

**Worktree:** `<wt>` = `C:/Users/enzo.jz/dev/patrimonio/f6-inventario` (branch `fase/f6-inventario`). Backend: `RBT="bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh <wt> f6"`. Frontend: os comandos rodam em `<wt>/frontend`.

## Decisões desta fase (o executor não decide de novo)

- **Número do patrimônio** = `items.id`, exibido como `#<id>`. É a coluna "Patrimônio" que a `StockPage` já mostra. A digitação aceita `123`, `#123` e `nº 123`.
- **Endereço do QR:** o backend é a fonte única. `POST /api/etiquetas` devolve `url = APP_PUBLIC_URL.rstrip('/') + '/e/' + id`, a mesma configuração usada nos links de e-mail. O frontend não monta a URL: com `window.location.origin`, quem imprimisse acessando por IP interno ou por http geraria QR errado. A tela avisa quando a origem de `base_url` é diferente de `window.location.origin`.
- **Itens de uma unidade** = `items.is_active = 1 AND items.revenda = unidades.nome`, porque o sistema liga item e unidade pelo nome. Em Python, a comparação usa `normalizar_texto`, que tira acento e caixa, como a collation do MySQL.
- **Só um inventário `aberto` por unidade:** garantido pela coluna gerada `unidade_aberta` com `UNIQUE`.
- **Não encontrados congelados:** no encerramento, os não encontrados são gravados em `inventario_nao_encontrados`, uma terceira tabela, aditiva (o spec deixa os detalhes para o plano). Assim, o relatório de um inventário encerrado não muda quando um item é transferido ou baixado depois. Itens `Em Manutenção` (F4) são ativos e entram, com o status exibido.
- **Um resultado por leitura**, com prioridade `nao_cadastrado` > `outra_unidade` > `usuario_diferente` > `local_errado` > `ok`. O campo `detalhe` lista todas as divergências. Um item baixado lido no inventário vira `nao_cadastrado` com o detalhe "Equipamento baixado no sistema.".
- **Local encontrado:** a tela tem o seletor fixo "Estou no local", cujo valor vai em cada leitura, e cada leitura pode ser corrigida por `PATCH`. O local só diverge quando o item tem local cadastrado. O **usuário** só é conferido quando o técnico o informa (`usuario_conferido`). Item antigo sem `colaborador_id` é comparado pelo nome em `assigned_to`.
- **Duplicidade:** cada leitura recebe uma `chave`: `item:<id>`, `id:<n>` (número que não existe) ou `codigo:<TEXTO NORMALIZADO>`, com `UNIQUE(inventario_id, chave)`, além do `UNIQUE(inventario_id, item_id)` do spec. Repetir uma leitura devolve `200` com `duplicada: true` e não grava nada.
- **Papéis:** abrir, ler, corrigir, desfazer leitura e encerrar ficam com Gestor e Técnico. Cancelar é **só do Gestor**. O Jovem Aprendiz recebe 403 na API e `RestrictedAccess` na tela. A etiqueta é impressa por Gestor e Técnico.
- **`/e/:id`:** a rota fica **fora** do `AppLayout`. Sem login, guarda o destino em `sessionStorage` e manda para `/login`. O `<PosLoginRedirect />`, montado no `App.tsx`, leva ao item quando o login termina em `/`. Assim nada é editado em `LoginPage`/`AuthContext` (F2) nem em `AppLayout` (F1).
- **Bibliotecas:**
  - `qrcode@1.5.4` (MIT): é a biblioteca de QR mais usada. Tem build próprio para navegador (`browser` → `lib/browser.js`, sem `pngjs`/`yargs`) e gera **SVG** vetorial, nítido em qualquer DPI de impressora. Fica na ordem de 25–30 KB minificados e é carregada só na `/etiquetas`.
  - Alternativas descartadas: `qrcode.react` (amarra ao React sem ganho), `qrcode-generator` (API antiga, sem SVG direto) e `uqr` (pouca adoção).
  - `@zxing/browser@0.2.1` (publicado em 07/2026, mantido): `BrowserQRCodeReader.decodeFromConstraints` devolve `IScannerControls.stop()`, então a tela é nossa e o fluxo da câmera é controlado por nós.
  - `html5-qrcode` foi descartado: a última versão é a 2.3.8, de 04/2023, sem manutenção, e ela injeta UI própria que não segue os tokens.
  - O executor registra no relatório o tamanho real dos chunks gerados por `npm run build`.
- **Impressão:** a folha é renderizada num portal em `document.body` (`div.etiquetas-impressao`), com o `<style>` de impressão **dentro do portal**. Assim, `@page { size: A4; margin: 0 }` só existe enquanto a página está montada. A classe `imprimindo-etiquetas` vai no `<body>` só nesse período.
  - Na folha impressa, as cores são `black`/`white` (palavras-chave CSS). É uma exceção consciente aos tokens: o papel é branco nos dois temas.
  - O layout fica salvo em `localStorage['revalle_etiquetas_layout']`, com `try/catch`, **sem** a posição inicial, que volta a 1 a cada visita.
- **Fora desta fase:**
  - Corrigir local ou usuário do equipamento a partir do inventário: é edição do item (F4).
  - Aba "Inventário" no detalhe do equipamento: a F7 liga a fonte da linha do tempo usando `InventarioDBManager.historico_do_item`, criado na Task 5.
  - Editar a `StockPage`: a seleção para etiquetas fica na `/etiquetas`.

## Global Constraints

- Migrações só na faixa `0600–0699`, como `backend/migrations/NNNN_descricao.sql` + `NNNN_descricao.down.sql`, só aditivas.
  - O executor divide os comandos por `;` no fim da linha: nada de `;` dentro de literais nem `; -- comentário` na mesma linha.
- "O QR contém só `APP_PUBLIC_URL/e/{id}`; nenhum dado pessoal." (spec 6.F6.1 e seção 8).
- Etiquetas: "folha A4 para impressão (`@media print`) com QR code, número do patrimônio e tipo/modelo. Padrão: 3 colunas × 8 linhas (70 × 37 mm, 24 por folha), com layout configurável (colunas, linhas, margens) e posição inicial".
- "`/e/{id}` abre o detalhe do equipamento (com login)."
- Inventário: "Gestor ou Técnico abre um inventário por unidade; no celular, a tela de leitura usa a câmera (`@zxing/browser`, carregado sob demanda; exige HTTPS) ou digitação/leitor USB. Cada leitura confere unidade, local e colaborador e marca o resultado; o técnico pode informar o local onde encontrou. Encerrar o inventário calcula os 'não encontrados'."
- Modelo: `inventarios.status ENUM('aberto','encerrado','cancelado')`; `inventario_leituras.resultado ENUM('ok','local_errado','usuario_diferente','outra_unidade','nao_cadastrado')`; `UNIQUE(inventario_id, item_id)`.
- Relatório (tela e Excel/PDF por `app/services/exports.py`): "conferidos, não encontrados, local errado, usuário diferente, de outra unidade e não cadastrados".
- Aceite: "leitura duplicada não conta duas vezes; encerramento lista os não lidos; câmera negada mostra a entrada manual; impressão sai com o layout configurado."
- Auditoria: toda escrita chama `audit.record(ctx, "inventario", inventario_id, acao, antes, depois)`, com `ctx: AuditContext = Depends(get_audit_context)`.
- 403 com a mensagem padrão: `Acesso restrito ao Gestor.` / `Acesso restrito a Gestor ou Técnico.`.
- Arquivos compartilhados (`backend/app/main.py`, `backend/app/dependencies.py`, `backend/migrations/README.md`, `frontend/src/App.tsx`, `frontend/src/components/layout/navigation.ts`, `frontend/nginx.conf.template`) só recebem **linhas novas**: nada é reordenado nem reformatado.
  - Não editar `StockPage`, `ItemDetailsModal`, `AppLayout` (F1), `LoginPage`/`AuthContext` (F2), `api/items.ts` (F4), nem os componentes e registros da F0.
- Frontend:
  - Usar os tokens da F0, `StatusPill`, `Pagination` + `usePageSize` (25/50/100), `Button loading` e `upper()` para unidade.
  - Texto azul só com `text-brand-text` (nunca `text-primary`).
  - Nenhum modificador de opacidade sobre token customizado (`bg-brand/50` gera CSS vazio; há guarda em `src/lib/tokens.test.ts`).
- Dependências novas: `qrcode@1.5.4`, `@zxing/browser@0.2.1`, `@zxing/library@0.23.0` (par obrigatório de `@zxing/browser`) e, como dev, `@types/qrcode@1.5.5`. `qrcode` e `@types/qrcode` usam **exatamente** as versões que a F2 fixou para o QR do 2FA, também carregado sob demanda, para o `package.json` e o lockfile se fundirem sem conflito. Nenhuma dependência no backend. A justificativa vai no relatório da fase.
- Migrações: nenhuma tabela nova declara `CHARACTER SET`/`COLLATE`; todas usam o padrão do database, como as existentes.
- Textos, mensagens e commits em português. Commits no formato `tipo(escopo): descrição`, terminando com uma linha em branco + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus

- **A mesma etiqueta lida duas vezes** (câmera disparando a cada quadro, QR seguido de `#id` digitado, ou dois técnicos ao mesmo tempo) conta uma vez. A resposta traz `duplicada: true` e a contagem não muda. Testes: Task 4 `test_leitura_concorrente_conta_uma_vez`; Task 12, sequência repetida em menos de 3 s.
- **Equipamento de produção sem `colaborador_id`/`local_id`** (só `assigned_to`) não gera falso "usuário diferente" nem falso "local errado". Testes: Task 2, casos legados; Task 4 `test_legado_assigned_to_sem_falso_usuario_diferente`.
- **Câmera negada, página em http sem TLS ou aparelho sem câmera:** aparece a mensagem clara e o campo de digitação fica focado. A câmera é desligada ao sair da tela, ao esconder a aba e quando o início termina depois do desmonte. Testes: Tasks 12 e 13.
- **Leitura, correção ou exclusão que chega depois do encerramento** (outro técnico com a tela aberta) recebe 409 e nada é gravado. Os "não encontrados" de um inventário encerrado não mudam com transferência ou baixa posterior. Testes: Task 5.
- **Impressão:**
  - Layout que não cabe na A4 ou posição inicial fora da folha mostram erro claro e deixam "Imprimir" desabilitado.
  - Sair da tela de etiquetas não deixa `@page` nem a classe no `<body>`, porque isso deixaria em branco a impressão das outras telas.
  - Testes: Tasks 8 e 9.

---

### Task 1: Migração 0600 — tabelas do inventário

**Files:**
- Create: `backend/migrations/0600_inventarios.sql`, `backend/migrations/0600_inventarios.down.sql`, `backend/tests/test_inventario_migracao.py`
- Modify: `backend/migrations/README.md` (acrescentar a linha `| 0600-0699 | Fase 6 (inventário físico) |` na tabela de faixas)

**Interfaces:**
- Consumes: tabelas `unidades`, `items` (baseline), `locais`, `colaboradores` (F0, migrações 0001/0003); `app.db.migrations.executar_sql(cur, texto)`.
- Produces: o DDL abaixo, que é o que as Tasks 3–6 consultam.

```sql
CREATE TABLE IF NOT EXISTS inventarios (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    unidade_id INT NOT NULL,
    status ENUM('aberto','encerrado','cancelado') NOT NULL DEFAULT 'aberto',
    iniciado_por VARCHAR(100) NOT NULL,
    iniciado_em DATETIME NOT NULL,
    encerrado_por VARCHAR(100) NULL,
    encerrado_em DATETIME NULL,
    observacao VARCHAR(500) NULL,
    motivo_cancelamento VARCHAR(255) NULL,
    unidade_aberta INT AS (IF(status = 'aberto', unidade_id, NULL)) STORED,
    UNIQUE KEY uq_inventario_aberto_por_unidade (unidade_aberta),
    INDEX idx_inventarios_unidade_status (unidade_id, status),
    INDEX idx_inventarios_iniciado_em (iniciado_em),
    CONSTRAINT fk_inventarios_unidade FOREIGN KEY (unidade_id) REFERENCES unidades(id)
);
CREATE TABLE IF NOT EXISTS inventario_leituras (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    inventario_id INT NOT NULL,
    item_id INT NULL,
    chave VARCHAR(120) NOT NULL,
    codigo_lido VARCHAR(255) NOT NULL,
    lido_por VARCHAR(100) NOT NULL,
    lido_em DATETIME NOT NULL,
    local_id_encontrado INT NULL,
    usuario_conferido TINYINT(1) NOT NULL DEFAULT 0,
    colaborador_id_encontrado INT NULL,
    revenda_item VARCHAR(100) NULL,
    local_id_esperado INT NULL,
    colaborador_id_esperado INT NULL,
    usuario_esperado VARCHAR(200) NULL,
    resultado ENUM('ok','local_errado','usuario_diferente','outra_unidade','nao_cadastrado') NOT NULL,
    detalhe VARCHAR(500) NULL,
    observacao VARCHAR(500) NULL,
    atualizado_em DATETIME NULL,
    UNIQUE KEY uq_leitura_item (inventario_id, item_id),
    UNIQUE KEY uq_leitura_chave (inventario_id, chave),
    CONSTRAINT fk_leituras_inventario FOREIGN KEY (inventario_id) REFERENCES inventarios(id) ON DELETE CASCADE,
    CONSTRAINT fk_leituras_item FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE SET NULL,
    CONSTRAINT fk_leituras_local FOREIGN KEY (local_id_encontrado) REFERENCES locais(id) ON DELETE SET NULL,
    CONSTRAINT fk_leituras_colaborador FOREIGN KEY (colaborador_id_encontrado) REFERENCES colaboradores(id) ON DELETE SET NULL
);
CREATE TABLE IF NOT EXISTS inventario_nao_encontrados (
    inventario_id INT NOT NULL,
    item_id INT NOT NULL,
    status_item VARCHAR(50) NULL,
    local_id INT NULL,
    colaborador_id INT NULL,
    usuario VARCHAR(200) NULL,
    PRIMARY KEY (inventario_id, item_id),
    CONSTRAINT fk_nao_encontrados_inventario FOREIGN KEY (inventario_id) REFERENCES inventarios(id) ON DELETE CASCADE,
    CONSTRAINT fk_nao_encontrados_item FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE CASCADE
);
```

Notas sobre o DDL:
- `fk_inventarios_unidade` não pode ter `CASCADE`/`SET NULL`: o MySQL proíbe essas ações na coluna-base de uma coluna gerada `STORED`.
- `local_id_esperado`, `colaborador_id_esperado` e `usuario_esperado` são fotografias tiradas no momento da leitura, por isso ficam sem FK.
- As tabelas não declaram `CHARACTER SET` nem `COLLATE`: herdam o padrão do database, como as tabelas legadas. Nunca usar collation exclusiva do MySQL 8 (ex.: `utf8mb4_0900_ai_ci`), e as comparações `items.revenda = unidades.nome` continuam na mesma collation.
- O `.down.sql` apaga as tabelas nesta ordem: `DROP TABLE IF EXISTS inventario_nao_encontrados;`, depois `inventario_leituras;`, depois `inventarios;`.

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_inventario_migracao.py` (marcador `integration`). Os imports de `app.*` ficam dentro dos testes; as leituras usam `from app.db.database_mysql import get_cursor` com `get_cursor(commit=False)`. Testes:
  - `test_tabelas_e_colunas_existem`: `information_schema.columns` traz todas as colunas acima nas três tabelas.
  - `test_um_inventario_aberto_por_unidade`: com a fixture `executar_sql`, insere um `aberto` para uma unidade (fixture `unidade`) e afirma que um segundo `aberto` na mesma unidade levanta `pymysql.err.IntegrityError`. Depois de `UPDATE ... SET status='encerrado'`, o segundo `INSERT` passa.
  - `test_leitura_unica_por_chave_e_por_item`: repetir `(inventario_id, chave)` levanta `IntegrityError`; duas leituras com `item_id NULL` e chaves diferentes passam.
  - `test_down_e_up`: lê `0600_inventarios.down.sql` e o `.sql`, aplica o down e depois o up com `app.db.migrations.executar_sql(cur, texto)` num `get_cursor()` (o up vai num `finally`) e afirma que as três tabelas existem de novo.
- [ ] **Step 2: Rodar e ver falhar:** `$RBT tests/test_inventario_migracao.py -q` → FAIL (tabelas inexistentes).
- [ ] **Step 3: Implementar** os dois arquivos de migração e a linha no README.
- [ ] **Step 4: Rodar** `$RBT tests/test_inventario_migracao.py tests/test_migrations.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(inventario): tabelas de inventário, leituras e não encontrados (migração 0600)`.

### Task 2: Regras puras da conferência

**Files:**
- Create: `backend/app/services/inventario.py`, `backend/tests/test_inventario_regras.py` (marcador `unit`)

**Interfaces:**
- Produces (em `app/services/inventario.py`):
  - `Resultado = Literal["ok", "local_errado", "usuario_diferente", "outra_unidade", "nao_cadastrado"]`; `Categoria = Literal[..., "nao_encontrado"]` (os 5 acima + `nao_encontrado`); `Divergencia = Literal["baixado", "sem_unidade", "outra_unidade", "usuario", "local"]`.
  - `ROTULOS_CATEGORIA = {"ok": "Conferido", "local_errado": "Local errado", "usuario_diferente": "Usuário diferente", "outra_unidade": "De outra unidade", "nao_cadastrado": "Não cadastrado", "nao_encontrado": "Não encontrado"}`; `ROTULO_AINDA_NAO_LIDO = "Ainda não lido"`.
  - `ORDEM_CATEGORIAS = ("nao_encontrado", "nao_cadastrado", "outra_unidade", "usuario_diferente", "local_errado", "ok")`.
  - `class CodigoInvalido(ValueError)`.
  - `@dataclass(frozen=True) class Alvo: tipo: Literal["id", "identificador"]; valor: int | str`.
  - `resolver_codigo(texto: str) -> Alvo`, que trata o texto nesta ordem:
    1. `strip()`; vazio levanta `CodigoInvalido("Informe o código lido.")`.
    2. URL da etiqueta: `re.search(r"[/;]e[/;](\d{1,10})[/;]?\s*$", t, re.I)` → `Alvo("id", int)`. O `;` cobre o leitor USB em configuração americana digitando num Windows ABNT2, que transforma `https://host/e/12` em `httpsÇ;;host;e;12`.
    3. Número digitado: `re.fullmatch(r"(?:#|n[º°o.]?)?\s*(\d{1,10})", t, re.I)` → `Alvo("id", int)`.
    4. Qualquer outro texto → `Alvo("identificador", " ".join(t.split()).upper())`.
  - `chave_da_leitura(alvo: Alvo, item_id: int | None) -> str`: `item:<item_id>` quando o item existe; senão `id:<n>` (alvo numérico) ou `codigo:<valor>`, cortada em 120 caracteres.
  - `normalizar_texto(v: str | None) -> str`: NFD sem marcas combinantes, `casefold()`, espaços colapsados, `strip()`; `None` → `""`.
  - `@dataclass(frozen=True) class Esperado: item_id: int; ativo: bool; revenda: str | None; local_id: int | None; colaborador_id: int | None; usuario_nome: str | None` (`usuario_nome` = nome do colaborador vinculado ou, sem vínculo, `items.assigned_to`).
  - `@dataclass(frozen=True) class Encontrado: local_id: int | None; usuario_conferido: bool; colaborador_id: int | None; colaborador_nome: str | None`.
  - `classificar(esperado: Esperado | None, unidade_nome: str, encontrado: Encontrado) -> tuple[Resultado, list[Divergencia]]`, com as regras:
    - `None` → `("nao_cadastrado", [])`.
    - `not ativo` → `("nao_cadastrado", ["baixado"])`.
    - Senão as divergências são acumuladas nesta ordem:
      - `"sem_unidade"` se `revenda` vazia.
      - `"outra_unidade"` se `normalizar_texto(revenda) != normalizar_texto(unidade_nome)`.
      - `"usuario"` se `usuario_conferido` e o usuário diverge. Divergir significa:
        - com `colaborador_id` esperado: o id encontrado é outro;
        - sem `colaborador_id`, mas com `usuario_nome`: ninguém encontrado ou nome normalizado diferente;
        - sem ninguém esperado: alguém foi encontrado.
      - `"local"` se não há divergência de unidade e os dois `local_id` estão preenchidos e são diferentes.
    - Resultado: `outra_unidade` se houver `sem_unidade`/`outra_unidade`; senão `usuario_diferente` se houver `usuario`; senão `local_errado` se houver `local`; senão `ok`.
  - `montar_detalhe(divergencias: list[Divergencia], *, revenda: str | None, local_esperado_nome: str | None, local_encontrado_nome: str | None, usuario_esperado: str | None, usuario_encontrado: str | None) -> str | None`. As frases são unidas por espaço; sem divergências devolve `None`:
    - `baixado` → `"Equipamento baixado no sistema."`
    - `sem_unidade` → `"Equipamento sem unidade cadastrada."`
    - `outra_unidade` → `f"Cadastrado na unidade {revenda.upper()}."`
    - `usuario` → `f"Usuário cadastrado: {usuario_esperado or 'nenhum'} · encontrado com: {usuario_encontrado or 'ninguém'}."`
    - `local` → `f"Local cadastrado: {local_esperado_nome} · encontrado em: {local_encontrado_nome}."`

- [ ] **Step 1: Escrever os testes** (`@pytest.mark.unit`):
  - `test_resolver_codigo` (parametrizado):
    - `"https://patrimonio.revalle.com.br/e/123"` → `Alvo("id", 123)`
    - `"httpsÇ;;patrimonio.revalle.com.br;e;123"` → `Alvo("id", 123)`
    - `"HTTPS://X/E/45/"` → `45`
    - `"#77"` → `77`; `" nº 8 "` → `8`; `"77"` → `77`
    - `"SN-E-1234"` → `Alvo("identificador", "SN-E-1234")`
    - `"abc  123"` → `Alvo("identificador", "ABC 123")`
  - `test_codigo_vazio_levanta`: `pytest.raises(CodigoInvalido, match="Informe o código lido.")` para `"   "`.
  - `test_chave_da_leitura`: `"item:5"`, `"id:99999"`, `"codigo:XYZ"`.
  - `test_classificar`, parametrizado com `unidade_nome="Revalle Juazeiro"`; nas linhas abaixo, "e" é o `Esperado` e "en" o `Encontrado`:
    - `None` → `("nao_cadastrado", [])`
    - inativo → `("nao_cadastrado", ["baixado"])`
    - revenda `"REVALLE JUAZEIRO "` → `ok`
    - revenda `"Petrolina"` → `("outra_unidade", ["outra_unidade"])`
    - revenda `None` → `("outra_unidade", ["sem_unidade"])`
    - local e=1 / en=2 → `("local_errado", ["local"])`
    - **local e=None / en=2 → `ok`**
    - local e=1 / en=None → `ok`
    - outra unidade com locais diferentes → `("outra_unidade", ["outra_unidade"])`
    - `usuario_conferido=False` com colaborador diferente → `ok`
    - colaborador e=5 / en=6 → `usuario_diferente`
    - **legado:** `colaborador_id=None`, `usuario_nome="Fulano de Tal"`, en `colaborador_nome="FULANO DE TAL"` → `ok`; en `"Beltrano"` → `usuario_diferente`; en ninguém → `usuario_diferente`
    - sem usuário esperado e en ninguém → `ok`; en colaborador 7 → `usuario_diferente`
    - usuário e local divergentes → `("usuario_diferente", ["usuario", "local"])`
  - `test_montar_detalhe`:
    - `["usuario", "local"]` com `"Fulano"`/`None` e `"Sala TI"`/`"Recepção"` → `"Usuário cadastrado: Fulano · encontrado com: ninguém. Local cadastrado: Sala TI · encontrado em: Recepção."`
    - `["outra_unidade"]` com `revenda="Juazeiro"` → `"Cadastrado na unidade JUAZEIRO."`
    - `[]` → `None`
- [ ] **Step 2: Rodar e ver falhar:** `$RBT tests/test_inventario_regras.py -q` → FAIL (módulo inexistente).
- [ ] **Step 3: Implementar** `app/services/inventario.py`, só com stdlib (`re`, `unicodedata`, `dataclasses`).
- [ ] **Step 4: Rodar** → PASS.
- [ ] **Step 5: Commit** `feat(inventario): regras de leitura e conferência do inventário`.

### Task 3: Abrir, listar, consultar e cancelar inventários (API)

**Files:**
- Create: `backend/app/db/inventario_db.py`, `backend/app/schemas/inventarios.py`, `backend/app/routers/inventarios.py`, `backend/tests/test_inventarios.py`
- Modify:
  - `backend/app/dependencies.py`: acrescentar o import de `InventarioDBManager` e `get_inventario_db()` com `@lru_cache`.
  - `backend/app/main.py`: acrescentar a linha `from app.routers import inventarios`, `app.include_router(inventarios.router)` e, em `_ROTULOS_DE_CAMPO`, só as chaves `"unidade_id": "Unidade"`, `"observacao": "Observação"`, `"motivo": "Motivo"`, `"codigo": "Código lido"`, `"local_id_encontrado": "Local"`, `"colaborador_id_encontrado": "Colaborador"`, `"usuario_conferido": "Usuário conferido"`, `"item_ids": "Equipamentos"`.

**Interfaces:**
- Consumes: `get_cursor` (`app.db.database_mysql`); `audit.record`, `AuditContext`, `get_audit_context`, `gestor_or_tecnico`, `gestor_only`, `Paginated`.
- Produces:
  - Em `inventario_db.py`:
    - Exceções: `InventarioNaoEncontrado(Exception)`, `InventarioFechado(Exception)` (a mensagem vem no `raise`), `InventarioJaAberto(Exception)` (atributo `inventario_id: int`), `UnidadeNaoEncontrada(Exception)`, `LeituraNaoEncontrada(Exception)`.
    - `class InventarioDBManager` (sem efeito no `__init__`; as tabelas vêm da migração), com:
      - `_travar(cur, inventario_id: int) -> dict`: `SELECT i.*, u.nome AS unidade_nome FROM inventarios i JOIN unidades u ON u.id = i.unidade_id WHERE i.id = %s FOR UPDATE`; sem linha levanta `InventarioNaoEncontrado`.
      - `abrir(unidade_id: int, observacao: str | None, por: str) -> int`: unidade inexistente levanta `UnidadeNaoEncontrada`; unidade inativa levanta `ValueError("Unidade inativa não pode ser inventariada.")`; erro 1062 no `INSERT` levanta `InventarioJaAberto(id do aberto)`.
      - `obter(inventario_id: int) -> dict | None`: `inventarios.*` + `unidade_nome` + `total_leituras`.
      - `listar(*, unidade_id: int | None, status: str | None, limit: int, offset: int) -> tuple[list[dict], int]`, ordenado por `iniciado_em DESC, id DESC`.
      - `resumo(inventario: dict) -> dict` com as chaves `esperados, lidos, encontrados, ok, local_errado, usuario_diferente, outra_unidade, nao_cadastrado, nao_encontrados`:
        - `lidos` = total de leituras; contagem por `resultado`; `encontrados = ok + local_errado + usuario_diferente`.
        - `aberto`: `esperados` = itens ativos da unidade agora; `nao_encontrados` = itens ativos da unidade sem leitura neste inventário.
        - `encerrado`: `nao_encontrados` = linhas de `inventario_nao_encontrados`; `esperados = encontrados + nao_encontrados`.
        - `cancelado`: `nao_encontrados = 0`; `esperados = encontrados`.
      - `cancelar(inventario_id: int, motivo: str, por: str) -> None`: status diferente de `aberto` levanta `InventarioFechado("Só inventários abertos podem ser cancelados.")`; grava `status='cancelado'`, `encerrado_por`, `encerrado_em=datetime.now()` e `motivo_cancelamento`.
  - Em `schemas/inventarios.py`:
    - `InventarioCreate(unidade_id: int, observacao: Optional[str] = None)`: `strip`; `""` vira `None`; mais de 500 caracteres → `ValueError("Observação muito longa (máximo 500 caracteres).")`.
    - `CancelarInventario(motivo: str)`: `strip`; menos de 3 caracteres → `"Informe o motivo do cancelamento."`; mais de 255 → `"Motivo muito longo (máximo 255 caracteres)."`.
    - `ResumoInventario` (as 9 chaves acima, `int`).
    - `InventarioResponse(id: int, unidade_id: int, unidade_nome: str, status: Literal["aberto", "encerrado", "cancelado"], iniciado_por: str, iniciado_em: datetime, encerrado_por: Optional[str] = None, encerrado_em: Optional[datetime] = None, observacao: Optional[str] = None, motivo_cancelamento: Optional[str] = None, total_leituras: int = 0)`.
    - `InventarioDetalhe(InventarioResponse)` + `resumo: ResumoInventario`.
  - Router `APIRouter(prefix="/api/inventarios", tags=["inventarios"])`, tudo com `gestor_or_tecnico` salvo onde indicado:
    - `GET ""` (`unidade_id`, `status: Optional[Literal[...]]`, `limit: int = Query(25, ge=1, le=100)`, `offset`) → `Paginated[InventarioResponse]`.
    - `POST ""` (`status_code=201`) → `InventarioDetalhe`.
    - `GET "/{inventario_id}"` → `InventarioDetalhe`.
    - `POST "/{inventario_id}/cancelar"` (**`gestor_only`**) → `InventarioDetalhe`.
    - Mapeamento de erros:
      - `InventarioNaoEncontrado` → 404 `"Inventário não encontrado."`
      - `UnidadeNaoEncontrada` → 404 `"Unidade não encontrada."`
      - `ValueError` → 422 com a mensagem
      - `InventarioJaAberto` → 409 `f"Já existe um inventário aberto para esta unidade (#{e.inventario_id})."`
      - `InventarioFechado` → 409 com a mensagem
    - Auditoria:
      - Abrir: `acao="criar"`, `depois={"unidade_id", "unidade_nome", "observacao"}`.
      - Cancelar: `acao="cancelar"`, `antes={"status": "aberto"}`, `depois={"status": "cancelado", "motivo": motivo}`.

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_inventarios.py`. Os itens vêm de `criar_item(revenda=unidade["nome"])`; a baixa é feita com `executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (id,))`. Testes:
  - `test_abrir_inventario_tecnico`: com 2 itens ativos na unidade, 1 baixado e 1 de outra unidade, `POST {"unidade_id": u}` → 201, `status == "aberto"`, `iniciado_por` = usuário do técnico, `resumo == {"esperados": 2, "lidos": 0, "encontrados": 0, "ok": 0, ..., "nao_encontrados": 2}`.
  - `test_segundo_aberto_na_mesma_unidade_409`: `detail == f"Já existe um inventário aberto para esta unidade (#{primeiro})."`.
  - `test_cancelar_libera_novo_inventario`.
  - `test_unidade_inexistente_404_e_inativa_422`.
  - `test_aprendiz_recebe_403`: `GET` e `POST` → 403 `"Acesso restrito a Gestor ou Técnico."`.
  - `test_cancelar_so_gestor`:
    - técnico → 403 `"Acesso restrito ao Gestor."`;
    - gestor com `{"motivo": "Contagem refeita"}` → 200, `status == "cancelado"`, `motivo_cancelamento == "Contagem refeita"`;
    - cancelar de novo → 409 `"Só inventários abertos podem ser cancelados."`;
    - `{"motivo": " "}` → 422 `"Motivo: Informe o motivo do cancelamento."`.
  - `test_listar_filtra_e_pagina`: com 3 inventários em 2 unidades, filtrar por `unidade_id` e por `status`; `total` correto; mais recente primeiro; `limit=101` → 422.
  - `test_inventario_inexistente_404`.
  - `test_auditoria_criar_e_cancelar`: há linhas em `audit_log` com `entidade="inventario"`, `entidade_id=str(id)`, `acao` `criar` e `cancelar` e o `username` do token.
- [ ] **Step 2: Rodar e ver falhar:** `$RBT tests/test_inventarios.py -q` → FAIL.
- [ ] **Step 3: Implementar** o manager, os schemas, o router e as linhas em `dependencies.py`/`main.py`.
- [ ] **Step 4: Rodar** `$RBT tests/test_inventarios.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(inventario): abrir, listar, consultar e cancelar inventários`.

### Task 4: Leituras (registrar, corrigir, desfazer, listar)

**Files:**
- Modify: `backend/app/db/inventario_db.py`, `backend/app/schemas/inventarios.py`, `backend/app/routers/inventarios.py`
- Test: `backend/tests/test_inventario_leituras.py`

**Interfaces:**
- Consumes: `resolver_codigo`, `chave_da_leitura`, `classificar`, `montar_detalhe`, `Esperado`, `Encontrado`, `CodigoInvalido` (Task 2); `_travar`, as exceções e `resumo` (Task 3).
- Produces:
  - Leitura (`dict` e `LeituraResponse`) com os campos:
    - da leitura: `id, inventario_id, item_id, codigo_lido, lido_por, lido_em, resultado, detalhe, observacao, atualizado_em`;
    - do que foi encontrado: `local_id_encontrado, local_encontrado_nome, usuario_conferido: bool, colaborador_id_encontrado, colaborador_encontrado_nome`;
    - do que era esperado: `local_id_esperado, local_esperado_nome, usuario_esperado`;
    - do item: `item_tipo, item_brand, item_model, item_identificador, item_revenda` (= `revenda_item`).
    - `LeituraResponse` acrescenta `duplicada: bool = False`.
  - `obter_leitura(leitura_id: int) -> dict | None`: com `LEFT JOIN` em `items`, `locais` (encontrado e esperado) e `colaboradores`.
  - `registrar_leitura(inventario_id: int, *, codigo: str, local_id_encontrado: int | None, colaborador_id_encontrado: int | None, usuario_conferido: bool, observacao: str | None, por: str) -> tuple[dict, bool]` devolve `(leitura, duplicada)`. Tudo acontece num único `get_cursor()`:
    1. `_travar` e status `aberto`; caso contrário, `InventarioFechado("Este inventário não está aberto: não aceita leituras.")`.
    2. Valida o local: precisa ser `locais.id` com `unidade_id` igual ao do inventário, senão `ValueError("O local informado não pertence à unidade do inventário.")`.
    3. Valida o colaborador: precisa existir, senão `ValueError("Colaborador não encontrado.")`.
    4. `resolver_codigo` (`CodigoInvalido` vira 422).
    5. Busca o item **inclusive inativo**, por `id` ou por `UPPER(TRIM(identificador)) = %s`. Mais de um por identificador → `ValueError("Há mais de um equipamento com esse número de série. Leia o QR code ou digite o número do patrimônio.")`.
    6. Calcula a `chave` e procura uma leitura existente com essa chave. Se existe, devolve `(leitura, True)`.
    7. Monta `Esperado`; `usuario_nome` = nome do colaborador do item, ou `assigned_to` se não houver.
    8. Roda `classificar` e `montar_detalhe` (o detalhe usa os nomes dos locais e do colaborador).
    9. `INSERT`, gravando também a fotografia do esperado: `revenda_item`, `local_id_esperado`, `colaborador_id_esperado`, `usuario_esperado`. Erro 1062 aqui (defesa) → relê pela chave e devolve `(existente, True)`.
    10. Devolve `(obter_leitura(id), False)`.
  - `atualizar_leitura(inventario_id: int, leitura_id: int, mudancas: dict, por: str) -> tuple[dict, dict]` devolve `(antes, depois)`:
    - Trava o inventário e exige que esteja aberto.
    - A leitura precisa ser deste inventário, senão `LeituraNaoEncontrada`.
    - Valida local e colaborador como no registro.
    - Chave ausente em `mudancas` não muda o campo. `colaborador_id_encontrado` presente sem `usuario_conferido` implica `usuario_conferido=True`. `usuario_conferido=False` zera `colaborador_id_encontrado`.
    - Reclassifica com a fotografia: `Esperado(item_id, ativo=True, revenda=revenda_item, local_id=local_id_esperado, colaborador_id=colaborador_id_esperado, usuario_nome=usuario_esperado)`. Uma leitura `nao_cadastrado` continua `nao_cadastrado`.
    - Grava `atualizado_em`.
  - `excluir_leitura(inventario_id: int, leitura_id: int) -> dict`: trava, exige aberto e devolve a leitura apagada.
  - `listar_leituras(inventario_id: int, *, resultado: str | None, limit: int, offset: int) -> tuple[list[dict], int]`, ordenado por `lido_em DESC, id DESC`.
  - Schemas:
    - `LeituraCreate(codigo: str, local_id_encontrado: Optional[int] = None, colaborador_id_encontrado: Optional[int] = None, usuario_conferido: bool = False, observacao: Optional[str] = None)`. O `codigo` passa por `strip`; vazio → `"Informe o código lido."`; mais de 255 caracteres → `"Código muito longo (máximo 255 caracteres)."`. `observacao` tem no máximo 500 caracteres.
    - `LeituraUpdate` com os mesmos campos menos `codigo`, todos opcionais; o router usa `model_dump(exclude_unset=True)`.
    - `LeituraResponse`.
  - Endpoints:
    - `GET "/{inventario_id}/leituras"` (`resultado: Optional[Resultado]`, `limit` 25 com `le=100`, `offset`) → `Paginated[LeituraResponse]`; inventário inexistente → 404.
    - `POST "/{inventario_id}/leituras"` → `LeituraResponse`: 201 para leitura nova, **200** para duplicada (por um parâmetro `response: Response`).
    - `PATCH "/{inventario_id}/leituras/{leitura_id}"` → `LeituraResponse`.
    - `DELETE "/{inventario_id}/leituras/{leitura_id}"` → `{"detail": "Leitura excluída."}`.
    - Erros: `InventarioFechado` → 409; `LeituraNaoEncontrada` → 404 `"Leitura não encontrada."`; `ValueError`/`CodigoInvalido` → 422 com a mensagem.
  - Auditoria (`entidade="inventario"`, `entidade_id=inventario_id`):
    - Registrar: `acao="registrar_leitura"`, `depois={"leitura_id", "item_id", "codigo_lido", "resultado"}`. Leitura duplicada **não** é auditada.
    - Corrigir: `"editar_leitura"`, com `audit.diff` sobre `local_id_encontrado, usuario_conferido, colaborador_id_encontrado, observacao, resultado` e `leitura_id` nos dois lados.
    - Desfazer: `"excluir_leitura"`, `antes={"leitura_id", "item_id", "codigo_lido", "resultado"}`.

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_inventario_leituras.py`:
  - Helpers no próprio arquivo: `_criar_local(unidade_id, nome) -> int` e `_criar_colaborador(nome, cpf) -> int`. Ambos usam `INSERT` direto com `get_cursor()` (import dentro da função) e devolvem `lastrowid`; o colaborador vai com `origem='manual'`, `situacao='ativo'`, `criado_em/atualizado_em=NOW()`.
  - `test_leitura_pelo_qr_confere_ok`: item da unidade com `local_id=L1`, `POST {"codigo": "http://localhost:5173/e/<id>", "local_id_encontrado": L1}` → 201, `resultado == "ok"`, `item_tipo == "Notebook"`, `duplicada is False`.
  - `test_leitura_duplicada_nao_conta_duas_vezes`: lê pela URL e depois `"#<id>"`; a segunda volta 200 com `duplicada is True` e o mesmo `id`; `GET /api/inventarios/{i}` → `resumo["lidos"] == 1`; só 1 linha `registrar_leitura` em `audit_log`.
  - `test_leitura_concorrente_conta_uma_vez`: `ThreadPoolExecutor(2)` chamando `InventarioDBManager().registrar_leitura(...)` com o mesmo código → `sorted(flags) == [False, True]` e `COUNT(*) == 1`.
  - `test_local_errado_com_detalhe`: locais "Sala TI" (no item) e "Recepção" (encontrado) → `local_errado`, `detalhe == "Local cadastrado: Sala TI · encontrado em: Recepção."`.
  - `test_item_de_outra_unidade`: `outra_unidade`, `detalhe == f"Cadastrado na unidade {outra['nome'].upper()}."`.
  - `test_nao_cadastrado_conta_uma_vez`:
    - `"XYZ-999"` duas vezes → 201 com `item_id is None`, depois 200 duplicada;
    - `"#99999"` e depois `"99999"` → a segunda é duplicada.
  - `test_item_baixado_vira_nao_cadastrado`: `detalhe == "Equipamento baixado no sistema."` e `item_id` preenchido.
  - `test_numero_de_serie_ambiguo_422`: dois itens com o mesmo `identificador`.
  - `test_codigo_vazio_422`: `detail == "Código lido: Informe o código lido."`.
  - `test_local_de_outra_unidade_422`: `"O local informado não pertence à unidade do inventário."`.
  - `test_informar_usuario_por_patch`: item com colaborador A.
    - `PATCH {"colaborador_id_encontrado": B}` → `usuario_diferente`, `usuario_conferido is True`, `colaborador_encontrado_nome == "B"`.
    - `PATCH {"usuario_conferido": false}` → `ok` e `colaborador_id_encontrado is None`.
  - `test_legado_assigned_to_sem_falso_usuario_diferente`: item com `assigned_to="Fulano de Tal"` e `colaborador_id NULL`.
    - `PATCH` com colaborador de nome `"FULANO DE TAL"` → `ok`.
    - `PATCH {"usuario_conferido": true, "colaborador_id_encontrado": null}` → `usuario_diferente`.
  - `test_desfazer_leitura_permite_ler_de_novo`: `DELETE` → 200 `{"detail": "Leitura excluída."}`; a nova leitura volta 201 com `duplicada is False`; existe auditoria `excluir_leitura`.
  - `test_listar_leituras_recentes_primeiro_e_por_resultado`.
  - `test_aprendiz_403_nas_leituras`.
- [ ] **Step 2: Rodar e ver falhar:** `$RBT tests/test_inventario_leituras.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `$RBT tests/test_inventario_leituras.py tests/test_inventarios.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(inventario): leituras com conferência de unidade, local e usuário`.

### Task 5: Encerramento, relatório e exportação Excel/PDF

**Files:**
- Modify: `backend/app/db/inventario_db.py`, `backend/app/schemas/inventarios.py`, `backend/app/routers/inventarios.py`
- Test: `backend/tests/test_inventario_relatorio.py`

**Interfaces:**
- Consumes: `exports.Coluna`, `gerar_excel(titulo, colunas, linhas, filtros)`, `gerar_pdf(titulo, colunas, linhas, filtros, paisagem=True)`, `resposta_arquivo(conteudo, nome_base, formato)` (F0 Task 8); `ROTULOS_CATEGORIA`, `ROTULO_AINDA_NAO_LIDO`, `ORDEM_CATEGORIAS` (Task 2).
- Produces:
  - `encerrar(inventario_id: int, por: str) -> None`, numa única transação:
    - `_travar`; status diferente de `aberto` → `InventarioFechado("Só inventários abertos podem ser encerrados.")`.
    - Grava a fotografia: `INSERT INTO inventario_nao_encontrados (inventario_id, item_id, status_item, local_id, colaborador_id, usuario) SELECT %s, i.id, i.status, i.local_id, i.colaborador_id, COALESCE(c.nome, i.assigned_to) FROM items i LEFT JOIN colaboradores c ON c.id = i.colaborador_id WHERE i.is_active = 1 AND i.revenda = %s AND NOT EXISTS (SELECT 1 FROM inventario_leituras l WHERE l.inventario_id = %s AND l.item_id = i.id)`.
    - `UPDATE` para `status='encerrado'`, `encerrado_por`, `encerrado_em`.
  - `linhas_relatorio(inventario: dict) -> list[dict]`. Chaves (= `RelatorioLinha`): `categoria, leitura_id, item_id, tipo, brand, model, identificador, status_item, unidade_item, local_esperado, local_encontrado, usuario_esperado, usuario_encontrado, codigo_lido, lido_por, lido_em, detalhe, observacao`. De onde vêm as linhas:
    - Leituras: `categoria = resultado`, `unidade_item = revenda_item`.
    - Não encontrados: `categoria = "nao_encontrado"`.
      - `aberto`: calculados na hora, com os itens ativos da unidade sem leitura.
      - `encerrado`: da fotografia, com `JOIN` em `items`/`locais` só para tipo, marca, modelo, série e nome do local.
      - `cancelado`: nenhum.
    - Ordem: posição em `ORDEM_CATEGORIAS`; depois `item_id` (os `None` por último); depois `leitura_id`.
  - `historico_do_item(item_id: int) -> list[dict]`, para a F7 ligar na linha do tempo. Chaves: `inventario_id, unidade_nome, status_inventario, evento ("leitura" | "nao_encontrado"), resultado, data, por, local_encontrado_nome`. Exclui inventários `cancelado`; `data` = `lido_em` ou `encerrado_em`; ordem crescente por `data`.
  - Schemas: `RelatorioLinha` e `RelatorioInventario(inventario: InventarioDetalhe, linhas: list[RelatorioLinha])`.
  - Endpoints:
    - `POST "/{inventario_id}/encerrar"` → `InventarioDetalhe`. Erro → 409 `"Só inventários abertos podem ser encerrados."`. Auditoria `acao="encerrar"`, `depois={"status": "encerrado", **resumo}`.
    - `GET "/{inventario_id}/relatorio"` → `RelatorioInventario`.
    - `GET "/{inventario_id}/relatorio/exportar"` (`formato: Literal["xlsx", "pdf"]`, sem padrão) → arquivo.
  - Exportação:
    - `titulo = f"Inventário #{id} — {unidade_nome.upper()}"`, com `" (parcial)"` no fim quando aberto.
    - `filtros = {"Situação": "Aberto" | "Encerrado" | "Cancelado", "Aberto em": "dd/mm/aaaa hh:mm", "Encerrado em": idem ou "—", "Resumo": "Conferidos N · Não encontrados N · Local errado N · Usuário diferente N · De outra unidade N · Não cadastrados N"}`.
    - Cada linha ganha:
      - `categoria_rotulo` = `ROTULOS_CATEGORIA`, ou `ROTULO_AINDA_NAO_LIDO` para `nao_encontrado` com o inventário aberto;
      - `patrimonio` = `f"#{item_id}"` ou `"—"`;
      - `equipamento` = `"marca modelo"`;
      - `lido_em` já formatado como `"%d/%m/%Y %H:%M"` (coluna de texto).
    - `COLUNAS_EXCEL`: Resultado, Patrimônio, Tipo, Marca/Modelo, Nº de série, Unidade do equipamento (`maiusculas=True`), Local cadastrado, Local encontrado, Usuário cadastrado, Usuário encontrado, Lido por, Lido em, Código lido, Detalhe, Observação.
    - `COLUNAS_PDF`: Resultado, Patrimônio, Tipo, Marca/Modelo, Nº de série, Local cadastrado, Local encontrado, Usuário cadastrado, Usuário encontrado, Detalhe.
    - `nome_base=f"inventario_{id}"`.

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_inventario_relatorio.py`:
  - `test_encerrar_congela_nao_encontrados`: unidade com itens A, B e C ativos e D baixado.
    - Lê A e encerra → `resumo["nao_encontrados"] == 2`, `["esperados"] == 3`, `["encontrados"] == 1`.
    - `{l["item_id"] for l in relatorio["linhas"] if l["categoria"] == "nao_encontrado"} == {B, C}`.
    - Depois `UPDATE items SET revenda='Outra' WHERE id=C` e baixa de B → o relatório traz as mesmas linhas.
  - `test_escritas_depois_de_encerrar_409`: `POST` leitura → 409 `"Este inventário não está aberto: não aceita leituras."`; `PATCH` e `DELETE` → 409; `COUNT(*)` de leituras não muda.
  - `test_encerrar_duas_vezes_e_cancelado_409`.
  - `test_relatorio_aberto_lista_ainda_nao_lidos`: um item novo cadastrado na unidade depois da abertura aparece como `nao_encontrado`.
  - `test_relatorio_ordena_por_categoria`.
  - `test_exportar_excel`: 200, `content-type` `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`. Com `openpyxl.load_workbook(BytesIO(resp.content))`, a planilha contém `"Não encontrado"`, `f"#{B}"` e o título com `unidade["nome"].upper()`.
  - `test_exportar_pdf`: `resp.content[:4] == b"%PDF"`.
  - `test_exportar_formato_invalido_422`: `?formato=csv` → 422.
  - `test_historico_do_item`: lido no inventário 1 (encerrado) e não lido no inventário 2 (encerrado) → `[e["evento"] for e in historico] == ["leitura", "nao_encontrado"]`.
  - `test_auditoria_encerrar`: há `acao="encerrar"` e `depois["nao_encontrados"] == 2`.
- [ ] **Step 2: Rodar e ver falhar:** `$RBT tests/test_inventario_relatorio.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `$RBT tests/test_inventario_relatorio.py tests/test_inventario_leituras.py tests/test_inventarios.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(inventario): encerramento com não encontrados e relatório com exportação Excel/PDF`.

### Task 6: Etiquetas (API)

**Files:**
- Create: `backend/app/routers/etiquetas.py`, `backend/tests/test_etiquetas.py`
- Modify: `backend/app/db/inventario_db.py`, `backend/app/schemas/inventarios.py`, `backend/app/main.py` (acrescentar a linha `from app.routers import etiquetas` e `app.include_router(etiquetas.router)`)

**Interfaces:**
- Consumes: `InventoryDBManager.list_items(*, tipo, status, revenda, search, limit=None, offset=0) -> (rows, total)` (`get_inventory_db`); `settings.APP_PUBLIC_URL`.
- Produces:
  - `InventarioDBManager.dados_etiquetas(item_ids: list[int]) -> list[dict]`: `id, tipo, brand, model, identificador, revenda`, com `is_active = 1` e `ORDER BY id`.
  - Schemas:
    - `EtiquetasRequest(item_ids: list[int])`: duplicados removidos. Vazio → `"Selecione ao menos um equipamento."`; mais de 2000 → `"Selecione no máximo 2000 equipamentos por vez."`.
    - `Etiqueta(id: int, tipo: Optional[str], brand: Optional[str], model: Optional[str], identificador: Optional[str], revenda: Optional[str], url: str)`.
    - `EtiquetasResponse(base_url: str, etiquetas: list[Etiqueta], ignorados: list[int])`.
    - `IdsEtiquetas(ids: list[int], total: int)`.
  - Router `APIRouter(prefix="/api/etiquetas", tags=["etiquetas"])`, `gestor_or_tecnico`:
    - `GET "/ids"` (`revenda`, `tipo`, `status`, `search`, todos opcionais na assinatura) → `IdsEtiquetas`, com os ids em ordem crescente. `revenda` vazia ou ausente → 422 `"Escolha a unidade."`.
    - `POST ""` → `EtiquetasResponse`:
      - `base_url = settings.APP_PUBLIC_URL.rstrip("/")`;
      - `url = f"{base_url}/e/{id}"`;
      - `ignorados` = ids pedidos e não devolvidos (inexistentes ou baixados), em ordem crescente.
    - Nenhum outro campo do item sai na resposta. Não há auditoria (não é escrita).

- [ ] **Step 1: Escrever os testes** em `backend/tests/test_etiquetas.py`:
  - `test_url_do_qr_so_tem_endereco_e_id`: `monkeypatch.setattr(settings, "APP_PUBLIC_URL", "https://patrimonio.revalle.com.br/")` e item com `assigned_to="Fulano de Tal"`, `cpf="12345678909"` → `etiquetas[0]["url"] == f"https://patrimonio.revalle.com.br/e/{id}"`, `"Fulano" not in resp.text`, `"12345678909" not in resp.text` e `set(etiquetas[0]) == {"id", "tipo", "brand", "model", "identificador", "revenda", "url"}`.
  - `test_ordena_e_ignora_baixados_e_inexistentes`: `item_ids=[c, a, baixado, 999999]` → ids `[a, c]` e `ignorados == sorted([baixado, 999999])`.
  - `test_limites_da_selecao`: `[]` → 422 `"Equipamentos: Selecione ao menos um equipamento."`; 2001 ids → 422 `"Equipamentos: Selecione no máximo 2000 equipamentos por vez."`.
  - `test_ids_da_unidade_com_filtros`: com 2 Notebook e 1 Monitor na unidade e 1 item fora dela, `?revenda=<nome>` → `total == 3`; com `&tipo=Notebook` → 2; sem `revenda` → 422 `"Escolha a unidade."`.
  - `test_aprendiz_403`.
- [ ] **Step 2: Rodar e ver falhar:** `$RBT tests/test_etiquetas.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `$RBT tests/test_etiquetas.py -q` → PASS; depois a suíte inteira, `$RBT -q -m "not mudanca_esperada"` → verde.
- [ ] **Step 5: Commit** `feat(etiquetas): dados das etiquetas com QR apontando para APP_PUBLIC_URL/e/{id}`.

### Task 7: Dependências, clientes da API e rótulos (frontend)

**Files:**
- Modify: `frontend/package.json`, `frontend/package-lock.json`
- Create: `frontend/src/api/inventarios.ts`, `frontend/src/api/etiquetas.ts`, `frontend/src/api/inventarios.test.ts`, `frontend/src/features/inventario/resultados.ts`, `frontend/src/features/inventario/resultados.test.ts`

**Interfaces:**
- Consumes: os contratos das Tasks 3–6; `api` e `downloadAuthenticated` (`api/client.ts`); `StatusTone` (`lib/status.ts`).
- Produces:
  - Em `api/inventarios.ts`:
    - Tipos:
      - `type StatusInventario = 'aberto' | 'encerrado' | 'cancelado'`
      - `type ResultadoLeitura = 'ok' | 'local_errado' | 'usuario_diferente' | 'outra_unidade' | 'nao_cadastrado'`
      - `type CategoriaRelatorio = ResultadoLeitura | 'nao_encontrado'`
    - Interfaces `ResumoInventario`, `Inventario`, `InventarioDetalhe`, `Leitura` e `RelatorioLinha`, com os campos exatos das Tasks 3–5 (opcionais como `?: T | null`), e `RelatorioInventario { inventario: InventarioDetalhe; linhas: RelatorioLinha[] }`.
    - Funções:
      - `listInventarios(params: { unidade_id?: number; status?: StatusInventario; limit?: number; offset?: number }): Promise<Paginated<Inventario>>`
      - `createInventario(data: { unidade_id: number; observacao?: string }): Promise<InventarioDetalhe>`
      - `getInventario(id: number): Promise<InventarioDetalhe>`
      - `encerrarInventario(id: number): Promise<InventarioDetalhe>`
      - `cancelarInventario(id: number, motivo: string): Promise<InventarioDetalhe>`
      - `listLeituras(id: number, params: { resultado?: ResultadoLeitura; limit?: number; offset?: number }): Promise<Paginated<Leitura>>`
      - `registrarLeitura(id: number, data: { codigo: string; local_id_encontrado?: number | null }): Promise<Leitura>`
      - `atualizarLeitura(id: number, leituraId: number, data: { local_id_encontrado?: number | null; colaborador_id_encontrado?: number | null; usuario_conferido?: boolean; observacao?: string | null }): Promise<Leitura>`
      - `excluirLeitura(id: number, leituraId: number): Promise<void>`
      - `getRelatorio(id: number): Promise<RelatorioInventario>`
      - `exportarRelatorio(id: number, formato: 'xlsx' | 'pdf'): Promise<void>` = `downloadAuthenticated(`/inventarios/${id}/relatorio/exportar?formato=${formato}`, `inventario_${id}.${formato}`)`
  - Em `api/etiquetas.ts`:
    - `interface Etiqueta { id: number; tipo?: string | null; brand?: string | null; model?: string | null; identificador?: string | null; revenda?: string | null; url: string }`
    - `interface EtiquetasResposta { base_url: string; etiquetas: Etiqueta[]; ignorados: number[] }`
    - `listIdsParaEtiquetas(params: { revenda: string; tipo?: string; status?: string; search?: string }): Promise<{ ids: number[]; total: number }>`
    - `gerarEtiquetas(itemIds: number[]): Promise<EtiquetasResposta>`
  - Em `features/inventario/resultados.ts`:
    - `ROTULO_RESULTADO: Record<CategoriaRelatorio, string>` (singular, igual a `ROTULOS_CATEGORIA` do backend)
    - `ROTULO_CATEGORIA: Record<CategoriaRelatorio, string>` (plural: `Conferidos`, `Local errado`, `Usuário diferente`, `De outra unidade`, `Não cadastrados`, `Não encontrados`)
    - `TOM_RESULTADO: Record<CategoriaRelatorio, StatusTone>` (`ok` available, `local_errado` pending, `usuario_diferente` return, `outra_unidade` maintenance, `nao_cadastrado` danger, `nao_encontrado` danger)
    - `ORDEM_CATEGORIAS: CategoriaRelatorio[]` (mesma ordem do backend)
    - `ROTULO_STATUS_INVENTARIO = { aberto: 'Aberto', encerrado: 'Encerrado', cancelado: 'Cancelado' }` (estes rótulos já mapeiam para pending/retired em `statusTone`)
    - `rotuloCategoria(c: CategoriaRelatorio, status: StatusInventario): string`: `'Ainda não lidos'` para `nao_encontrado` com status `aberto`; senão `ROTULO_CATEGORIA[c]`.

- [ ] **Step 1: Instalar:** `npm install qrcode@1.5.4 @zxing/browser@0.2.1 @zxing/library@0.23.0` e `npm install -D @types/qrcode@1.5.5`. Conferir que `package.json` ganhou só essas 4 entradas.
- [ ] **Step 2: Escrever os testes:**
  - `resultados.test.ts`: `rotuloCategoria('nao_encontrado', 'aberto') === 'Ainda não lidos'`; `rotuloCategoria('nao_encontrado', 'encerrado') === 'Não encontrados'`; `TOM_RESULTADO.ok === 'available'`; `ORDEM_CATEGORIAS[0] === 'nao_encontrado'`.
  - `inventarios.test.ts` (MSW):
    - `registrarLeitura(12, { codigo: '#5', local_id_encontrado: 3 })` faz `POST /api/inventarios/12/leituras` com esse corpo.
    - `exportarRelatorio(12, 'pdf')` faz `GET /api/inventarios/12/relatorio/exportar?formato=pdf`, com `URL.createObjectURL`/`revokeObjectURL` mockados por `vi.fn`.
- [ ] **Step 3: Rodar e ver falhar:** `npx vitest run src/api/inventarios.test.ts src/features/inventario` → FAIL.
- [ ] **Step 4: Implementar.**
- [ ] **Step 5: Rodar** os mesmos testes e `npx tsc --noEmit` → verdes.
- [ ] **Step 6: Commit** `feat(inventario): clientes da API de inventário e etiquetas e dependências de QR`.

### Task 8: Layout da folha, QR em SVG e camada de impressão

**Files:**
- Create em `frontend/src/features/etiquetas/`: `layout.ts`, `layout.test.ts`, `qr.ts`, `qr.test.ts`, `FolhasEtiquetas.tsx`, `FolhasEtiquetas.test.tsx`, `ImpressaoEtiquetas.tsx`, `ImpressaoEtiquetas.test.tsx`

**Interfaces:**
- Consumes: `Etiqueta` (Task 7).
- Produces:
  - Em `layout.ts`:
    - `interface LayoutEtiquetas { colunas: number; linhas: number; larguraMm: number; alturaMm: number; margemSuperiorMm: number; margemEsquerdaMm: number; espacoHorizontalMm: number; espacoVerticalMm: number }`
    - `FOLHA_A4 = { larguraMm: 210, alturaMm: 297 }`
    - `LAYOUT_PADRAO = { colunas: 3, linhas: 8, larguraMm: 70, alturaMm: 37, margemSuperiorMm: 0.5, margemEsquerdaMm: 0, espacoHorizontalMm: 0, espacoVerticalMm: 0 }`
    - `CHAVE_LAYOUT = 'revalle_etiquetas_layout'`
  - `validarLayout(l: LayoutEtiquetas): string | null`, que devolve a primeira falha nesta ordem:
    - `colunas` inteiro de 1 a 10, senão `"Colunas: informe um número inteiro de 1 a 10."`
    - `linhas` inteiro de 1 a 30, senão `"Linhas: informe um número inteiro de 1 a 30."`
    - `larguraMm` de 10 a 210, senão `"Largura: informe de 10 a 210 mm."`
    - `alturaMm` de 10 a 297, senão `"Altura: informe de 10 a 297 mm."`
    - margens e espaços de 0 a 50, senão `"Margens e espaçamentos: informe de 0 a 50 mm."`
    - largura total `margemEsquerdaMm + colunas*larguraMm + (colunas-1)*espacoHorizontalMm` até `210 + 0.01`, e a altura equivalente até `297 + 0.01`, senão `"O layout não cabe numa folha A4 (210 × 297 mm)."`
  - `etiquetasPorFolha(l): number` (= `colunas * linhas`).
  - `validarPosicaoInicial(l, p: number): string | null`: devolve `` `A posição inicial deve estar entre 1 e ${n}.` `` se `p` não for inteiro de 1 a n.
  - `interface PosicaoEtiqueta { folha: number; linha: number; coluna: number; esquerdaMm: number; topoMm: number }`.
  - `calcularPosicoes(l, quantidade: number, posicaoInicial: number): PosicaoEtiqueta[]`:
    - preenche linha a linha, da esquerda para a direita;
    - `slot = posicaoInicial - 1 + k`, `folha = floor(slot / porFolha)`, `idx = slot % porFolha`;
    - `esquerdaMm = margemEsquerdaMm + coluna*(larguraMm + espacoHorizontalMm)` e `topoMm` análogo.
  - `totalFolhas(l, quantidade, posicaoInicial): number`.
  - `carregarLayout(): LayoutEtiquetas`: `try/catch`; JSON inválido, campo faltando ou `validarLayout` falhando devolve `LAYOUT_PADRAO`.
  - `salvarLayout(l): void`: `try/catch`, silencioso.
  - `qr.ts`:
    - `gerarQrSvg(conteudo: string): Promise<string>`: `const mod = await import('qrcode'); const QRCode = (mod as { default?: typeof mod }).default ?? mod`, depois `QRCode.toString(conteudo, { type: 'svg', errorCorrectionLevel: 'M', margin: 1 })`. As cores do QR ficam no padrão preto/branco da biblioteca. É o mesmo `import()` dinâmico que a F2 usa para o QR do 2FA; não pode haver import estático de `qrcode` em lugar nenhum.
    - `gerarQrs(etiquetas: Etiqueta[]): Promise<Record<number, string>>`.
  - `FolhasEtiquetas({ etiquetas, qrs, layout, posicaoInicial }: { etiquetas: Etiqueta[]; qrs: Record<number, string>; layout: LayoutEtiquetas; posicaoInicial: number })`:
    - Uma `div.folha` (`data-testid="folha"`, `width: 210mm; height: 297mm; position: relative`) por folha.
    - Cada etiqueta (`data-testid="etiqueta"`) tem `position: absolute` e `left/top/width/height` em `mm` vindos de `calcularPosicoes`. Por dentro:
      - à esquerda, o QR (`dangerouslySetInnerHTML={{ __html: qrs[id] }}`, o SVG gerado pela própria biblioteca), num quadrado de `Math.min(alturaMm - 4, larguraMm * 0.45)` mm com `aria-label` `QR do patrimônio #<id>`;
      - à direita, `#<id>` (12pt, negrito, monoespaçado), `tipo` (8pt) e `brand model` (7pt, no máximo 2 linhas).
  - `ImpressaoEtiquetas({ children }: { children: ReactNode })`: no mount, cria `div.etiquetas-impressao`, anexa a `document.body` e faz `document.body.classList.add('imprimindo-etiquetas')`. Renderiza por `createPortal` o `<style>{CSS_IMPRESSAO}</style>` seguido dos `children`. No unmount, remove a div e a classe. `export const CSS_IMPRESSAO`:

```css
@page { size: A4 portrait; margin: 0; }
@media screen { .etiquetas-impressao { display: none; } }
@media print {
  body.imprimindo-etiquetas { margin: 0; background: white; }
  body.imprimindo-etiquetas > *:not(.etiquetas-impressao) { display: none !important; }
  .etiquetas-impressao { display: block; }
  .etiquetas-impressao, .etiquetas-impressao * { color: black !important; background: white !important; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
  .etiquetas-impressao .folha { break-after: page; page-break-after: always; }
  .etiquetas-impressao .folha:last-child { break-after: auto; page-break-after: auto; }
}
```

- [ ] **Step 1: Escrever os testes:**
  - `layout.test.ts`:
    - `validarLayout(LAYOUT_PADRAO) === null`; `etiquetasPorFolha(LAYOUT_PADRAO) === 24`.
    - `{...LAYOUT_PADRAO, colunas: 4}` → `"O layout não cabe numa folha A4 (210 × 297 mm)."`; `colunas: 0` → `"Colunas: informe um número inteiro de 1 a 10."`.
    - `validarPosicaoInicial(LAYOUT_PADRAO, 25) === "A posição inicial deve estar entre 1 e 24."`.
    - `calcularPosicoes(LAYOUT_PADRAO, 2, 5)` → `[{ folha: 0, linha: 1, coluna: 1, esquerdaMm: 70, topoMm: 37.5 }, { folha: 0, linha: 1, coluna: 2, esquerdaMm: 140, topoMm: 37.5 }]` (`toBeCloseTo` nos mm).
    - Com 25 etiquetas a partir da posição 1, a última fica em `{ folha: 1, linha: 0, coluna: 0 }`.
    - `totalFolhas(LAYOUT_PADRAO, 20, 10) === 2`.
    - `carregarLayout()`: com `'{x'` no storage → padrão; com JSON válido de `colunas: 2` → `colunas === 2`; com `getItem` lançando (`vi.spyOn(Storage.prototype, 'getItem')`) → padrão; `salvarLayout` não lança com `setItem` lançando.
  - `qr.test.ts`: `gerarQrSvg('https://patrimonio.revalle.com.br/e/12')` começa com `'<svg'` e contém `'viewBox'`; `Object.keys(await gerarQrs([...2 etiquetas]))` → `['1', '2']`.
  - `FolhasEtiquetas.test.tsx`:
    - 26 etiquetas no padrão → 2 `folha`s.
    - Com `posicaoInicial=5`, a primeira etiqueta tem `style.left === '70mm'` e `style.top === '37.5mm'`.
    - Aparecem `'#12'`, `'Notebook'` e `'Dell Latitude 5420'`.
  - `ImpressaoEtiquetas.test.tsx`: montado → `document.body.classList.contains('imprimindo-etiquetas')` e `document.querySelector('.etiquetas-impressao style')?.textContent` contém `'@page'`; desmontado → nem a classe nem a div.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/etiquetas` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run src/features/etiquetas` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(etiquetas): layout da folha A4 em mm, QR em SVG e camada de impressão`.

### Task 9: Tela `/etiquetas`

**Files:**
- Create: `frontend/src/pages/EtiquetasPage.tsx`, `frontend/src/pages/EtiquetasPage.test.tsx`, `frontend/src/features/etiquetas/LayoutForm.tsx`
- Modify:
  - `frontend/src/App.tsx`: import + rota `etiquetas` com `RequireRole roles={['Gestor', 'Técnico']}`.
  - `frontend/src/components/layout/navigation.ts`: nova linha `import { QrCode } from 'lucide-react'` (um import separado, sem editar o existente). No grupo `Inventário`, depois de `/locais`, a linha `{ to: '/etiquetas', label: 'Etiquetas', icon: QrCode, roles: ['Gestor', 'Técnico'] }`.

**Interfaces:**
- Consumes: `listItemsPaginated` (`api/items.ts`, só leitura), `listUnidades`, `useConstants().equipment_types`, `listIdsParaEtiquetas`, `gerarEtiquetas` (Task 7); tudo da Task 8; `Pagination`, `usePageSize('etiquetas')`, `StatusPill`, `Checkbox`, `Button`, `upper`, `useDebouncedValue`, `toast`, `getErrorMessage`.
- Produces: `EtiquetasPage` (default export) e `LayoutForm({ layout, posicaoInicial, onLayoutChange, onPosicaoChange, onRestaurar })`. O `LayoutForm` tem campos numéricos (`step="0.5"` nos mm) com os rótulos `Colunas`, `Linhas`, `Largura (mm)`, `Altura (mm)`, `Margem superior (mm)`, `Margem esquerda (mm)`, `Espaço entre colunas (mm)`, `Espaço entre linhas (mm)`, `Posição inicial`, além do botão `Restaurar padrão`.
- Comportamento da tela:
  - **Filtros:** `Unidade` (`aria-label="Unidade"`; o valor é `unidade.nome`, exibido com `upper`), `Tipo`, `Status` (opções fixas `['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução', 'Em Manutenção']`) e `Buscar` (debounce de 300 ms). Mudar um filtro volta à página 1. A consulta é `listItemsPaginated({ revenda, tipo, status, search, limit: pageSize, offset })`.
  - **Tabela:**
    - Colunas: checkbox, Patrimônio (`#id`), Tipo, Equipamento, Nº de série, Unidade (`upper`), Status (`StatusPill`).
    - Cabeçalho com checkbox `aria-label="Selecionar a página"`; cada linha com `aria-label={`Selecionar #${id}`}`.
    - A seleção é um `Set<number>` e sobrevive à troca de filtro e de página.
  - **Barra de seleção:**
    - Texto `` `${n} selecionado(s)` `` e o botão `Limpar seleção`.
    - `Selecionar todos da unidade`: desabilitado sem unidade. Chama `listIdsParaEtiquetas` com os filtros atuais, soma os ids à seleção e mostra o toast `` `${total} equipamento(s) selecionado(s).` ``.
    - `Gerar etiquetas` (`loading`): desabilitado com 0; com mais de 2000 mostra `"Selecione no máximo 2000 equipamentos por vez."`. Chama `gerarEtiquetas([...sel].sort((a, b) => a - b))` e depois `gerarQrs`. Com `ignorados.length > 0`, mostra o toast `` `${n} equipamento(s) ignorado(s): baixados ou inexistentes.` ``.
  - **Seção "Folha de etiquetas"** (depois de gerar):
    - O `LayoutForm`. A cada alteração válida, chama `salvarLayout`. A posição inicial começa em 1 e não é salva.
    - Resumo `` `${n} etiqueta(s) em ${f} folha(s)` ``.
    - Erros de `validarLayout`/`validarPosicaoInicial` aparecem em `role="alert"` e desabilitam `Imprimir`.
    - Aviso quando `new URL(base_url).origin !== window.location.origin`: `` `Atenção: os QR codes vão apontar para ${base_url}, diferente do endereço desta página (${window.location.origin}). Confira APP_PUBLIC_URL no servidor antes de imprimir.` ``.
    - Dica: `Na janela de impressão, escolha tamanho real (100%) e margens "Nenhuma".`
    - Prévia: `FolhasEtiquetas` dentro de um contêiner com `overflow-auto` e `transform: scale(0.5)` (`origin-top-left`).
    - `<ImpressaoEtiquetas><FolhasEtiquetas .../></ImpressaoEtiquetas>`.
    - `Imprimir` → `window.print()`.

- [ ] **Step 1: Escrever os testes** (`EtiquetasPage.test.tsx`; MSW e `useAuth` mockado como em `LocaisPage.test.tsx`; render via `App` em `MemoryRouter` com `['/etiquetas']`; `window.print = vi.fn()`; `localStorage.clear()` no `beforeEach`):
  - `filtra pela unidade e pagina no servidor`: `GET /api/items` recebe `revenda=Juazeiro&limit=25&offset=0`.
  - `gera as etiquetas dos selecionados em ordem crescente`: marca `#5` e depois `#3` → `POST /api/etiquetas` com `{ item_ids: [3, 5] }`; a prévia mostra `'#3'`.
  - `selecionar todos da unidade usa os filtros e soma à seleção`: o botão começa desabilitado; com a unidade escolhida → `GET /api/etiquetas/ids?revenda=Juazeiro` (resposta com 40 ids) → aparece `'40 selecionado(s)'`.
  - `layout inválido bloqueia a impressão`: `Colunas` = 4 → `role="alert"` com `"O layout não cabe numa folha A4 (210 × 297 mm)."` e `Imprimir` desabilitado; voltando a 3, clicar `Imprimir` chama `window.print`.
  - `guarda o layout e não a posição inicial`: `Colunas` 2 e `Posição inicial` 5. `JSON.parse(localStorage.getItem('revalle_etiquetas_layout')!).colunas === 2` e o JSON não tem `posicaoInicial`. Remontado e gerado de novo → `Colunas` 2 e `Posição inicial` 1.
  - `avisa quando o endereço do QR difere`: com `base_url: 'https://outro.exemplo'`, aparece o texto `'Confira APP_PUBLIC_URL no servidor'`.
  - `sair da tela limpa a impressão`: `unmount()` → `document.body.classList.contains('imprimindo-etiquetas') === false`.
  - `Jovem Aprendiz não acessa`: aparece `'Acesso restrito'`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/EtiquetasPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(etiquetas): tela de seleção e impressão de etiquetas QR em folha A4`.

### Task 10: Rota `/e/:id` e retorno ao item depois do login

**Files:**
- Create: `frontend/src/features/etiquetas/posLogin.ts`, `frontend/src/features/etiquetas/posLogin.test.ts`, `frontend/src/features/etiquetas/PosLoginRedirect.tsx`, `frontend/src/pages/EtiquetaRedirectPage.tsx`, `frontend/src/pages/EtiquetaRedirectPage.test.tsx`
- Modify: `frontend/src/App.tsx`:
  - import;
  - `<Route path="/e/:id" element={<EtiquetaRedirectPage />} />` logo depois da rota `/login`, **fora** do `AppLayout`;
  - a linha `<PosLoginRedirect />` antes de `<ToastContainer />`.

**Interfaces:**
- Consumes: `useAuth()` (`user`, `isLoading`), `react-router-dom`.
- Produces:
  - `posLogin.ts`:
    - `CHAVE_DESTINO = 'revalle_destino_pos_login'`.
    - `guardarDestinoPosLogin(destino: string): void`: grava `{ destino, em: Date.now() }` só se `destino` casar `/^\/items\/\d+$/`, com `try/catch`.
    - `consumirDestinoPosLogin(agora = Date.now()): string | null`: lê e **apaga** o valor; devolve `null` se o formato for inválido ou se passou de 10 min (`10 * 60 * 1000`); usa `try/catch`.
  - `PosLoginRedirect(): null`: com `user`, `location.pathname === '/'` e destino válido, chama `navigate(destino, { replace: true })`. Sem usuário, não consome nada.
  - `EtiquetaRedirectPage`:
    - `id` que não seja só dígitos → `<Navigate to="/" replace />`.
    - `isLoading` → `LoadingState` "Abrindo equipamento...".
    - Com usuário → `<Navigate to={`/items/${id}`} replace />`.
    - Sem usuário → `guardarDestinoPosLogin(`/items/${id}`)` (chamada idempotente) e `<Navigate to="/login" replace />`.

- [ ] **Step 1: Escrever os testes:**
  - `posLogin.test.ts`:
    - guarda e consome uma única vez (o segundo `consumir` dá `null`);
    - `consumirDestinoPosLogin(Date.now() + 10 * 60 * 1000 + 1) === null`;
    - `guardarDestinoPosLogin('//evil.com')` e `('/users')` não gravam;
    - `sessionStorage` lançando não quebra.
  - `EtiquetaRedirectPage.test.tsx`: `MemoryRouter` com `Routes` de teste (`/e/:id` → página; `/login` → `<p>tela de login</p>`; `/items/:id` → `<p>detalhe</p>`; `/` → `<p>início</p>`) e `useAuth` mockado.
    - logado → mostra `'detalhe'`;
    - deslogado → `'tela de login'` e `consumirDestinoPosLogin() === '/items/42'`;
    - `/e/abc` → `'início'`.
  - `PosLoginRedirect`: logado em `/` com destino guardado → `'detalhe'`; logado em `/stock` → não navega; deslogado → o destino continua guardado.
  - Integração: `App` em `MemoryRouter(['/e/42'])` com usuário logado → chama `GET /api/items/42` (MSW).
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/etiquetas/posLogin.test.ts src/pages/EtiquetaRedirectPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(etiquetas): atalho /e/{id} para o detalhe do equipamento com retorno após o login`.

### Task 11: Lista de inventários e abertura

**Files:**
- Create: `frontend/src/pages/InventariosPage.tsx`, `frontend/src/pages/InventariosPage.test.tsx`, `frontend/src/components/inventario/NovoInventarioDialog.tsx`
- Modify:
  - `frontend/src/App.tsx`: rota `inventarios`, Gestor e Técnico.
  - `frontend/src/components/layout/navigation.ts`: nova linha `import { ClipboardCheck } from 'lucide-react'` e, no grupo `Inventário`, a linha `{ to: '/inventarios', label: 'Inventários', icon: ClipboardCheck, roles: ['Gestor', 'Técnico'] }`.

**Interfaces:**
- Consumes: `listInventarios`, `createInventario`, `ROTULO_STATUS_INVENTARIO` (Task 7); `listUnidades`; `Pagination`, `usePageSize('inventarios')`, `StatusPill`, `Button loading`, `Dialog`, `upper`, `formatDateTime`, `toast`, `getErrorMessage`.
- Produces:
  - `InventariosPage`:
    - Filtros: `Unidade` (por id) e `Status`. A troca de filtro volta à página 1.
    - Tabela: Nº (`#id`), Unidade (`upper`), Status (`StatusPill status={ROTULO_STATUS_INVENTARIO[s]}`), Aberto por, Aberto em, Encerrado em, Leituras.
    - Ações por linha: link `Ler` para `/inventarios/:id/leitura` (só quando aberto) e link `Relatório` para `/inventarios/:id`.
    - Botão `Novo inventário`.
  - `NovoInventarioDialog({ open, onOpenChange })`:
    - select `aria-label="Unidade do inventário"` (só unidades ativas) e textarea `Observação` opcional;
    - `Abrir inventário` (`loading`) → `navigate(`/inventarios/${id}/leitura`)`;
    - em erro, toast com `getErrorMessage`.

- [ ] **Step 1: Escrever os testes** (MSW):
  - `lista com unidade em caixa alta e status`: `'JUAZEIRO'` e `'Aberto'`; a API é chamada com `limit=25&offset=0`.
  - `filtro de status volta à página 1`: `status=aberto&offset=0`.
  - `abrir inventário vai para a leitura`: render isolado em `MemoryRouter` com a rota de teste `/inventarios/:id/leitura` → `<p>tela de leitura</p>`; `POST` com `{ unidade_id: 2 }` → aparece `'tela de leitura'`.
  - `inventário já aberto mostra a mensagem da API`: 409 com `detail` `"Já existe um inventário aberto para esta unidade (#7)."` → o texto aparece.
  - `Jovem Aprendiz não acessa`: via `App` → `'Acesso restrito'`.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(inventario): lista de inventários e abertura por unidade`.

### Task 12: Leitor de QR pela câmera

**Files:**
- Create: `frontend/src/features/inventario/scanner.ts`, `frontend/src/features/inventario/scanner.test.ts`, `frontend/src/components/inventario/LeitorCamera.tsx`, `frontend/src/components/inventario/LeitorCamera.test.tsx`

**Interfaces:**
- Produces:
  - Em `scanner.ts`:
    - `type TipoErroCamera = 'inseguro' | 'negada' | 'sem_camera' | 'falha'`
    - `class ErroCamera extends Error { constructor(readonly tipo: TipoErroCamera) }` (a `message` é `MENSAGEM_ERRO_CAMERA[tipo]`)
    - `MENSAGEM_ERRO_CAMERA`:
      - `inseguro`: `'A câmera só funciona em conexão segura (HTTPS). Use a digitação ou um leitor USB.'`
      - `negada`: `'Permissão da câmera negada. Use a digitação abaixo ou libere a câmera nas configurações do navegador.'`
      - `sem_camera`: `'Nenhuma câmera disponível neste aparelho. Use a digitação ou um leitor USB.'`
      - `falha`: `'Não foi possível abrir a câmera. Use a digitação ou um leitor USB.'`
    - `interface ControlesLeitor { parar(): void }`
  - `iniciarLeitor(video: HTMLVideoElement, onCodigo: (texto: string) => void): Promise<ControlesLeitor>`, nesta ordem:
    1. `!window.isSecureContext` → `ErroCamera('inseguro')`.
    2. Sem `navigator.mediaDevices?.getUserMedia` → `ErroCamera('sem_camera')`.
    3. `const { BrowserQRCodeReader } = await import('@zxing/browser')` e `new BrowserQRCodeReader(undefined, { delayBetweenScanAttempts: 150 })`.
    4. `await reader.decodeFromConstraints({ audio: false, video: { facingMode: { ideal: 'environment' } } }, video, (result) => { if (result) onCodigo(result.getText()) })`.
    5. Erros pelo `name`: `NotAllowedError`/`SecurityError` → `negada`; `NotFoundError`/`OverconstrainedError` → `sem_camera`; o resto → `falha`.
    6. `parar()` é idempotente: chama `controls.stop()` e ainda para cada trilha de `video.srcObject` e zera `srcObject`.
  - `LeitorCamera({ onCodigo, onErro, onPausar }: { onCodigo: (texto: string) => void; onErro: (e: ErroCamera) => void; onPausar: () => void })`:
    - Renderiza `<video muted playsInline autoPlay aria-label="Imagem da câmera" className="aspect-square w-full max-w-sm rounded-md bg-surface-alt object-cover" />`.
    - Inicia no mount.
    - Se for desmontado antes de `iniciarLeitor` resolver, chama `parar()` assim que resolver.
    - No unmount, chama `parar()`.
    - `visibilitychange` com `document.visibilityState === 'hidden'` → `parar()` + `onPausar()`.
    - Ignora o mesmo texto repetido em menos de 3000 ms (`Date.now()`).
    - Converte erros que não são `ErroCamera` em `ErroCamera('falha')`.

- [ ] **Step 1: Escrever os testes:**
  - `scanner.test.ts`: `vi.mock('@zxing/browser')` com uma classe falsa cujo `decodeFromConstraints` é `vi.fn()`; `Object.defineProperty` para `window.isSecureContext` e `navigator.mediaDevices`.
    - inseguro → rejeita com `tipo === 'inseguro'` sem instanciar o leitor;
    - `DOMException('x', 'NotAllowedError')` → `'negada'`;
    - o callback com `{ getText: () => 'https://x/e/1' }` chama `onCodigo('https://x/e/1')`;
    - `parar()` chama `stop` e o `stop()` de cada trilha (`video.srcObject` falso com `getTracks()`); chamar `parar()` duas vezes não quebra.
  - `LeitorCamera.test.tsx`: `vi.mock('@/features/inventario/scanner', async (orig) => ({ ...(await orig()), iniciarLeitor: vi.fn() }))`.
    - desmontado antes de resolver → depois de resolver, `parar` chamado 1×;
    - aba escondida → `parar` e `onPausar` chamados;
    - o mesmo código duas vezes em menos de 3 s → `onCodigo` 1×; com `Date.now` avançado 3001 ms → 2×;
    - rejeição `ErroCamera('negada')` → `onErro` com `tipo 'negada'`.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/inventario/scanner.test.ts src/components/inventario/LeitorCamera.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** → PASS; `npx tsc --noEmit` → verde.
- [ ] **Step 5: Commit** `feat(inventario): leitor de QR pela câmera com @zxing/browser sob demanda`.

### Task 13: Tela de leitura (celular) e política de câmera no nginx

**Files:**
- Create: `frontend/src/pages/InventarioLeituraPage.tsx`, `frontend/src/pages/InventarioLeituraPage.test.tsx`, `frontend/src/components/inventario/UltimaLeituraCard.tsx`, `frontend/src/test/nginxPolicy.test.ts`
- Modify:
  - `frontend/src/App.tsx`: rota `inventarios/:id/leitura`, Gestor e Técnico.
  - `frontend/nginx.conf.template`: acrescentar, depois de `client_max_body_size`, as linhas abaixo.

```nginx
    # Câmera só para o próprio site (leitura de QR do inventário, F6). O navegador
    # só libera a câmera em HTTPS (ou localhost); em produção o TLS termina no proxy
    # do EasyPanel. Vale para o index.html (location /); o bloco de assets tem
    # add_header próprio e não herda este, o que não importa para a câmera.
    add_header Permissions-Policy "camera=(self), microphone=(), geolocation=()" always;
```

**Interfaces:**
- Consumes: `getInventario`, `registrarLeitura`, `atualizarLeitura`, `excluirLeitura`, `listLeituras`, `ROTULO_RESULTADO`, `TOM_RESULTADO` (Task 7); `LeitorCamera`, `MENSAGEM_ERRO_CAMERA`, `ErroCamera` (Task 12); `LocalSelect({ unidadeId, value, onChange })`, `ColaboradorSelect({ value, onChange })`, `StatusPill`, `Button loading`, `AlertDialog`, `EmptyState`, `upper`, `formatDateTime`.
- Produces:
  - `InventarioLeituraPage` (`/inventarios/:id/leitura`), em uma coluna `max-w-xl mx-auto`, botões `size="lg"`:
    - Inventário que não está `aberto` → `EmptyState` com o título `"Este inventário não está aberto."` e o link `Ver relatório` para `/inventarios/:id`, sem campo de leitura.
    - Cabeçalho `` `Inventário #${id} · ${upper(unidade_nome)}` ``, o texto `` `Encontrados ${encontrados} de ${esperados} · Faltam ${nao_encontrados}` `` e o link `Relatório`.
    - `Estou no local`: `LocalSelect` da unidade do inventário. O valor vai como `local_id_encontrado` em cada leitura.
    - Câmera:
      - Botão `Ler com a câmera`, desabilitado com a mensagem `inseguro` visível quando `!window.isSecureContext`. Quando ativo, renderiza `LeitorCamera` + o botão `Parar câmera`.
      - `onErro` → desliga a câmera, mostra `MENSAGEM_ERRO_CAMERA[tipo]` em `role="alert"` e foca o campo de digitação.
      - `onPausar` → desliga e mostra `Câmera pausada. Toque em "Ler com a câmera" para continuar.`.
      - O que a câmera lê vai para a mesma mutação, e é ignorado enquanto houver uma pendente.
    - Digitação: `form` com o rótulo `Código ou número do patrimônio`, `autoFocus`, `autoComplete="off"`, `enterKeyHint="send"` e o botão `Registrar` (`loading`). O `Enter` envia, o que já atende o leitor USB. Depois de enviar, limpa o campo e devolve o foco.
    - Depois de cada mutação, invalida `['inventario', id]` e `['inventario-leituras', id]`.
    - `Últimas leituras`: `listLeituras(id, { limit: 10 })`.
  - `UltimaLeituraCard({ leitura, onInformarUsuario, onSemUsuario, onDesfazer })`:
    - Mostra:
      - `StatusPill status={ROTULO_RESULTADO[r]} tone={TOM_RESULTADO[r]}`;
      - `` `#${item_id} · ${item_tipo} · ${item_brand} ${item_model}` `` (sem item, mostra `codigo_lido`);
      - o `detalhe`;
      - `` `Usuário cadastrado: ${usuario_esperado ?? 'nenhum'}` ``;
      - quando `duplicada`, `` `Já lido em ${formatDateTime(lido_em)} por ${lido_por} — não contou de novo.` ``.
    - Ações (as duas primeiras só quando há `item_id`):
      - `Informar usuário`: abre o `ColaboradorSelect`; a escolha manda `PATCH { colaborador_id_encontrado: c.id }`.
      - `Ninguém usando`: manda `PATCH { usuario_conferido: true, colaborador_id_encontrado: null }`.
      - `Desfazer leitura`: abre o `AlertDialog` "Desfazer esta leitura?", com o botão `Desfazer` → `DELETE`.

- [ ] **Step 1: Escrever os testes:**
  - `nginxPolicy.test.ts`: lê `resolve(__dirname, '../../nginx.conf.template')`. O texto contém `Permissions-Policy` e `camera=(self)`, e não casa `/camera=\(\)/`.
  - `InventarioLeituraPage.test.tsx`: MSW; `scanner` mockado como na Task 12; `Object.defineProperty(window, 'isSecureContext', { value: true, configurable: true })` no `beforeEach`; render via `App` em `['/inventarios/12/leitura']`.
    - `digitação registra com o local atual`: escolher `Local` 3 e digitar `#12{Enter}` → `POST` com `{ codigo: '#12', local_id_encontrado: 3 }`. Aparecem `'Conferido'` e `'#12 · Notebook'`; o campo fica vazio e com foco.
    - `leitura duplicada avisa que não contou de novo`.
    - `câmera negada mostra a mensagem e foca a digitação`: `iniciarLeitor` rejeita `new ErroCamera('negada')` → aparece `'Permissão da câmera negada'` e o campo tem foco.
    - `código lido pela câmera é registrado`: chamar o `onCodigo` capturado com `'https://p.revalle.com.br/e/12'` → `POST` com esse `codigo`.
    - `"Ninguém usando" corrige a leitura`: `PATCH` com `{ usuario_conferido: true, colaborador_id_encontrado: null }`; aparece `'Usuário diferente'`.
    - `desfazer leitura chama DELETE`.
    - `inventário encerrado não oferece leitura`: aparece `'Este inventário não está aberto.'`; não há o campo `Código ou número do patrimônio`.
    - `sem HTTPS a câmera fica desabilitada e a digitação continua`.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(inventario): tela de leitura no celular com câmera, digitação e leitor USB`.

### Task 14: Relatório do inventário (tela)

**Files:**
- Create: `frontend/src/pages/InventarioRelatorioPage.tsx`, `frontend/src/pages/InventarioRelatorioPage.test.tsx`
- Modify: `frontend/src/App.tsx` (rota `inventarios/:id`, Gestor e Técnico)

**Interfaces:**
- Consumes: `getRelatorio`, `encerrarInventario`, `cancelarInventario`, `exportarRelatorio`, `ORDEM_CATEGORIAS`, `rotuloCategoria`, `ROTULO_STATUS_INVENTARIO` (Task 7); `DetailTabs({ tabs, activeId, onSelect, label, children })`, `StatBlock({ label, value })`, `Pagination` com fatia no cliente, `usePageSize('inventario-relatorio')`, `StatusPill`, `AlertDialog`, `Dialog`, `Button loading`, `useAuth`.
- Produces: `InventarioRelatorioPage` (`/inventarios/:id`):
  - **Cabeçalho:** `` `Inventário #${id}` ``, unidade com `upper`, `StatusPill`, aberto por/em, encerrado por/em, observação e motivo do cancelamento.
  - **Indicadores (`StatBlock`):** Esperados e um por categoria (`rotuloCategoria`).
  - **Abas:**
    - Uma `DetailTabs` (`label="Resultados do inventário"`) por categoria, na ordem de `ORDEM_CATEGORIAS`, com o rótulo `` `${rotuloCategoria(c, status)} (${n})` ``.
    - O estado fica em `?aba=`. Sem `?aba=`, abre a primeira categoria com `n > 0`; se todas estiverem vazias, `nao_encontrado`.
  - **Tabela:** dentro de `overflow-x-auto`.
    - Colunas: Patrimônio (link para `/items/:id` quando há item), Tipo, Equipamento, Nº de série, Unidade do equipamento (`upper`), Local cadastrado, Local encontrado, Usuário cadastrado, Usuário encontrado, Lido por / em, Detalhe / observação.
    - Página no cliente com `Pagination`; a troca de aba volta à página 1.
  - **Ações:**
    - Aberto: link `Continuar leitura` para `/inventarios/:id/leitura`.
    - Aberto: `Encerrar inventário`, com `AlertDialog` cujo texto é `` `${resumo.nao_encontrados} equipamento(s) ainda não lido(s) serão marcados como não encontrados. Depois de encerrado, o inventário não aceita novas leituras.` ``. Confirmar → `encerrarInventario` → toast `Inventário encerrado.` e recarga.
    - Gestor e aberto: `Cancelar inventário`, com um `Dialog` de motivo obrigatório → `cancelarInventario`.
    - Sempre: `Exportar Excel` e `Exportar PDF` (`loading`).

- [ ] **Step 1: Escrever os testes** (MSW, `App` em `['/inventarios/12']`):
  - `mostra o resumo e as abas com contagem`: inventário encerrado → `'Não encontrados (2)'` e `'Conferidos (1)'`; clicar em `Não encontrados` mostra `'#7'`.
  - `inventário aberto chama de "Ainda não lidos" e oferece continuar`: aparece `'Ainda não lidos (2)'` e o link `Continuar leitura` com `href` `/inventarios/12/leitura`.
  - `encerrar pede confirmação e chama a API`: o diálogo contém `'2 equipamento(s) ainda não lido(s)'` → `POST /api/inventarios/12/encerrar`.
  - `cancelar só para o Gestor`: o Técnico não vê o botão; o Gestor envia `{ motivo: 'Contagem refeita' }`.
  - `exportar Excel baixa o arquivo`: `GET .../relatorio/exportar?formato=xlsx`, com `URL.createObjectURL` mockado.
- [ ] **Step 2: Rodar e ver falhar.**
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(inventario): relatório por inventário com encerramento e exportação`.

### Task 15: Verificação final da fase

**Files:** nenhum novo. Só correções que as verificações apontarem, cada uma com teste.

- [ ] **Step 1: Backend inteiro:** `$RBT -q -m "not mudanca_esperada"` → todos passam. Anotar o total: o de antes da F6 mais os novos.
- [ ] **Step 2: Frontend inteiro:** em `frontend/`, `npx vitest run`, `npx tsc --noEmit` e `npm run build` → verdes.
- [ ] **Step 3: Carga sob demanda:** depois do build, `grep -l "got interrupted by new loading request" dist/assets/*.js` (zxing) e `grep -l "Unknown EC Level" dist/assets/*.js` (qrcode) só podem listar chunks que **não** sejam `index-*.js`. Anotar no relatório o tamanho dos dois chunks que o `npm run build` imprime.
- [ ] **Step 4: Sintaxe do nginx:** `docker run --rm -e BACKEND_HOST=backend -e BACKEND_PORT=8000 -v "<wt>/frontend/nginx.conf.template:/etc/nginx/templates/default.conf.template:ro" nginx:alpine nginx -t` → `syntax is ok`. O container é descartado (`--rm`), e não se cria imagem nova.
- [ ] **Step 5: Conferência manual (registrar no relatório o que foi feito e o que não foi possível):**
  - **Impressão:** imprimir em PDF no Chrome uma folha com 26 etiquetas, layout padrão, posição inicial 5. A grade de 3 × 8 tem 70 × 37 mm, a primeira etiqueta fica na 2ª coluna da 2ª linha, e a 2ª página traz as restantes. O QR impresso, lido pela câmera de um celular, abre `APP_PUBLIC_URL/e/<id>`.
  - **Celular (390 px no DevTools):** `/inventarios/:id/leitura` e `/etiquetas` sem rolagem horizontal da página.
  - **Câmera** em `http://localhost:5173`, que é contexto seguro (webcam): lê o QR, para ao trocar de aba e ao sair da tela (o indicador de câmera do navegador apaga).
  - **Produção:** a câmera exige HTTPS. Antes de usar no celular, rodar `curl -sI https://<dominio>/ | grep -i permissions-policy` e confirmar `camera=(self)`, sem nenhum `camera=()` vindo do proxy do EasyPanel.
- [ ] **Step 6: Checklist contra o spec**, item a item, com a task e o teste de cada um no relatório:
  - **11.1** Etiquetas:
    - seleção na listagem ou todos da unidade (T6, T9);
    - folha A4 com `@media print` (T8, T9);
    - QR + número + tipo/modelo (T8);
    - padrão 3×8, 70×37 mm, 24 por folha (T8);
    - colunas, linhas, margens e posição inicial configuráveis (T8, T9);
    - QR só com `APP_PUBLIC_URL/e/{id}` (T6).
  - **11.2** `/e/{id}` abre o detalhe com login (T10).
  - **11.3** Inventário:
    - Gestor ou Técnico abre por unidade (T3, T11);
    - câmera com `@zxing/browser` sob demanda e HTTPS (T12, T13, Step 3);
    - digitação e leitor USB (T13);
    - confere unidade, local e colaborador (T2, T4);
    - local informado (T4, T13);
    - encerrar calcula os não encontrados (T5).
  - **11.4 / 12.4** Relatório com as seis categorias, em tela e em Excel/PDF (T5, T14).
  - **Aceite:**
    - leitura duplicada não conta duas vezes (T4);
    - encerramento lista os não lidos (T5);
    - câmera negada mostra a entrada manual (T13);
    - impressão com o layout configurado (T8, T9 e Step 5).
  - **Auditoria `inventario`** (T3–T5). **Migrações** na faixa 0600 (T1).
- [ ] **Step 7: Relatório da fase.** Inclui:
  - a justificativa das dependências novas (qrcode, @zxing/browser, @zxing/library como par obrigatório e @types/qrcode);
  - as linhas acrescentadas em arquivos compartilhados;
  - os pontos para a integração e para a F7 (seção abaixo);
  - a lembrança de que o `README.md` (passo 6 do plano mestre) deve citar "câmera do inventário exige HTTPS".

## Pontos de integração com outras fases

- **Plano mestre:** a tabela de fases diz que a F6 toca a `StockPage` ("seleção para etiquetas"). Vale a decisão posterior do contexto comum: a tela própria é `/etiquetas`, e a `StockPage` (F1) não é editada.
- **F1:** `GET /api/etiquetas/ids` reaproveita `InventoryDBManager.list_items(tipo=, status=, revenda=, search=, limit=None)`. A F1 pode acrescentar parâmetros, mas deve manter esses e o `limit=None`.
- **F2:**
  - O retorno a `/items/:id` depois do login depende de o login terminar em `/`. Se a F2 criar uma tela obrigatória de cadastro do 2FA, o redirecionamento só acontece quando o usuário chegar a `/` (o destino vale por 10 min).
  - A F2 já fixou `qrcode@1.5.4` + `@types/qrcode@1.5.5` (carregado sob demanda) para o QR do TOTP. A F6 usa as mesmas versões e o mesmo `import()` dinâmico. Na integração, `package.json`/lockfile ficam com uma entrada só de cada.
- **F4:** itens `Em Manutenção` entram como esperados ou não encontrados, com o status exibido. A F4 pode mexer na `ItemResponse`; a F6 não depende de nada além de `local_id`/`colaborador_id` da F0.
- **F7:**
  - A exportação "Inventário" da seção 6.F7.4 já é entregue aqui (`GET /api/inventarios/{id}/relatorio/exportar`).
  - A fonte da linha do tempo usa `InventarioDBManager.historico_do_item(item_id)`.
- **`_ROTULOS_DE_CAMPO`:** as chaves genéricas (`unidade_id`, `observacao`, `motivo`) podem ser acrescentadas também por outras fases. Na integração, manter uma só de cada.
- **`nginx.conf.template`:** a linha de `Permissions-Policy` é nova. Se outra fase acrescentar cabeçalhos de segurança, nenhum deles pode trazer `camera=()`.
