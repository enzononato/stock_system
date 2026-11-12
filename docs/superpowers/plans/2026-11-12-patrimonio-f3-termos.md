# Fase 3 — Termos e RevSign — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Empréstimo e devolução passam a usar o colaborador do cadastro (com e-mail obrigatório), cada entrega/devolução/pedido avulso vira uma solicitação de assinatura em `termo_solicitacoes`, enviada ao RevSign quando a integração está ligada (e manual quando não está), com retorno por webhook e consulta periódica, e o termo vigente/histórico aparece no equipamento e no colaborador.

**Architecture:** Módulos novos isolados: cliente do RevSign (`app/services/revsign_client.py`, interface + falso + HTTP), tabela e manager próprios (`termo_solicitacoes`, `app/db/termo_db.py`), três serviços (`termos.py` = solicitar/enviar/cancelar; `termos_confirmacao.py` = a confirmação que o upload manual e o RevSign compartilham; `termos_retorno.py` = webhook, consulta e jobs) e dois routers (`termos.py`, `integracoes_revsign.py`). O `InventoryDBManager` só ganha ganchos (gravar/limpar `colaborador_id`, gerar o .docx sem mudar status). No frontend, `api/termos.ts` + componentes em `components/termos/` reaproveitados pelas abas "Termos" (equipamento e colaborador), `LoanPage`, `ReturnPage` e `TermsPage`.

**Tech Stack:** FastAPI 0.115 + PyMySQL (SQL cru) + MySQL 8; `httpx` 0.28 para o RevSign (Documenso v2.21.1, API v2); React 18 + TS + TanStack Query 5 + Vitest/MSW.

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md` — seção 6 "F3", matriz da seção 3 (linhas 3.1, 3.4, 3.5, 6.1, 6.2, 6.4, 6.5, 6.6), seção 5 (`termo_solicitacoes`), seções 7–9. Contexto comum: `C:/Users/enzo.jz/dev/patrimonio/_tools/contexto-planejamento.md`. Interfaces da F0 consumidas: plano `2026-11-12-patrimonio-f0-base.md`.

**Ambiente:** worktree da fase `C:/Users/enzo.jz/dev/patrimonio/f3-termos` (branch `fase/f3-termos`, criada a partir de `feat/patrimonio-melhorias` com a F0 integrada; ajuste o caminho se o coordenador usar outro nome). Backend: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh C:/Users/enzo.jz/dev/patrimonio/f3-termos f3 <args do pytest>` (abreviado abaixo como `RBT <args>`). Frontend: em `frontend/`, `npx vitest run <arquivo>`, `npx tsc --noEmit`, `npm run build`.

## Global Constraints

- Migrações da F3 na faixa `0300–0399`, arquivo `NNNN_descricao.sql` + par `NNNN_descricao.down.sql`; só mudanças aditivas; cada comando termina com `;` no fim da linha (nunca `; -- comentário` na mesma linha, nem `;` dentro de literal).
- `REVSIGN_ENABLED` (padrão `false`), `REVSIGN_BASE_URL`, `REVSIGN_API_TOKEN`, `REVSIGN_WEBHOOK_SECRET` já existem em `app/core/config.py`; não criar variável nova.
- "Com `REVSIGN_ENABLED=false` ou falha, a solicitação fica como `manual` e o fluxo atual (baixar o termo, subir o PDF assinado) continua funcionando."
- "Webhook do RevSign validado por segredo; payload tratado como dado não confiável."
- "E-mail do colaborador — Obrigatório para receber equipamento; pode ser pessoal." / "o e-mail é **obrigatório** e validado antes de seguir. Os campos antigos continuam gravados para compatibilidade."
- "Integração isolada em `app/services/revsign_client.py` com a interface `RevSignClient` (`criar_documento`, `consultar`, `cancelar`, `reenviar`, `baixar_assinado`, `validar_webhook`) e uma implementação falsa para testes."
- "Uma solicitação nova cancela a pendente anterior do mesmo item." / "todo empréstimo (entrega a um colaborador, inclusive troca de aparelho ou de usuário) e toda devolução criam a solicitação correspondente sem ação extra do técnico."
- "nenhum fluxo atual (empréstimo, devolução, termos, baixa, estorno, relatórios) quebrado": `tests/test_loan_workflow.py`, `tests/test_termos.py`, `tests/test_audit_fluxos.py` e `tests/test_reversal.py` passam sem alteração.
- Toda escrita nova grava `audit.record(ctx, "termo", <id>, <acao>, antes, depois)`; ações: `criar`, `enviar`, `reenviar`, `cancelar`, `assinar`, `atualizar`. Jobs e webhook usam `AuditContext.sistema()`. CPF nunca entra cru (o `record` mascara chaves com "cpf"); e-mail do destinatário não entra na auditoria.
- Endpoints novos sob `/api/`; leitura e escrita de termos: `gestor_or_tecnico` (403 "Acesso restrito a Gestor ou Técnico."); só o webhook é público (autenticado pelo segredo).
- Telas da F3 usam os componentes da F0: `Pagination` + `usePageSize(tela, 25)` (25/50/100), `StatusPill`, `Button` com `loading`, `upper()` para unidade/revenda; texto azul só com `text-brand-text`. Não editar `StatusPill`, `Pagination`, `Button` nem os registros de abas.
- Arquivos compartilhados (`backend/app/main.py`, `backend/app/dependencies.py`, `backend/app/jobs.py`, `backend/.env.example`, `backend/requirements.txt`, `frontend/src/test/handlers.ts`) só recebem linhas novas.
- Dependência nova: `httpx==0.28.1` em `backend/requirements.txt` (cliente HTTP síncrono com timeout e `MockTransport` para testar o cliente sem rede; já está na imagem de teste e no `requirements-dev.txt`).
- Textos de interface, mensagens e commits em português; commits `tipo(escopo): descrição` + linha em branco + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus

- `REVSIGN_WEBHOOK_SECRET` vazio no servidor e o RevSign mandando `X-Documenso-Secret: ""` (é o que o Documenso envia quando o webhook não tem segredo) → 401, nunca aceito (teste na Task 8).
- O mesmo evento "assinado" chegando duas vezes, ou webhook e consulta periódica ao mesmo tempo → uma única confirmação: um `Confirmação Empréstimo` no histórico e uma linha `emprestimo/confirmar` na auditoria (testes na Task 8: repetição e trava ocupada).
- Colaborador desligado → empréstimo recusado com mensagem clara; devolução aceita (o seletor da devolução mostra desligados) (testes nas Tasks 6 e 13).
- Item que muda de situação entre o pedido e o envio/assinatura (estorno do empréstimo, devolução, agendado vencido) → solicitação cancelada com motivo, também no RevSign; assinatura que chega para solicitação cancelada é ignorada (testes nas Tasks 6, 8 e 9).
- RevSign fora do ar, erro HTTP ou timeout no envio → solicitação `manual` com motivo e o empréstimo segue normalmente; falha só no download do PDF mantém `enviado` para nova tentativa (testes nas Tasks 4 e 9).

## Decisões fixadas por este plano

| Tema | Decisão |
|---|---|
| Status pendentes | `PENDENTES = ("agendado", "enviado", "manual", "erro")`; finais: `assinado`, `recusado`, `expirado`, `cancelado` |
| `manual` | Assinatura pelo fluxo atual (baixar .docx, subir PDF). Sempre com `motivo` explicando por quê |
| `erro` | Assinado no RevSign, mas a confirmação no Patrimônio foi recusada (ex.: item mudou de status). O PDF fica guardado (`arquivo_assinado_key`) |
| Tipo pelo status do item | `Pendente` → `entrega`; `Indisponível` → `uso`; `Pendente Devolução` → `devolucao`. `POST /api/items/{id}/termos` deriva o tipo assim (para `Indisponível` é o `uso` do spec; nos outros dois é um novo pedido do mesmo termo) |
| Horário | Tudo de `termo_solicitacoes` em horário de Brasília sem fuso (`agora_local()`); `agendado_para` chega do navegador como `AAAA-MM-DDTHH:MM` local |
| `modelo_versao` | `"<nome do arquivo>@<8 primeiros hex do sha256 do arquivo>"`, ex. `termo.docx@1a2b3c4d`; `None` se o modelo não existe |
| Webhook | Só gatilho: identifica a solicitação por `externalId` (`patrimonio-termo-<id>`) ou id do documento e chama `sincronizar`, que consulta o RevSign (fonte da verdade) — o status do corpo nunca é usado |
| Concorrência | `sincronizar` e o job de agendados rodam dentro de `GET_LOCK('patrimonio_termo_<id>', 0)`; atualizações de status são condicionais (`WHERE status IN (...)`) |
| Confirmação automática | Mesmo caminho do upload manual; `history.operador = "RevSign"`; auditoria com `AuditContext.sistema()` |
| Payload antigo do empréstimo | Continua aceito. Sem `colaborador_id`, a solicitação tenta achar o colaborador pelo CPF (sem gravar `items.colaborador_id`) |
| Termo vigente | Última solicitação não cancelada do item, só enquanto o item está `Pendente`/`Indisponível`/`Pendente Devolução`; sem solicitação e com `Confirmação Empréstimo` ativa no histórico → termo "legado" (`id: null`, `legado: true`, `status: assinado`) |
| Estorno | Estornar `Empréstimo`/`Devolução` cancela as solicitações pendentes ligadas àquele `history_id` (e no RevSign); estornar `Empréstimo` limpa `items.colaborador_id` |

## Mapa de arquivos

Backend (criar): `migrations/0300_termo_solicitacoes.sql` (+ `.down.sql`), `app/core/tempo.py`, `app/services/revsign_client.py`, `app/db/termo_db.py`, `app/services/termos.py`, `app/services/termos_confirmacao.py`, `app/services/termos_retorno.py`, `app/schemas/termos.py`, `app/routers/termos.py`, `app/routers/integracoes_revsign.py`; testes `tests/fixtures_termos.py`, `tests/test_revsign_client.py`, `tests/test_termo_db.py`, `tests/test_termos_inventario.py`, `tests/test_termos_servico.py`, `tests/test_termos_confirmacao.py`, `tests/test_loan_colaborador.py`, `tests/test_termos_api.py`, `tests/test_revsign_webhook.py`, `tests/test_termos_jobs.py`.
Backend (alterar): `app/db/inventory_manager_db.py` (só métodos de empréstimo/devolução/termo e a linha do estorno do empréstimo), `app/routers/loans.py`, `app/schemas/loans.py`, `app/routers/history.py` (2 linhas), `app/dependencies.py`, `app/main.py`, `app/jobs.py`, `requirements.txt`, `.env.example`.
Frontend (criar): `src/api/termos.ts`, `src/lib/email.ts`, `src/components/termos/{TermoStatusPill,TermoVigenteCard,SolicitarAssinaturaDialog,AnexarAssinadoDialog,TermosLista,ColaboradorEmailField}.tsx`, `src/components/termos/colaboradorEmail.ts`, `src/features/item-detail/tabs/termos.tsx` (+ `.test.tsx`), `src/features/colaborador-detail/tabs/termos.tsx` (+ `.test.tsx`), `src/pages/ReturnPage.test.tsx`, `src/pages/TermsPage.test.tsx`.
Frontend (alterar): `src/api/loans.ts`, `src/pages/LoanPage.tsx` (+ teste), `src/pages/ReturnPage.tsx`, `src/pages/TermsPage.tsx`, `src/components/equipment/ConfirmacaoTermo.tsx`, `src/pages/ItemDetailPage.test.tsx` e `src/pages/ColaboradorDetailPage.test.tsx` (uma linha cada).

---

### Task 1: Cliente do RevSign (interface, falso e HTTP)

**Files:**
- Create: `backend/app/services/revsign_client.py`, `backend/tests/test_revsign_client.py`
- Modify: `backend/requirements.txt` (acrescentar, no bloco "Utilitários HTTP/rede", `httpx==0.28.1` com o comentário `# Cliente HTTP do RevSign (assinatura eletrônica)`)

**Interfaces:**
- Produces (todos em `app.services.revsign_client`):
  - `class RevSignError(Exception)`; `class WebhookInvalido(RevSignError)`; `class WebhookMalformado(RevSignError)`
  - `PDF_FAKE = b"%PDF-1.4\n% termo assinado (falso)\n"`; `CABECALHO_SEGREDO = "X-Documenso-Secret"`
  - `@dataclass(frozen=True) NovoDocumento(titulo: str, arquivo: bytes, nome_arquivo: str, signatario_nome: str, signatario_email: str, referencia_externa: str)`
  - `@dataclass(frozen=True) DocumentoCriado(documento_id: str)`
  - `Situacao = Literal["pendente", "assinado", "recusado", "expirado", "cancelado"]`; `@dataclass(frozen=True) StatusDocumento(situacao: Situacao, assinado_em: datetime | None = None, motivo: str | None = None)`
  - `@dataclass(frozen=True) EventoWebhook(tipo: str, documento_id: str | None, referencia_externa: str | None)`
  - `class RevSignClient(Protocol)`: `criar_documento(doc: NovoDocumento) -> DocumentoCriado`; `consultar(documento_id: str) -> StatusDocumento`; `cancelar(documento_id: str) -> None`; `reenviar(documento_id: str) -> None`; `baixar_assinado(documento_id: str) -> bytes`; `validar_webhook(headers: Mapping[str, str], corpo: bytes) -> EventoWebhook`
  - `class _ValidadorWebhook` (base de `HttpRevSignClient` e `FakeRevSignClient`, recebe `webhook_secret: str`): segredo configurado vazio → `WebhookInvalido("Segredo do webhook não configurado.")`; cabeçalho `X-Documenso-Secret` (busca sem caixa) ausente ou diferente (`hmac.compare_digest` em bytes UTF-8) → `WebhookInvalido("Segredo do webhook inválido.")`; JSON inválido, não-objeto ou sem `event` texto → `WebhookMalformado("Corpo do webhook inválido.")`; `documento_id = str(payload["id"])` quando `payload.id` é int/str, `referencia_externa = payload["externalId"]` quando texto.
  - `class FakeRevSignClient(_ValidadorWebhook)`: `__init__(webhook_secret: str = "")`; listas públicas `criados: list[NovoDocumento]`, `cancelados: list[str]`, `reenviados: list[str]`; `falhar_em: set[str]` (nomes dos métodos que devem levantar `RevSignError("Falha simulada em <método>.")`); ids `"1001"`, `"1002"`…; `marcar_assinado(documento_id: str, pdf: bytes = PDF_FAKE, assinado_em: datetime | None = None) -> None`; `marcar(documento_id: str, situacao: Situacao, motivo: str | None = None) -> None`; `consultar` de id desconhecido → `RevSignError("Documento não encontrado no RevSign.")`; `baixar_assinado` de documento não assinado → `RevSignError("Documento ainda não assinado.")`.
  - `class HttpRevSignClient(_ValidadorWebhook)`: `__init__(base_url: str, api_token: str, webhook_secret: str, *, transport: httpx.BaseTransport | None = None)`; constantes `API_PREFIXO = "/api/v2"`, `TIMEOUT = httpx.Timeout(15.0, connect=5.0)`, `TIMEOUT_DOWNLOAD = httpx.Timeout(30.0, connect=5.0)`, `MAX_PDF_BYTES = 20 * 1024 * 1024`, `MIME_DOCX`, `CAMPO_ASSINATURA = {"type": "SIGNATURE", "pageNumber": 1, "pageX": 10, "pageY": 80, "width": 35, "height": 8}`.
  - `get_revsign_client() -> RevSignClient` (override de teste ou `HttpRevSignClient(settings.REVSIGN_BASE_URL, settings.REVSIGN_API_TOKEN, settings.REVSIGN_WEBHOOK_SECRET)`); `usar_cliente(cliente: RevSignClient | None) -> None` (só testes).

- [ ] **Step 1: Escrever os testes** (`pytestmark = pytest.mark.unit`), com `httpx.MockTransport` gravando as requisições numa lista:
  - `test_criar_documento_envia_multipart_e_distribui`: o handler responde `{"envelopeId": "env_1", "id": 42}` ao create e `{}` ao distribute. Afirmar: `DocumentoCriado("42")`; 1ª requisição `POST https://rs.teste/api/v2/document/create` com `Authorization == "Bearer tok-123"`, corpo contendo `b'name="payload"'`, `b'name="file"; filename="termo.docx"'`, `'"externalId": "patrimonio-termo-5"'.encode()`, `b'"role": "SIGNER"'`, `b'"type": "SIGNATURE"'`, `'"email": "maria.souza@gmail.com"'.encode()`, `b'"language": "pt-BR"'`; 2ª `POST .../api/v2/document/distribute` com JSON `{"documentId": 42}`.
  - `test_distribuir_falha_apaga_rascunho_e_levanta`: distribute → 500; levanta `RevSignError`; 3ª requisição é `POST .../api/v2/document/delete` com `{"documentId": 42}`.
  - `test_consultar_mapeia_status` (parametrizado sobre o JSON de `GET .../api/v2/document/42`): `COMPLETED` + `completedAt` → `assinado` com `assinado_em`; `REJECTED` + `recipients[0].rejectionReason="Dados errados"` → `recusado`, `motivo == "Dados errados"`; `CANCELLED` → `cancelado`; `PENDING` → `pendente`; `PENDING` com destinatário `NOT_SIGNED` e `expiresAt` no passado → `expirado`.
  - `test_reenviar_redistribui_so_quem_nao_assinou`: GET devolve recipients `[{id: 7, signingStatus: "NOT_SIGNED"}, {id: 8, signingStatus: "SIGNED"}]` → POST `/api/v2/document/redistribute` com `{"documentId": 42, "recipients": [7]}`.
  - `test_cancelar_trata_404_como_sucesso`.
  - `test_baixar_assinado_exige_pdf`: `GET .../api/v2/document/42/download` com `b"%PDF-1.7..."` → bytes; com `b"<html>"` → `RevSignError`.
  - `test_erro_de_rede_vira_revsign_error_sem_token`: handler levanta `httpx.ConnectError("x")` → `RevSignError` cuja mensagem não contém `"tok-123"`; status 401 idem.
  - `test_sem_configuracao_levanta`: `HttpRevSignClient("", "", "s").criar_documento(...)` → `RevSignError("RevSign não configurado (REVSIGN_BASE_URL/REVSIGN_API_TOKEN).")`.
  - `test_validar_webhook` (tabela, usando `FakeRevSignClient(webhook_secret="s3")`): cabeçalho certo + corpo `{"event": "DOCUMENT_COMPLETED", "payload": {"id": 42, "externalId": "patrimonio-termo-5"}}` → `EventoWebhook("DOCUMENT_COMPLETED", "42", "patrimonio-termo-5")`; `x-documenso-secret` em minúsculas também vale; segredo errado e cabeçalho ausente → `WebhookInvalido`; `FakeRevSignClient(webhook_secret="")` com cabeçalho `""` → `WebhookInvalido`; corpo `b"nao-json"` e `b"[]"` → `WebhookMalformado`.
  - `test_fake_ciclo`: `criar_documento` devolve `"1001"`; `consultar` → `pendente`; `marcar_assinado("1001")` → `assinado` e `baixar_assinado == PDF_FAKE`; `falhar_em={"criar_documento"}` → `RevSignError`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_revsign_client.py -q` → FAIL (módulo inexistente).
- [ ] **Step 3: Implementar** `revsign_client.py`. Docstring do módulo lista TODAS as premissas do contrato, para ajuste quando a sessão do RevSign entregar o contrato: (1) API v2 do Documenso v2.21.1, rotas "document" (marcadas como deprecated em favor da Envelope API); (2) `POST /api/v2/document/create` multipart (`payload` JSON + `file` .docx, convertido pelo Gotenberg do RevSign) cria rascunho e `POST /document/distribute` envia o e-mail; (3) autenticação `Authorization: Bearer <REVSIGN_API_TOKEN>`; (4) campo de assinatura obrigatório (o Documenso recusa enviar sem campo `SIGNATURE`), posição fixa `CAMPO_ASSINATURA` na página 1 até o contrato dizer outra coisa; (5) signatário com e-mail pessoal depende da alteração do RevSign (spec, seção 2); (6) `GET /document/{id}` traz `status` (`DRAFT|PENDING|COMPLETED|REJECTED|CANCELLED`), `completedAt` e `recipients[]` (`id`, `signingStatus`, `signedAt`, `rejectionReason`, `expiresAt`); (7) `GET /document/{id}/download` devolve o PDF assinado; (8) `POST /document/delete` cancela/apaga; `POST /document/redistribute` reenvia; (9) webhook com corpo `{event, payload: {id, externalId, ...}, createdAt, webhookEndpoint}` e cabeçalho `X-Documenso-Secret`; eventos úteis `DOCUMENT_COMPLETED`, `DOCUMENT_REJECTED`, `DOCUMENT_CANCELLED`, `RECIPIENT_EXPIRED`; (10) o Documenso recusa URL de webhook privada: cadastrar a URL pública `<APP_PUBLIC_URL>/api/integrations/revsign/webhook`. Corpo do create: `{"title", "externalId", "recipients": [{"email", "name", "role": "SIGNER", "signingOrder": 1, "fields": [CAMPO_ASSINATURA]}], "meta": {"subject": f"Assinatura: {titulo}", "message": "Revalle · Patrimônio: assine o termo do equipamento sob sua responsabilidade.", "timezone": "America/Sao_Paulo", "language": "pt-BR"}}` serializado com `json.dumps(..., ensure_ascii=False)`. Qualquer `httpx.HTTPError` ou status ≥ 400 vira `RevSignError(f"RevSign respondeu {status}")`/`RevSignError("RevSign indisponível.")`, nunca com o token. PDF maior que `MAX_PDF_BYTES` → `RevSignError`.
- [ ] **Step 4: Rodar** `RBT tests/test_revsign_client.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(revsign): cliente do RevSign com interface, implementação falsa e HTTP`.

### Task 2: Tabela `termo_solicitacoes`, manager e fixtures de teste

**Files:**
- Create: `backend/migrations/0300_termo_solicitacoes.sql`, `backend/migrations/0300_termo_solicitacoes.down.sql`, `backend/app/core/tempo.py`, `backend/app/db/termo_db.py`, `backend/tests/fixtures_termos.py`, `backend/tests/test_termo_db.py`
- Modify: `backend/app/dependencies.py` (acrescentar `get_termo_db`)

**Interfaces:**
- Consumes: `ColaboradorDBManager.criar_manual(dados: dict) -> int`, `obter(id) -> dict | None` (F0 Task 6); fixtures `criar_item`, `criar_unidade`, `executar_sql`, `limpar_banco`, `_test_schema` (conftest); `FakeRevSignClient`, `usar_cliente`, `PDF_FAKE` (Task 1).
- Produces:
  - Tabela (texto exato do `.sql`):
    ```sql
    CREATE TABLE IF NOT EXISTS termo_solicitacoes (
        id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        item_id INT NOT NULL,
        colaborador_id INT NULL,
        tipo ENUM('entrega','devolucao','uso') NOT NULL,
        history_id INT NULL,
        modelo_versao VARCHAR(100) NULL,
        status ENUM('agendado','enviado','assinado','recusado','expirado','cancelado','erro','manual') NOT NULL,
        destinatario_nome VARCHAR(200) NULL,
        destinatario_email VARCHAR(200) NULL,
        solicitado_por VARCHAR(100) NOT NULL,
        solicitado_em DATETIME NOT NULL,
        agendado_para DATETIME NULL,
        enviado_em DATETIME NULL,
        assinado_em DATETIME NULL,
        cancelado_em DATETIME NULL,
        revsign_documento_id VARCHAR(64) NULL,
        documento_key VARCHAR(255) NULL,
        arquivo_assinado_key VARCHAR(255) NULL,
        motivo VARCHAR(500) NULL,
        atualizado_em DATETIME NOT NULL,
        UNIQUE KEY uq_termo_revsign_doc (revsign_documento_id),
        INDEX idx_termo_item (item_id, id),
        INDEX idx_termo_colaborador (colaborador_id),
        INDEX idx_termo_status (status, agendado_para),
        CONSTRAINT fk_termo_item FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE CASCADE,
        CONSTRAINT fk_termo_colaborador FOREIGN KEY (colaborador_id) REFERENCES colaboradores(id) ON DELETE SET NULL,
        CONSTRAINT fk_termo_history FOREIGN KEY (history_id) REFERENCES history(id) ON DELETE SET NULL
    ) CHARACTER SET utf8mb4;
    ```
    `.down.sql`: `DROP TABLE IF EXISTS termo_solicitacoes;`
  - `app/core/tempo.py`: `FUSO_BRASILIA = ZoneInfo("America/Sao_Paulo")`; `agora_local() -> datetime` (agora em Brasília, sem `tzinfo`).
  - `app/db/termo_db.py`: `PENDENTES = ("agendado", "enviado", "manual", "erro")`; `class TermoDBManager` com:
    - `criar(*, item_id: int, tipo: str, status: str, solicitado_por: str, solicitado_em: datetime, colaborador_id: int | None = None, history_id: int | None = None, agendado_para: datetime | None = None, destinatario_nome: str | None = None, destinatario_email: str | None = None) -> int`
    - `obter(termo_id: int) -> dict | None` — linha + `colaborador_nome` (JOIN `colaboradores`) + `item_descricao` (`CONCAT_WS(' ', i.tipo, i.brand, i.model)`); `obter_por_documento(revsign_documento_id: str) -> dict | None`
    - `atualizar(termo_id: int, *, somente_se_status: tuple[str, ...] | None = None, **campos) -> bool` — só colunas de `COLUNAS_ATUALIZAVEIS = {"status", "colaborador_id", "modelo_versao", "destinatario_nome", "destinatario_email", "agendado_para", "enviado_em", "assinado_em", "cancelado_em", "revsign_documento_id", "documento_key", "arquivo_assinado_key", "motivo"}` (outra → `ValueError`); sempre grava `atualizado_em = agora_local()`; com `somente_se_status` acrescenta `AND status IN (...)`; devolve se alguma linha mudou.
    - `pendentes_do_item(item_id: int, *, tipo: str | None = None) -> list[dict]` (status em `PENDENTES`, `ORDER BY id DESC`); `por_history(history_id: int) -> list[dict]`
    - `listar_por_item(item_id: int, *, limit: int, offset: int) -> tuple[list[dict], int]` e `listar_por_colaborador(colaborador_id: int, *, limit: int, offset: int) -> tuple[list[dict], int]` (`ORDER BY id DESC`, mesmas colunas de `obter`)
    - `vigente(item_id: int) -> dict | None` — regra da tabela "Decisões"; o legado vem da última `history` `Confirmação Empréstimo` com `is_reversed=0` e devolve `{"id": None, "item_id", "tipo": "entrega", "status": "assinado", "solicitado_por": operador, "solicitado_em": data_operacao, "assinado_em": data_operacao, "arquivo_assinado_key": termo_assinado_anexo, "legado": True}` (demais campos `None`); linhas normais levam `"legado": False`.
    - `agendados_vencidos(agora: datetime) -> list[dict]` (`status='agendado' AND agendado_para <= agora`); `enviados() -> list[dict]` (`status='enviado' AND revsign_documento_id IS NOT NULL`)
  - `app.dependencies.get_termo_db() -> TermoDBManager` (`@lru_cache`, mesmo padrão dos outros).
  - `tests/fixtures_termos.py` (os arquivos de teste importam daqui com `# noqa: F401`; **não** registrar no `conftest.py`): `SEGREDO_WEBHOOK = "segredo-webhook-de-teste-123"`; fixtures `revsign_fake(monkeypatch) -> FakeRevSignClient` (seta `settings.REVSIGN_ENABLED = True` e `settings.REVSIGN_WEBHOOK_SECRET = SEGREDO_WEBHOOK` com `monkeypatch.setattr`, `usar_cliente(fake)`, e `usar_cliente(None)` no teardown), `criar_colaborador(_test_schema, limpar_banco)` (fábrica via `ColaboradorDBManager().criar_manual` com padrões `cpf="52998224725"`, `nome="Maria Souza"`, `cargo="Analista Financeiro"`, `setor="Financeiro"`, `email="maria.souza@gmail.com"`, `situacao="ativo"`; devolve `obter(id)`), `colaborador(criar_colaborador)`, `unidade_juazeiro(criar_unidade)` (`criar_unidade(nome="Revalle Juazeiro")`, para o termo poder ser gerado); helpers `corpo_emprestimo_colaborador(item_id: int, colaborador_id: int, **extra) -> dict` (`center_cost="101 - Puxada"`, `revenda="Revalle Juazeiro"`, `date_issue` = hoje `dd/mm/aaaa`), `evento_webhook(tipo: str, documento_id: str, termo_id: int) -> dict` (`{"event": tipo, "payload": {"id": int(documento_id), "externalId": f"patrimonio-termo-{termo_id}", "status": "COMPLETED"}, "createdAt": "2026-11-12T10:00:00Z", "webhookEndpoint": "https://patrimonio.teste/api/integrations/revsign/webhook"}`), `postar_webhook(client, corpo: dict, segredo: str | None = SEGREDO_WEBHOOK)` (posta em `/api/integrations/revsign/webhook`, com o cabeçalho só quando `segredo is not None`), `linhas_audit(entidade: str, acao: str | None = None) -> list[dict]` (mesmo formato do `_linhas` de `test_audit_fluxos.py`).

- [ ] **Step 1: Testes** (`integration`) em `test_termo_db.py`:
  - `test_migracao_cria_tabela_com_enums`: `information_schema.columns` de `termo_solicitacoes` tem `status` com `'agendado','enviado','assinado','recusado','expirado','cancelado','erro','manual'` e `tipo` com `'entrega','devolucao','uso'`.
  - `test_criar_obter_e_atualizar`: cria para um item (`criar_item`) com colaborador; `obter` traz `colaborador_nome == "Maria Souza"` e `item_descricao == "Notebook Dell Latitude 5420"`; `atualizar(id, status="enviado", revsign_documento_id="1001")` → `True`; `atualizar(id, somente_se_status=("agendado",), status="cancelado")` → `False` e o status continua `enviado`; `atualizar(id, item_id=2)` → `ValueError`.
  - `test_pendentes_do_item_e_por_tipo` (três linhas: `manual`/entrega, `assinado`/entrega, `enviado`/uso → `pendentes_do_item` devolve as ids de `enviado` e `manual` nessa ordem; com `tipo="uso"` só a de `enviado`).
  - `test_vigente_ignora_cancelada_e_item_disponivel`: item `Indisponível` (via `executar_sql`) com uma `assinado` e depois uma `cancelado` → vigente é a `assinado`; item `Disponível` → `None`.
  - `test_vigente_legado_vem_do_historico` (`item_indisponivel` da conftest, sem solicitação) → `legado is True`, `status == "assinado"`, `arquivo_assinado_key == "termos_assinados/fixture.pdf"`, `solicitado_por == "fixture_teste"`.
  - `test_agendados_vencidos_e_enviados`: `agendado_para = agora + 1h` não aparece com `agora`, aparece com `agora + 61min`; `enviados()` só traz `enviado` com documento.
  - `test_listar_por_colaborador_pagina`: 3 linhas do colaborador, `limit=2, offset=0` → 2 linhas e `total == 3`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_termo_db.py -q` → FAIL.
- [ ] **Step 3: Implementar** migração, `tempo.py`, `termo_db.py`, `get_termo_db`, `fixtures_termos.py`.
- [ ] **Step 4: Rodar** `RBT tests/test_termo_db.py tests/test_migrations.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(termos): tabela de solicitações de assinatura e manager`.

### Task 3: Ganchos do empréstimo/devolução no `InventoryDBManager`

**Files:**
- Modify: `backend/app/db/inventory_manager_db.py` (`issue`, `confirm_return`, ramo `'Empréstimo'` de `reverse_history_entry`, `generate_loan_term_bytes`, `generate_return_term_bytes`; métodos novos)
- Test: `backend/tests/test_termos_inventario.py`

**Interfaces:**
- Produces:
  - `issue(self, pid, user, cpf, center_cost, cargo, setor, revenda, date_issue, logged_user: str, pessoa_juridica: bool = False, colaborador_id: int | None = None)` — o `UPDATE items` passa a gravar `colaborador_id=%s` (o valor recebido, `None` no payload antigo). Mensagens e retorno iguais.
  - `confirm_return(...)` — o `UPDATE items` acrescenta `colaborador_id=NULL`. `reverse_history_entry` — o `UPDATE` do ramo `'Empréstimo'` acrescenta `colaborador_id=NULL` (única mudança nesse método).
  - `ultimo_history_id(self, item_id: int, operation: str) -> int | None` (último não estornado, `ORDER BY data_operacao DESC, id DESC`).
  - `vincular_colaborador(self, item_id: int, colaborador_id: int) -> None`.
  - `render_loan_term(self, item_id: int, *, status_permitidos: tuple[str, ...] = ("Pendente",)) -> tuple[bool, bytes | str, str | None]` — corpo atual de `generate_loan_term_bytes`; status fora da lista → `"Este equipamento não está pendente de empréstimo."` quando `status_permitidos == ("Pendente",)`, senão `"Este equipamento não está com um colaborador."`. `generate_loan_term_bytes(item_id)` passa a ser `return self.render_loan_term(item_id)`.
  - `render_return_term(self, item_id: int, *, status_permitidos: tuple[str, ...] = ("Indisponível",)) -> tuple[bool, bytes | str, str | None]` — a parte de geração de `generate_return_term_bytes` (sem tocar no banco); status fora → mensagem atual `"Apenas itens com status 'Indisponível' podem ter termo de devolução gerado."` no padrão, senão `"Este equipamento não está com devolução pendente."`. `generate_return_term_bytes` = `render_return_term` + a transição atual (mesmas mensagens, mesmo nome de arquivo).

- [ ] **Step 1: Testes** (`integration`; importar `criar_colaborador`, `colaborador`, `unidade_juazeiro` de `tests.fixtures_termos`):
  - `test_issue_grava_colaborador_id` (`inv_manager.issue(..., colaborador_id=colaborador["id"])` → `find(...)["colaborador_id"] == colaborador["id"]`); `test_issue_sem_colaborador_grava_null`.
  - `test_confirm_return_limpa_colaborador_id` (com `forcar_devolucao_iniciada`).
  - `test_estorno_do_emprestimo_limpa_colaborador_id` (via `inv_manager.reverse_history_entry(ultimo_history_id(item, "Empréstimo"), "teste")`).
  - `test_ultimo_history_id_ignora_estornados`; `test_vincular_colaborador`.
  - `test_render_loan_term_indisponivel_so_quando_permitido` (`item_indisponivel` + `unidade_juazeiro`): padrão → `(False, "Este equipamento não está pendente de empréstimo.", None)`; com `("Indisponível",)` → `ok` e `conteudo[:2] == b"PK"`.
  - `test_render_return_term_nao_muda_status` (`item_indisponivel`; pula com `pytest.skip` se `termo_devolucao_disponivel` for falso): `ok`; status continua `Indisponível`; nenhuma linha `Devolução` nova em `history`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_termos_inventario.py -q` → FAIL.
- [ ] **Step 3: Implementar** (só os métodos listados; nenhuma mensagem existente muda).
- [ ] **Step 4: Rodar** `RBT tests/test_termos_inventario.py tests/test_termos.py tests/test_loan_workflow.py tests/test_reversal.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(emprestimo): vínculo do colaborador no item e geração do termo sem mudar status`.

### Task 4: Serviço de solicitação (solicitar, enviar, cancelar, reenviar)

**Files:**
- Create: `backend/app/services/termos.py`, `backend/tests/test_termos_servico.py`

**Interfaces:**
- Consumes: Tasks 1–3; `get_inventory_db`, `get_termo_db`, `get_colaborador_db` (`app.dependencies`); `ColaboradorDBManager.obter`, `obter_por_cpf` (F0); `storage` (`app.core.storage`); `audit.record`; `TERMO_MODELOS`, `TERMO_DEVOLUCAO_MODELOS`.
- Produces (`app.services.termos`):
  - `TipoTermo = Literal["entrega", "devolucao", "uso"]`; `STATUS_ITEM_POR_TIPO = {"entrega": "Pendente", "uso": "Indisponível", "devolucao": "Pendente Devolução"}`; `TIPO_POR_STATUS_ITEM` (inverso); `ROTULO_TIPO = {"entrega": "entrega", "devolucao": "devolução", "uso": "uso"}`; `PREFIXO_REFERENCIA = "patrimonio-termo-"`; `OPERADOR_REVSIGN = "RevSign"`.
  - Motivos (texto exato): `MOTIVO_DESLIGADO = "Integração com o RevSign desligada: baixe o termo e anexe o PDF assinado."`, `MOTIVO_SEM_COLABORADOR = "Colaborador não encontrado no cadastro: assinatura manual."`, `MOTIVO_SEM_EMAIL = "Colaborador sem e-mail cadastrado: assinatura manual."`, `MOTIVO_SITUACAO_MUDOU = "O equipamento mudou de situação antes do envio."`, `MOTIVO_ESTORNO = "Operação estornada no histórico."`; formatos `f"Não foi possível gerar o termo: {erro}"`, `f"Falha ao enviar ao RevSign ({erro}): assinatura manual."`, `f"Substituída pela solicitação #{novo_id}."`.
  - `class TermoErro(Exception)` (mensagem para o usuário; routers → 400).
  - `@dataclass(frozen=True) DocumentoGerado(conteudo: bytes, nome_arquivo: str, chave: str)`.
  - `referencia_externa(termo_id: int) -> str`; `versao_modelo(caminho: str | None) -> str | None`; `caminho_modelo(tipo: TipoTermo, revenda: str | None) -> str | None`.
  - `solicitar(item_id: int, tipo: TipoTermo, ctx: AuditContext, *, history_id: int | None = None, agendado_para: datetime | None = None, documento: DocumentoGerado | None = None) -> dict` — resolve o colaborador (`item.colaborador_id`, senão `obter_por_cpf(item.cpf)`), cria a linha (`status="agendado"` com data, senão `"manual"` provisório), cancela as outras pendentes do item (`cancelar_pendentes_do_item(..., exceto=novo_id)`), audita `termo/criar` (`depois`: `tipo`, `item_id`, `colaborador_id`, `history_id`, `status`, `agendado_para`) e, sem agendamento, chama `enviar`. Devolve `obter(novo_id)`.
  - `enviar(termo_id: int, ctx: AuditContext, *, documento: DocumentoGerado | None = None) -> dict` — só age em `agendado`/`manual`; item fora de `STATUS_ITEM_POR_TIPO[tipo]` → `cancelar(..., MOTIVO_SITUACAO_MUDOU)`; resolve de novo o colaborador (grava `colaborador_id`, `destinatario_nome`, `destinatario_email`); grava `modelo_versao`; gera o .docx quando `documento` é `None` (`render_loan_term`/`render_return_term` com `status_permitidos=(STATUS_ITEM_POR_TIPO[tipo],)`, salvo em `storage` na categoria `terms` — entrega/uso — ou `termos_devolucao`) e grava `documento_key`; depois, na ordem: desligado → `manual` + `MOTIVO_DESLIGADO`; geração falhou → `manual` + "Não foi possível gerar o termo: …"; sem colaborador → `MOTIVO_SEM_COLABORADOR`; sem e-mail → `MOTIVO_SEM_EMAIL`; `criar_documento(NovoDocumento(titulo=f"Termo de {ROTULO_TIPO[tipo]} — {tipo do item} {marca} {modelo} (patrimônio #{item_id})", ..., referencia_externa=referencia_externa(id)))` → `enviado` (`enviado_em`, `revsign_documento_id`, `motivo=None`); `RevSignError` → `manual` + "Falha ao enviar ao RevSign (…)". Atualização condicional `somente_se_status=("agendado", "manual")`; audita `termo/enviar` (`status`, `motivo`). Nunca propaga erro do RevSign.
  - `cancelar(termo_id: int, ctx: AuditContext, motivo: str) -> dict` — status fora de `PENDENTES` → `TermoErro("Esta solicitação não está pendente.")`; `enviado` com documento e integração ligada → `get_revsign_client().cancelar(...)` em melhor esforço (falha: log + `motivo` ganha ` (não foi possível cancelar no RevSign)`); grava `cancelado`, `cancelado_em`, `motivo`; audita `termo/cancelar`.
  - `cancelar_pendentes_do_item(item_id: int, ctx: AuditContext, motivo: str, *, exceto: int | None = None) -> list[int]`; `cancelar_por_estorno(history_id: int, ctx: AuditContext) -> list[int]` (pendentes de `por_history(history_id)`, motivo `MOTIVO_ESTORNO`).
  - `reenviar(termo_id: int, ctx: AuditContext) -> dict` — `enviado` → `get_revsign_client().reenviar(doc)` (`RevSignError` → `TermoErro(f"Não foi possível reenviar: {erro}")`); `manual` com integração ligada → `enviar(termo_id, ctx)`; `manual` com integração desligada → `TermoErro("Integração com o RevSign desligada: use a assinatura manual.")`; outro status → `TermoErro("Só é possível reenviar uma solicitação enviada ou em assinatura manual.")`. Audita `termo/reenviar`.
  - `para_resposta(linha: dict) -> dict` — campos de `TermoResponse` (Task 7): remove `documento_key`, `arquivo_assinado_key`, `revsign_documento_id`, `destinatario_nome`, `atualizado_em`; acrescenta `tem_documento = bool(documento_key)`, `tem_assinado = bool(arquivo_assinado_key)`, `legado = linha.get("legado", False)`.

- [ ] **Step 1: Testes** (`integration`; o empréstimo base é `inv_manager.issue(..., colaborador_id=colaborador["id"])` com `item_disponivel` + `unidade_juazeiro`; `ctx = AuditContext(1, "gestor.teste", "Gestor", "127.0.0.1")`):
  - `test_versao_modelo` (`unit`, `tmp_path`): arquivo com `b"abc"` → `f"modelo.docx@{hashlib.sha256(b'abc').hexdigest()[:8]}"`; caminho inexistente → `None`.
  - `test_solicitar_desligado_fica_manual_com_motivo`: `status == "manual"`, `motivo == MOTIVO_DESLIGADO`, `modelo_versao.startswith("termo.docx@")`, `documento_key.startswith("terms/")`, `colaborador_id == colaborador["id"]`.
  - `test_solicitar_ligado_envia_ao_revsign` (`revsign_fake`): `status == "enviado"`, `revsign_documento_id == "1001"`, `enviado_em` preenchido; `fake.criados[0]`: `signatario_email == "maria.souza@gmail.com"`, `nome_arquivo.endswith(".docx")`, `arquivo[:2] == b"PK"`, `referencia_externa == f"patrimonio-termo-{id}"`, `"Termo de entrega" in titulo`.
  - `test_sem_email_fica_manual` (`criar_colaborador(email=None)`) → `motivo == MOTIVO_SEM_EMAIL`, nada em `fake.criados`.
  - `test_payload_antigo_acha_colaborador_pelo_cpf` (issue sem `colaborador_id`, `cpf` igual ao do colaborador) → `colaborador_id == colaborador["id"]`; CPF desconhecido → `colaborador_id is None` e, com `revsign_fake`, `motivo == MOTIVO_SEM_COLABORADOR`.
  - `test_falha_do_revsign_fica_manual_e_nao_propaga` (`fake.falhar_em = {"criar_documento"}`) → `manual`, `motivo.startswith("Falha ao enviar ao RevSign (")`.
  - `test_falha_ao_gerar_termo_fica_manual` (sem `unidade_juazeiro`, com `revsign_fake`) → `motivo.startswith("Não foi possível gerar o termo: A unidade 'Revalle Juazeiro' não está cadastrada")`.
  - `test_nova_solicitacao_cancela_pendente_anterior_tambem_no_revsign` (`revsign_fake`): duas `solicitar(..., "entrega")` → a 1ª `cancelado`, `motivo == f"Substituída pela solicitação #{id2}."`, `cancelado_em` preenchido; `fake.cancelados == ["1001"]`; a 2ª `enviado` com `"1002"`.
  - `test_cancelar_nao_pendente_levanta` (`TermoErro` com "Esta solicitação não está pendente.").
  - `test_enviar_agendado_com_item_em_outra_situacao_cancela` (solicitação `uso` agendada; item volta a `Disponível` via `executar_sql`; `enviar` → `cancelado`, `motivo == MOTIVO_SITUACAO_MUDOU`).
  - `test_reenviar` (`revsign_fake`): `enviado` → `fake.reenviados == ["1001"]`; `manual` por falta de e-mail, depois `executar_sql` grava o e-mail → `reenviar` → `enviado`; integração desligada → `TermoErro` com a mensagem do desligado.
  - `test_auditoria_termo` → `linhas_audit("termo", "criar")` e `linhas_audit("termo", "enviar")` com `entidade_id == str(id)` e `username == "gestor.teste"`; nenhum `depois` contém `"maria.souza@gmail.com"`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_termos_servico.py -q` → FAIL.
- [ ] **Step 3: Implementar** `app/services/termos.py`.
- [ ] **Step 4: Rodar** `RBT tests/test_termos_servico.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(termos): solicitação de assinatura com envio ao RevSign e plano manual`.

### Task 5: Confirmação compartilhada (upload manual e RevSign)

**Files:**
- Create: `backend/app/services/termos_confirmacao.py`, `backend/tests/test_termos_confirmacao.py`
- Modify: `backend/app/routers/loans.py` (`confirm_loan` e `confirm_return` passam a chamar o serviço; `_auditar_etapa` e `_safe_filename` saem do router)

**Interfaces:**
- Consumes: Tasks 1–4; `audit.diff`, `audit.record`.
- Produces (`app.services.termos_confirmacao`):
  - `nome_seguro(nome: str) -> str` (o atual `_safe_filename`); `auditar_etapa(ctx: AuditContext, entidade: str, acao: str, item_id: int, antes_item: dict | None, extra: dict | None = None) -> None` (o atual `_auditar_etapa`, lendo o item com `get_inventory_db().find`). `routers/loans.py` passa a importar os dois.
  - `confirmar_emprestimo(item_id: int, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext, *, termo_id: int | None = None, origem: Literal["manual", "revsign"] = "manual") -> tuple[bool, str, str]` — salva em `termos_assinados` com o nome atual `f"termo_{item_id}_{AAAAmmddHHMMSS}_{nome_seguro(nome)}"`, chama `inv.confirm_loan(item_id, operador, chave)`; sucesso → `auditar_etapa(ctx, "emprestimo", "confirmar", ..., extra={"termo_assinado": chave})` e marca a solicitação (`termo_id`, ou a pendente mais recente `entrega` do item) como `assinado` (`assinado_em`, `arquivo_assinado_key=chave`, `motivo=None`, condicional a `PENDENTES`) + auditoria `termo/assinar` (`depois`: `status`, `origem`, `tipo`); se a marcada estava `enviado` e `origem == "manual"` → `cancelar` no RevSign em melhor esforço. Devolve `(ok, msg, chave)`.
  - A chave do arquivo é devolvida mesmo quando `confirm_loan`/`confirm_return` recusa (o arquivo já foi salvo, como hoje).
  - `confirmar_devolucao(...)` — mesma assinatura e regras, categoria `termos_devolucao_assinados`, nome `f"termo_devolucao_{item_id}_{ts}_{nome}"`, `inv.confirm_return`, auditoria `devolucao/confirmar`, solicitação `devolucao`.
  - `registrar_uso_assinado(termo_id: int, conteudo: bytes, nome_arquivo: str, ctx: AuditContext, *, origem: Literal["manual", "revsign"] = "manual") -> tuple[bool, str, str]` — salva em `termos_assinados` como `f"termo_uso_{item_id}_{ts}_{nome}"` e marca `assinado` (mesma auditoria e cancelamento remoto das outras); não mexe no item.
  - `concluir_assinatura(termo: dict, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext, *, origem: Literal["manual", "revsign"]) -> tuple[bool, str, str]` — despacha pelo `tipo`: `entrega` → `confirmar_emprestimo(..., termo_id=termo["id"], origem=origem)`, `devolucao` → `confirmar_devolucao(...)`, `uso` → `registrar_uso_assinado(...)`. Usada pelo upload manual (abaixo) e pelo RevSign (Task 8).
  - `anexar_assinado(termo_id: int, conteudo: bytes, nome_arquivo: str, operador: str, ctx: AuditContext) -> dict` — status fora de `("enviado", "manual", "erro")` → `TermoErro("Esta solicitação não aceita anexo.")`; senão `concluir_assinatura(..., origem="manual")`; `not ok` → `TermoErro(msg)`. Devolve `obter(termo_id)`.

- [ ] **Step 1: Testes** (`integration`):
  - `test_confirmacao_manual_marca_entrega_assinada` (empréstimo pelo serviço, integração desligada; `POST /api/loans/{id}/confirm` com `ARQUIVO_PDF`) → 200 com o `detail` atual; solicitação `assinado` com `arquivo_assinado_key` igual ao `termo_assinado_anexo` da `Confirmação Empréstimo`; `len(linhas_audit("emprestimo", "confirmar")) == 1`; uma linha `termo/assinar` com `depois["origem"] == "manual"`.
  - `test_upload_manual_com_solicitacao_enviada_cancela_no_revsign` (`revsign_fake`) → `assinado` e `fake.cancelados == ["1001"]`.
  - `test_confirmacao_sem_solicitacao_continua_funcionando` (`item_pendente` legado) → 200, nenhuma linha em `termo_solicitacoes`.
  - `test_confirmar_devolucao_marca_devolucao_assinada` (solicitação `devolucao` criada pelo serviço com `forcar_devolucao_iniciada`) → `assinado`, chave começa com `termos_devolucao_assinados/`.
  - `test_anexar_assinado_uso_grava_pdf` → `status == "assinado"`, `arquivo_assinado_key.startswith("termos_assinados/termo_uso_")`, item continua `Indisponível`; `anexar_assinado` numa `cancelado` → `TermoErro("Esta solicitação não aceita anexo.")`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_termos_confirmacao.py -q` → FAIL.
- [ ] **Step 3: Implementar** o serviço e trocar o corpo de `confirm_loan`/`confirm_return` por `ok, msg, _ = confirmar_*(item_id, content, signed_pdf.filename, current_user.username, ctx)` (400 com `msg` quando `not ok`; `_check_size` continua no router).
- [ ] **Step 4: Rodar** `RBT tests/test_termos_confirmacao.py tests/test_loan_workflow.py tests/test_audit_fluxos.py tests/test_uploads.py -q` → PASS.
- [ ] **Step 5: Commit** `refactor(termos): confirmação de empréstimo e devolução num caminho único`.

### Task 6: Empréstimo e devolução com colaborador do cadastro e solicitação automática

**Files:**
- Modify: `backend/app/schemas/loans.py`, `backend/app/routers/loans.py`, `backend/app/routers/history.py` (2 linhas), `backend/app/main.py` (`_ROTULOS_DE_CAMPO`: `"colaborador_id": "Colaborador"`)
- Test: `backend/tests/test_loan_colaborador.py`

**Interfaces:**
- Consumes: Tasks 3–5; `get_colaborador_db` (F0).
- Produces:
  - `LoanRequest`: `item_id: int`, `colaborador_id: Optional[int] = None`, `usuario: Optional[str] = None`, `cpf: Optional[str] = None`, `center_cost: str`, `cargo: Optional[str] = None`, `setor: Optional[str] = None`, `revenda: str`, `date_issue: str`, `pessoa_juridica: bool = False`; validador de CPF ignora `None`; `model_validator(mode="after")`: sem `colaborador_id` e com algum de `usuario`, `cpf`, `cargo`, `setor` ausente (`None`) → `ValueError("Informe o colaborador do cadastro ou nome, CPF, cargo e setor.")`.
  - `ReturnInitiateRequest(colaborador_id: Optional[int] = None)`.
  - `POST /api/loans` com `colaborador_id`: colaborador inexistente → 400 `"Colaborador não encontrado."`; `situacao == "desligado"` → 400 `"Colaborador desligado não pode receber equipamento."`; sem e-mail → 400 `"O colaborador precisa ter e-mail cadastrado para receber o termo."`; `usuario`/`cpf` = do cadastro (o corpo é ignorado), `cargo`/`setor` = do cadastro quando preenchidos, senão do corpo, senão `""`; `inv.issue(..., colaborador_id=...)`. Em qualquer payload, depois do sucesso e da auditoria atual: `termo = termos.solicitar(item_id, "entrega", ctx, history_id=inv.ultimo_history_id(item_id, "Empréstimo"))` dentro de `try/except Exception` (log; o empréstimo não falha). Resposta: `{"detail": msg, "termo": termos.para_resposta(termo) | None}`.
  - `POST /api/loans/{item_id}/return/initiate` com corpo opcional (`body: Optional[ReturnInitiateRequest] = Body(None)`): com `colaborador_id` → inexistente 400 `"Colaborador não encontrado."`, sem e-mail 400 `"O colaborador precisa ter e-mail cadastrado para receber o termo."`, item já vinculado a outro → 400 `"O equipamento está vinculado a outro colaborador."` (desligado é aceito); geração/transição atuais; só depois do sucesso, se o item não tinha vínculo, `inv.vincular_colaborador`; `termos.solicitar(item_id, "devolucao", ctx, history_id=inv.ultimo_history_id(item_id, "Devolução"), documento=DocumentoGerado(result, filename, storage_key))` em `try/except`. Resposta ganha `"termo"`.
  - `routers/history.py`: depois do `audit.record` do estorno, `termos.cancelar_por_estorno(history_id, ctx)` dentro de `try/except Exception` com `logger.exception` (o estorno nunca falha por causa disso).

- [ ] **Step 1: Testes** (`integration`):
  - `test_emprestimo_com_colaborador_preenche_dados_do_cadastro`: corpo `corpo_emprestimo_colaborador(item, colab, cargo="Outro", setor="TI", usuario="Ignorado")` → 200; item `assigned_to == "Maria Souza"`, `cpf == "52998224725"`, `colaborador_id == colab`; `history` `Empréstimo` com `cargo == "Analista Financeiro"`, `setor == "Financeiro"`; `resp.json()["termo"]["status"] == "manual"` e `["motivo"] == MOTIVO_DESLIGADO`.
  - `test_colaborador_sem_email_400`, `test_colaborador_desligado_400`, `test_colaborador_inexistente_400` (mensagens acima; item continua `Disponível`).
  - `test_payload_antigo_continua_aceito`: corpo antigo escrito no próprio teste (mesmos campos de `_corpo_emprestimo` em `test_loan_workflow.py`, CPF `11144477735`; não importar de outro módulo de teste) → 200, mesmo `detail`, `termo["tipo"] == "entrega"`, `termo["colaborador_id"] is None`.
  - `test_sem_colaborador_e_sem_campos_antigos_422`: `{"item_id", "center_cost", "revenda", "date_issue"}` → 422 com `detail == "Informe o colaborador do cadastro ou nome, CPF, cargo e setor."`.
  - `test_emprestimo_com_revsign_envia_para_email_pessoal` (`revsign_fake`, `unidade_juazeiro`) → `termo["status"] == "enviado"`, `fake.criados[0].signatario_email == "maria.souza@gmail.com"`.
  - `test_devolucao_confirmada_limpa_colaborador_id`.
  - `test_devolucao_de_item_legado_vincula_colaborador_escolhido` (`item_indisponivel`; pula sem modelo de devolução): `json={"colaborador_id": colab}` → 200, item `colaborador_id == colab`, `termo["tipo"] == "devolucao"`, `termo["tem_documento"] is True`.
  - `test_devolucao_com_colaborador_diferente_do_vinculado_400`.
  - `test_devolucao_aceita_colaborador_desligado` (Review Focus): colaborador `situacao="desligado"` com e-mail → 200.
  - `test_devolucao_cancela_uso_pendente` (`revsign_fake`; "troca de usuário/aparelho"): `uso` enviado (`termos.solicitar(..., "uso", ...)`), depois devolução → `uso` `cancelado` com `"Substituída pela solicitação #"`, `"1001" in fake.cancelados`, nova `devolucao` `enviado`.
  - `test_estorno_do_emprestimo_cancela_solicitacao_e_no_revsign` (`revsign_fake`): estorno da linha `Empréstimo` (`json={"password": "SenhaForte#123"}`) → 200; solicitação `cancelado`, `motivo == MOTIVO_ESTORNO`, `fake.cancelados == ["1001"]`; item `Disponível` e `colaborador_id is None`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_loan_colaborador.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar a suíte inteira:** `RBT -q -m "not mudanca_esperada"` → verde (inclui `test_loan_workflow.py`, `test_termos.py`, `test_audit_fluxos.py`, `test_reversal.py` sem alteração).
- [ ] **Step 5: Commit** `feat(emprestimo): colaborador do cadastro com e-mail e solicitação de termo automática`.

### Task 7: API de termos (equipamento, colaborador e ações)

**Files:**
- Create: `backend/app/schemas/termos.py`, `backend/app/routers/termos.py`, `backend/tests/test_termos_api.py`
- Modify: `backend/app/main.py` (import + `app.include_router(termos.router)`; `_ROTULOS_DE_CAMPO`: `"agendado_para": "Data agendada"`)

**Interfaces:**
- Consumes: Tasks 2, 4, 5.
- Produces:
  - Schemas: `TipoTermo`, `StatusTermo` (Literals); `TermoResponse(id: int | None, item_id: int, item_descricao: str | None = None, colaborador_id: int | None = None, colaborador_nome: str | None = None, tipo: TipoTermo, history_id: int | None = None, modelo_versao: str | None = None, status: StatusTermo, destinatario_email: str | None = None, solicitado_por: str, solicitado_em: datetime, agendado_para: datetime | None = None, enviado_em: datetime | None = None, assinado_em: datetime | None = None, cancelado_em: datetime | None = None, motivo: str | None = None, tem_documento: bool = False, tem_assinado: bool = False, legado: bool = False)`; `TermosDoItemResponse(revsign_ativo: bool, vigente: TermoResponse | None, historico: Paginated[TermoResponse])`; `SolicitarTermoRequest(agendado_para: datetime | None = None)` (valor com fuso é convertido para Brasília e fica sem fuso); `CancelarTermoRequest(motivo: str | None = Field(None, max_length=300))`; `ArquivosTermoResponse(documento_url: str | None, assinado_url: str | None)`.
  - Router `APIRouter(tags=["termos"])`, todos `gestor_or_tecnico`, erros `TermoErro` → 400:
    - `GET /api/items/{item_id}/termos?limit=25&offset=0` (`limit` 1–100) → `TermosDoItemResponse`; item inexistente → 404 `"Equipamento não encontrado."`.
    - `POST /api/items/{item_id}/termos` (201) → `TermoResponse`; tipo por `TIPO_POR_STATUS_ITEM` (fora → 400 `"Só é possível solicitar assinatura de equipamento com colaborador."`); `agendado_para` com integração desligada → 400 `"Agendamento disponível só com a integração do RevSign ligada."`; `agendado_para <= agora_local()` → 400 `"A data agendada precisa ser no futuro."`; `history_id` = `ultimo_history_id` de `Empréstimo` (entrega) ou `Devolução` (devolucao), `None` para uso.
    - `GET /api/colaboradores/{colaborador_id}/termos?limit=25&offset=0` → `Paginated[TermoResponse]`; inexistente → 404 `"Colaborador não encontrado."`.
    - `POST /api/termos/{termo_id}/cancelar` (corpo opcional; motivo padrão `f"Cancelada por {usuario}."`), `POST /api/termos/{termo_id}/reenviar`, `POST /api/termos/{termo_id}/assinado` (multipart `signed_pdf`, limite 20 MB → 413 com a mensagem atual) → `TermoResponse`; termo inexistente → 404 `"Solicitação de termo não encontrada."`.
    - `GET /api/termos/{termo_id}/arquivos` → `ArquivosTermoResponse` com `storage.get_url(chave)` quando a chave existe no storage, senão `None`.

- [ ] **Step 1: Testes** (`integration`):
  - `test_termos_do_item_traz_vigente_e_historico` (`revsign_fake`, empréstimo com colaborador) → `revsign_ativo is True`; `vigente["tipo"] == "entrega"`, `["status"] == "enviado"`, `["modelo_versao"].startswith("termo.docx@")`, `["solicitado_por"] == "gestor.teste"`, `["colaborador_nome"] == "Maria Souza"`; `historico["total"] == 1`; nenhuma chave `documento_key`/`revsign_documento_id` na resposta.
  - `test_termos_item_inexistente_404`; `test_aprendiz_recebe_403` (`client_aprendiz`, `"Acesso restrito a Gestor ou Técnico."`).
  - `test_vigente_legado` (`item_indisponivel`) → `vigente["legado"] is True`, `["id"] is None`, `["status"] == "assinado"`.
  - `test_solicitar_uso_em_item_indisponivel` (`revsign_fake`) → 201, `tipo == "uso"`, `status == "enviado"`; a 2ª chamada cancela a 1ª.
  - `test_solicitar_em_item_disponivel_400`; `test_agendar_com_integracao_desligada_400`; `test_agendar_no_passado_400` (com `revsign_fake`).
  - `test_cancelar_e_reenviar` (`revsign_fake`): `POST .../cancelar` com `{"motivo": "Pedido duplicado"}` → `cancelado`, `motivo == "Pedido duplicado"`; cancelar de novo → 400 `"Esta solicitação não está pendente."`.
  - `test_anexar_assinado_e_arquivos`: `uso` manual (integração desligada) → `POST .../assinado` → `tem_assinado is True`; `GET .../arquivos` → `assinado_url.startswith("/api/documents/files/termos_assinados/")`, `documento_url.startswith("/api/documents/files/terms/")`.
  - `test_anexar_assinado_entrega_confirma_emprestimo` → item `Indisponível`.
  - `test_termos_do_colaborador` → itens com `item_descricao == "Notebook Dell Latitude 5420"`, `total` certo; colaborador inexistente 404.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_termos_api.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `RBT tests/test_termos_api.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(termos): API do termo vigente, histórico e ações de assinatura`.

### Task 8: Retorno do RevSign (sincronização e webhook)

**Files:**
- Create: `backend/app/services/termos_retorno.py`, `backend/app/routers/integracoes_revsign.py`, `backend/tests/test_revsign_webhook.py`
- Modify: `backend/app/main.py` (import + `app.include_router(integracoes_revsign.router)`), `backend/.env.example` (acrescentar, abaixo de `REVSIGN_WEBHOOK_SECRET=`, o comentário `# URL a cadastrar no webhook do RevSign: <APP_PUBLIC_URL>/api/integrations/revsign/webhook (eventos DOCUMENT_COMPLETED, DOCUMENT_REJECTED, DOCUMENT_CANCELLED, RECIPIENT_EXPIRED)`)

**Interfaces:**
- Consumes: Tasks 1, 2, 4, 5.
- Produces:
  - `termos_retorno.trava(termo_id: int) -> ContextManager[bool]` — `GET_LOCK('patrimonio_termo_<id>', 0)` numa conexão própria do pool mantida durante o bloco; `RELEASE_LOCK` só se pegou.
  - `sincronizar(termo_id: int) -> str` — resultados: `"em_processamento"` (trava ocupada), `"ja_processado"` (status ≠ `enviado`), `"falha_consulta"` (`RevSignError` no `consultar`: grava só `motivo`), `"pendente"`, `"assinado"`, `"erro"`, `"falha_download"`, `"recusado"`, `"expirado"`, `"cancelado"`. `assinado` → `baixar_assinado` (falha: mantém `enviado`, `motivo = "Falha ao baixar o PDF assinado do RevSign; nova tentativa na próxima consulta."`, devolve `"falha_download"`) → `concluir_assinatura(termo, pdf, f"revsign_{documento_id}.pdf", OPERADOR_REVSIGN, AuditContext.sistema(), origem="revsign")` (Task 5); confirmação recusada → status `erro` (condicional a `enviado`), `motivo = msg`, `arquivo_assinado_key` = chave salva. `recusado`/`expirado`/`cancelado` → grava o status e o `motivo` vindo do RevSign; auditoria `termo/atualizar` com `AuditContext.sistema()`.
  - `processar_evento(evento: EventoWebhook) -> str` — acha a solicitação por `referencia_externa` (`patrimonio-termo-<id>`; id do documento, se vier, precisa bater com `revsign_documento_id`) ou por `obter_por_documento`; sem match → `"nao_encontrado"`; senão `sincronizar(id)`.
  - `POST /api/integrations/revsign/webhook` (sem login): integração desligada → 404 `"Integração com o RevSign desligada."`; corpo > 1 MB → 413 `"Corpo do webhook muito grande."`; `WebhookInvalido` → 401 `"Webhook não autorizado."`; `WebhookMalformado` → 400 `"Corpo do webhook inválido."`; senão `{"ok": True, "resultado": <str>}` (processamento em `run_in_threadpool`).

- [ ] **Step 1: Testes** (`integration`, todos com `revsign_fake`, `unidade_juazeiro`, `colaborador`):
  ```python
  def test_webhook_assinado_conclui_emprestimo_e_grava_pdf(client, client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake):
      item_id = item_disponivel["id"]
      termo = client_gestor.post("/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador["id"])).json()["termo"]
      revsign_fake.marcar_assinado("1001")
      resp = postar_webhook(client, evento_webhook("DOCUMENT_COMPLETED", "1001", termo["id"]))  # client SEM login
      assert resp.status_code == 200 and resp.json() == {"ok": True, "resultado": "assinado"}
      assert client_gestor.get(f"/api/items/{item_id}").json()["status"] == "Indisponível"
      # history: uma 'Confirmação Empréstimo', operador "RevSign", termo_assinado_anexo começa com "termos_assinados/" e storage.read(chave) == PDF_FAKE
      # solicitação: status "assinado", assinado_em preenchido; auditoria emprestimo/confirmar e termo/assinar com username "sistema"
  ```
  - `test_mesmo_evento_duas_vezes_confirma_uma_vez` → 2º `resultado == "ja_processado"`; exatamente uma `Confirmação Empréstimo` e uma `emprestimo/confirmar`.
  - `test_trava_ocupada_nao_processa` (Review Focus): dentro de `with get_cursor(commit=False) as cur: cur.execute("SELECT GET_LOCK(%s, 0)", (f"patrimonio_termo_{id}",))` → `sincronizar(id) == "em_processamento"` e item continua `Pendente`.
  - `test_segredo_errado_401`, `test_sem_cabecalho_401` (`postar_webhook(..., segredo=None)`).
  - `test_segredo_vazio_no_servidor_recusa_tudo` (Review Focus): `monkeypatch.setattr(settings, "REVSIGN_WEBHOOK_SECRET", "")` e fake recriado com `webhook_secret=""`; cabeçalho `""` → 401.
  - `test_integracao_desligada_404` (sem `revsign_fake`); `test_corpo_invalido_400` (`content=b"nao-json"`).
  - `test_documento_desconhecido_ignorado` → 200, `resultado == "nao_encontrado"`.
  - `test_recusado_marca_recusado` (`fake.marcar("1001", "recusado", "Dados errados")`) → `recusado`, `motivo == "Dados errados"`, item continua `Pendente`.
  - `test_assinatura_de_solicitacao_cancelada_e_ignorada` (cancelar pela API; depois assinado + webhook) → `ja_processado`, item `Pendente`.
  - `test_confirmacao_recusada_vira_erro_com_pdf_guardado` (item forçado a `Disponível` via `executar_sql`) → `erro`, `motivo == "Apenas itens com status 'Pendente' podem ser confirmados."`, `GET /api/items/{id}/termos` mostra `tem_assinado is True` no histórico.
  - `test_devolucao_assinada_pelo_revsign_conclui_devolucao` (pula sem modelo de devolução) → item `Disponível`, `colaborador_id is None`, chave em `termos_devolucao_assinados/`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_revsign_webhook.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `RBT tests/test_revsign_webhook.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(revsign): webhook autenticado e sincronização idempotente das assinaturas`.

### Task 9: Jobs do agendador (envio agendado e consulta de reserva)

**Files:**
- Modify: `backend/app/services/termos_retorno.py`, `backend/app/jobs.py`
- Test: `backend/tests/test_termos_jobs.py`

**Interfaces:**
- Consumes: `registrar_job`, `jobs_registrados`, `limpar_jobs` (`app.services.scheduler`, F0 Task 7); `registrar_jobs()` (`app.jobs`).
- Produces:
  - `enviar_agendados(agora: datetime | None = None) -> int` — integração desligada → `0`; para cada `agendados_vencidos(agora or agora_local())`: dentro de `trava(id)`, relê e, se ainda `agendado`, `termos.enviar(id, AuditContext.sistema())`; erro de uma solicitação é logado e não para as outras; devolve quantas saíram de `agendado`.
  - `consultar_enviados() -> int` — integração desligada → `0`; `sincronizar` em cada `enviados()` com `try/except` por item; devolve quantas mudaram de status (resultado fora de `pendente`, `em_processamento`, `falha_download`, `falha_consulta`, `ja_processado`).
  - Em `registrar_jobs()`, acrescentar ao fim do corpo:
    ```python
    if settings.REVSIGN_ENABLED:
        from app.services import termos_retorno
        registrar_job("termos_envios_agendados", termos_retorno.enviar_agendados, a_cada_segundos=60)
        registrar_job("termos_consulta_revsign", termos_retorno.consultar_enviados, a_cada_segundos=600)
    ```

- [ ] **Step 1: Testes** (`integration`, `revsign_fake`):
  - `test_envio_agendado_sai_no_tique_certo` (aceite): item `Indisponível` com colaborador (empréstimo + confirmação manual); `agora = agora_local()`; `POST /api/items/{id}/termos` com `{"agendado_para": (agora + timedelta(hours=1)).strftime("%Y-%m-%dT%H:%M")}` → 201 `agendado`, `len(fake.criados) == 1` (só a entrega); `enviar_agendados(agora + timedelta(minutes=30)) == 0` e continua `agendado`; `enviar_agendados(agora + timedelta(minutes=61)) == 1`, status `enviado`, `len(fake.criados) == 2`.
  - `test_agendado_de_item_que_mudou_de_situacao_e_cancelado` (item forçado a `Disponível`) → `cancelado`, `motivo == MOTIVO_SITUACAO_MUDOU`.
  - `test_consulta_periodica_conclui_enviado` → `marcar_assinado("1001")`; `consultar_enviados() == 1`; item `Indisponível`.
  - `test_falha_no_download_mantem_enviado_e_tenta_de_novo` (Review Focus): `fake.falhar_em = {"baixar_assinado"}` → `consultar_enviados() == 0`, status `enviado`, `motivo` começa com `"Falha ao baixar o PDF assinado"`; limpa `falhar_em` → `consultar_enviados() == 1`, `assinado`.
  - `test_jobs_desligados_nao_fazem_nada` (sem `revsign_fake`): ambos devolvem `0`.
  - `test_registrar_jobs_so_com_revsign_ligado`: `limpar_jobs()`; com `REVSIGN_ENABLED=False` → `registrar_jobs()` não registra `termos_envios_agendados` nem `termos_consulta_revsign`; `limpar_jobs()` e com `True` → os dois nomes em `jobs_registrados()`; `limpar_jobs()` no fim.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_termos_jobs.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar a suíte inteira:** `RBT -q -m "not mudanca_esperada"` → verde.
- [ ] **Step 5: Commit** `feat(termos): envio agendado e consulta periódica ao RevSign no agendador`.

### Task 10: Frontend — API de termos, componentes e aba "Termos" do equipamento

**Files:**
- Create: `frontend/src/api/termos.ts`, `frontend/src/components/termos/TermoStatusPill.tsx`, `frontend/src/components/termos/TermoVigenteCard.tsx`, `frontend/src/components/termos/SolicitarAssinaturaDialog.tsx`, `frontend/src/components/termos/AnexarAssinadoDialog.tsx`, `frontend/src/components/termos/TermosLista.tsx`, `frontend/src/features/item-detail/tabs/termos.tsx`, `frontend/src/features/item-detail/tabs/termos.test.tsx`
- Modify: `frontend/src/pages/ItemDetailPage.test.tsx` (só a linha `expect(tabs.map((t) => t.textContent)).toEqual(['Dados', 'Histórico'])` vira `expect(tabs.map((t) => t.textContent)).toEqual(expect.arrayContaining(['Dados', 'Histórico']))` — linha idêntica deve ser usada pelas outras fases que acrescentam abas)

**Interfaces:**
- Consumes: contratos da Task 7; `StatusPill`, `Pagination`, `usePageSize`, `Button loading`, `formatDateTime`, `downloadAuthenticated`, `downloadSignedTerm` (`api/loans.ts`), `getErrorMessage`, `Dialog`, `FileUpload`, `toast`.
- Produces:
  - `api/termos.ts`: `type TermoTipo = 'entrega' | 'devolucao' | 'uso'`; `type TermoStatus = 'agendado' | 'enviado' | 'assinado' | 'recusado' | 'expirado' | 'cancelado' | 'erro' | 'manual'`; `interface TermoSolicitacao` (campos de `TermoResponse`, datas como `string`, opcionais `| null`); `interface TermosDoItem { revsign_ativo: boolean; vigente: TermoSolicitacao | null; historico: Paginated<TermoSolicitacao> }`; `STATUS_TERMO_PENDENTES: TermoStatus[] = ['agendado', 'enviado', 'manual', 'erro']`; `TIPO_TERMO_ROTULO = { entrega: 'Termo de entrega', devolucao: 'Termo de devolução', uso: 'Termo de uso' }`; `STATUS_TERMO_ROTULO = { agendado: 'Agendado', enviado: 'Enviado', assinado: 'Assinado', recusado: 'Recusado', expirado: 'Expirado', cancelado: 'Cancelado', erro: 'Erro', manual: 'Assinatura manual' }`; funções `getTermosDoItem(itemId: number, params: { limit: number; offset: number }): Promise<TermosDoItem>`, `solicitarTermo(itemId: number, data: { agendado_para?: string }): Promise<TermoSolicitacao>`, `getTermosDoColaborador(colaboradorId: number, params: { limit: number; offset: number }): Promise<Paginated<TermoSolicitacao>>`, `cancelarTermo(id: number, motivo?: string)`, `reenviarTermo(id: number)`, `anexarTermoAssinado(id: number, pdf: File)` (campo `signed_pdf`), `getArquivosTermo(id: number): Promise<{ documento_url: string | null; assinado_url: string | null }>`.
  - `TermoStatusPill({ status }: { status: TermoStatus })` → `<StatusPill status={STATUS_TERMO_ROTULO[status]} tone={status === 'manual' ? 'pending' : undefined} />`.
  - `TermoVigenteCard({ itemId, itemStatus, vigente, revsignAtivo }: { itemId: number; itemStatus?: string | null; vigente: TermoSolicitacao | null; revsignAtivo: boolean })` — título "Termo vigente"; sem vigente: "Nenhum termo vigente para este equipamento."; mostra Tipo, Versão do modelo (`—` se nula), Status, "Solicitado por <quem> em <data>", Enviado em, Agendado para, Assinado em, Destinatário, Motivo; ações: "Solicitar assinatura" (abre `SolicitarAssinaturaDialog`; desabilitado com o texto "Disponível só para equipamento com colaborador." quando `itemStatus` não é `Pendente`/`Indisponível`/`Pendente Devolução`), "Reenviar" (`enviado`, ou `manual` com `revsignAtivo`), "Cancelar" (status em `STATUS_TERMO_PENDENTES`), "Anexar assinado" (`enviado`/`manual`/`erro`, abre `AnexarAssinadoDialog`), "Baixar termo" (`tem_documento`), "Ver assinado" (`tem_assinado`; legado usa `downloadSignedTerm(itemId)`). Mutações com `loading` e invalidação de `['item-termos', itemId]` e `['items']`.
  - `SolicitarAssinaturaDialog({ itemId, open, onOpenChange, revsignAtivo, temPendente })` — rádios "Enviar agora" / "Agendar envio" (o segundo desabilitado sem RevSign, com o texto "Agendamento exige a integração com o RevSign ligada."), `input type="datetime-local"` com `aria-label="Data e hora do envio"`, aviso "A solicitação pendente deste equipamento será cancelada." quando `temPendente`, botão "Solicitar" (`loading`); toasts: agendado → `Envio agendado para ${formatDateTime(agendado_para)}.`; enviado → `Termo enviado para assinatura em ${destinatario_email}.`; manual → `Solicitação registrada para assinatura manual: ${motivo}`.
  - `AnexarAssinadoDialog({ termo, open, onOpenChange })` — `FileUpload` + "Anexar" (`loading`).
  - `TermosLista({ termos, mostrarItem = false }: { termos: TermoSolicitacao[]; mostrarItem?: boolean })` — tabela com rolagem horizontal: (Equipamento com link `/items/<id>?aba=termos` e `text-brand-text`, quando `mostrarItem`) | Tipo | Status | Solicitado por / em | Enviado em | Assinado em | Motivo.
  - Aba `features/item-detail/tabs/termos.tsx`: `{ id: 'termos', label: 'Termos', order: 30, roles: ['Gestor', 'Técnico'], Component }`; `usePageSize('item-termos')`, página 1-based (volta a 1 ao trocar tamanho/equipamento, corrige quando passa da última), query `['item-termos', item.id, page, pageSize]`; renderiza `TermoVigenteCard` + "Histórico de solicitações" (`TermosLista` + `Pagination`); vazio: "Nenhuma solicitação de assinatura registrada."

- [ ] **Step 1: Testes** (`termos.test.tsx`, MSW, renderizando `tab.Component` com um `Item` `{ id: 12, status: 'Indisponível', ... }` + `ToastContainer`):
  - `mostra o termo vigente com tipo, versão, status, solicitante e assinatura` → textos "Termo de uso", "termo.docx@1a2b3c4d", "Assinado", `/Solicitado por tecnico\.teste/` e `/12\/11\/2026/` (data de `assinado_em`); a requisição a `/api/items/12/termos` leva `limit=25&offset=0`.
  - `solicitar com agendamento envia agendado_para` (`revsign_ativo: true`): "Solicitar assinatura" → "Agendar envio" → `fireEvent.change(getByLabelText('Data e hora do envio'), { target: { value: '2026-11-20T09:30' } })` → "Solicitar" → corpo do POST `/api/items/12/termos` igual a `{ agendado_para: '2026-11-20T09:30' }`; toast casa `/Envio agendado para 20\/11\/2026/`.
  - `sem RevSign o agendamento fica indisponível` → rádio "Agendar envio" desabilitado e o texto do aviso aparece.
  - `cancelar solicitação pendente` (vigente `enviado`, id 31) → POST `/api/termos/31/cancelar` chamado; refaz o GET.
  - `equipamento disponível não permite solicitar` (`status: 'Disponível'`) → botão desabilitado e "Disponível só para equipamento com colaborador.".
  - `histórico paginado` (2 itens) → "Mostrando 1–2 de 2" e pílulas "Enviado" e "Cancelado".
  - `ItemDetailPage.test.tsx` continua passando com a linha ajustada.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/item-detail/tabs/termos.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(termos): aba Termos no detalhe do equipamento`.

### Task 11: Frontend — aba "Termos" do colaborador

**Files:**
- Create: `frontend/src/features/colaborador-detail/tabs/termos.tsx`, `frontend/src/features/colaborador-detail/tabs/termos.test.tsx`
- Modify: `frontend/src/pages/ColaboradorDetailPage.test.tsx` (só a linha `toEqual(['Dados', 'Equipamentos'])` vira `toEqual(expect.arrayContaining(['Dados', 'Equipamentos']))`)

**Interfaces:**
- Consumes: `getTermosDoColaborador`, `TermosLista` (Task 10).
- Produces: `{ id: 'termos', label: 'Termos', order: 30, roles: ['Gestor', 'Técnico'], Component }` — `usePageSize('colaborador-termos')`, query `['colaborador-termos', colaborador.id, page, pageSize]`, `TermosLista mostrarItem` + `Pagination`; vazio: "Nenhum termo registrado para este colaborador."

- [ ] **Step 1: Testes** (MSW, dentro de `MemoryRouter`): lista dois termos com o link do equipamento apontando para `/items/12?aba=termos` e as pílulas certas; a requisição leva `limit=25`; lista vazia mostra a mensagem; `ColaboradorDetailPage.test.tsx` continua verde.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/colaborador-detail/tabs/termos.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(termos): aba Termos no detalhe do colaborador`.

### Task 12: Frontend — empréstimo com colaborador do cadastro

**Files:**
- Create: `frontend/src/lib/email.ts`, `frontend/src/components/termos/ColaboradorEmailField.tsx`, `frontend/src/components/termos/colaboradorEmail.ts`
- Modify: `frontend/src/api/loans.ts`, `frontend/src/pages/LoanPage.tsx`, `frontend/src/pages/LoanPage.test.tsx`, `frontend/src/components/equipment/ConfirmacaoTermo.tsx`

**Interfaces:**
- Consumes: `ColaboradorSelect`, `updateColaborador`, `Colaborador` (F0); `TermoSolicitacao` (Task 10).
- Produces:
  - `api/loans.ts`: `interface LoanPayload { item_id: number; colaborador_id?: number; usuario?: string; cpf?: string; center_cost: string; cargo?: string; setor?: string; revenda: string; date_issue: string; pessoa_juridica?: boolean }`; `initiateLoan(data: LoanPayload): Promise<{ detail: string; termo: TermoSolicitacao | null }>`; `initiateReturn(itemId: number, body?: { colaborador_id?: number }): Promise<{ detail: string; download_url: string; filename: string; termo: TermoSolicitacao | null }>` (sem `body` não envia corpo); `downloadReturnTerm` e o resto ficam iguais.
  - `lib/email.ts`: `isValidEmail(v: string): boolean` (`/^[^@\s]+@[^@\s]+\.[^@\s]+$/`, mesmo padrão do backend).
  - `colaboradorEmail.ts`: `garantirEmailDoColaborador(colaborador: Colaborador, email: string): Promise<Colaborador>` — vazio → `Error('Informe o e-mail do colaborador.')`; inválido → `Error('E-mail inválido.')`; igual ao atual (sem caixa) → devolve o colaborador; senão `updateColaborador(colaborador.id, { email: email.trim().toLowerCase() })`.
  - `ColaboradorEmailField({ colaborador, value, onChange, disabled })` — `Label` "E-mail do colaborador *" (`htmlFor="colaborador-email"`), `Input type="email"`, ajuda "O RevSign envia o link de assinatura para este e-mail (pode ser pessoal)."; não renderiza sem colaborador.
  - `LoanPage`: passo 2 (título mantido "2. Dados do Colaborador Responsável", descrição do passo no `STEPS`: "Cadastro e e-mail") troca nome/CPF digitados por `ColaboradorSelect` (rótulo "Colaborador *") + resumo do cadastro (NOME, CPF formatado, cargo, setor, UNIDADE) + `ColaboradorEmailField` + o toggle "É pessoa jurídica" (mantido). "Próximo" do passo 2: sem colaborador → toast "Selecione o colaborador."; `garantirEmailDoColaborador` (botão com `loading` enquanto salva; erro local mostra a mensagem do `Error`, erro da API usa `getErrorMessage`); depois preenche `revenda` com `unidade_nome` do colaborador quando ela está em `revendas`, e `cargo`/`setor` com os do cadastro. Passo 3: cargo e setor do cadastro aparecem só leitura com "Do cadastro do colaborador"; sem valor no cadastro, os campos atuais (setor obrigatório: "Selecione o setor."). Passo 4: linhas Colaborador, CPF, "E-mail para assinatura", Revenda/Setor (`upper`), e `StatusPill status="Pendente"` no lugar do badge "PENDENTE TERMO"; "Confirmar Empréstimo" com `loading`. Envio: `{ item_id, colaborador_id, center_cost, cargo, setor, revenda, date_issue, pessoa_juridica }` (sem `usuario`/`cpf`). Sucesso: `termo?.status === 'enviado'` → toast `Empréstimo iniciado. Termo enviado para assinatura em ${termo.destinatario_email}.` (painel de confirmação não abre); senão abre o painel `ConfirmacaoTermo` e toast atual. Tabela "Empréstimos Pendentes de Termo": `Pagination` + `usePageSize('emprestimo-pendentes')` (página 1-based, correção da última página), `upper(item.revenda)`. Barra lateral: o item "O CPF é validado…" vira "O colaborador vem do cadastro; o e-mail é obrigatório (pode ser pessoal) e recebe o link de assinatura."
  - `ConfirmacaoTermo`: botão "Confirmar Empréstimo" com `loading={confirmMutation.isPending}` (sem o texto alternativo "Confirmando...").

- [ ] **Step 1: Testes** (`LoanPage.test.tsx`; mocks de `/api/items`, `/api/colaboradores`, `/api/colaboradores/:id` (GET/PUT), `/api/loans`, `/api/constants` quando preciso; helper `escolherColaborador(user, nome)` = clicar em `/Selecione um colaborador/`, digitar no placeholder `/Buscar por nome, CPF ou matrícula/`, clicar no nome):
  - mantido: `não avança da etapa Equipamento sem selecionar um item disponível`.
  - `não avança da etapa Colaborador sem escolher alguém do cadastro` → "Selecione o colaborador." (MUDANÇA INTENCIONAL F3: substitui o teste "CPF inválido barra o avanço" — CPF agora vem do cadastro, já validado pela F0; comentário no teste dizendo isso).
  - `exige e-mail válido e grava o e-mail no cadastro antes de seguir` (colaborador id 8 sem e-mail): digita `joao@` → "E-mail inválido."; corrige para `joao.silva@gmail.com` → PUT `/api/colaboradores/8` com `{ email: 'joao.silva@gmail.com' }` → aparece "3. Condições Operacionais e Alocação".
  - adaptado: `preenche a revenda a partir do item e barra sem setor quando o cadastro não tem setor` → "Selecione o setor." e nenhuma "Selecione a revenda.".
  - `usa setor e cargo do cadastro e envia colaborador_id` (colaborador 7 com setor "Financeiro", cargo "Analista", e-mail "maria@gmail.com"): passo 3 mostra "Do cadastro do colaborador"; confirmar → corpo do POST contém `{ item_id: 42, colaborador_id: 7, setor: 'Financeiro', cargo: 'Analista', revenda: 'Matriz' }` e não tem `usuario` nem `cpf`; resposta `termo.status: 'enviado'` → toast "Empréstimo iniciado. Termo enviado para assinatura em maria@gmail.com.".
  - `com assinatura manual abre o painel de confirmação` (`termo.status: 'manual'`) → heading "Confirmar Empréstimo — Item #42".
  - `pendentes de termo usam a paginação padrão` → a busca `status=Pendente` leva `limit=25`; "Mostrando 1–1 de 1".
  - adaptados (helpers novos): voltar etapa/não pular; toggle PJ desmarcado.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/LoanPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(emprestimo): colaborador do cadastro com e-mail obrigatório no empréstimo`.

### Task 13: Frontend — devolução com colaborador do cadastro

**Files:**
- Modify: `frontend/src/pages/ReturnPage.tsx`
- Create: `frontend/src/pages/ReturnPage.test.tsx`

**Interfaces:**
- Consumes: `initiateReturn` (Task 12), `ColaboradorSelect`, `getColaborador`, `ColaboradorEmailField`, `garantirEmailDoColaborador`, `downloadAuthenticated`.
- Produces: "Iniciar Devolução" abre o diálogo "Iniciar devolução" (equipamento e responsável registrado); item com `colaborador_id` → `ColaboradorSelect` desabilitado com esse colaborador (dados via `getColaborador`); sem vínculo → `ColaboradorSelect` com `somenteAtivos={false}` e o aviso "Equipamento sem colaborador do cadastro: escolha quem está devolvendo (responsável registrado: <assigned_to>)."; `ColaboradorEmailField`; botão "Gerar termo e iniciar" (`loading`): sem colaborador → "Selecione o colaborador."; `garantirEmailDoColaborador`; `initiateReturn(item.id, { colaborador_id })`. Resposta `termo.status === 'enviado'` → toast `Termo de devolução enviado para assinatura em ${email}.` sem download; senão `downloadAuthenticated(download_url, filename)`, abre a "Conferência Física & Termo de Devolução" atual e o toast atual. As duas tabelas: `Pagination` com `usePageSize('devolucao-pendentes')` e `usePageSize('devolucao-em-campo')` (página 1-based; a busca volta para a página 1); `upper(revenda)`; o badge "RECOLHIMENTO" vira `StatusPill status="Pendente Devolução"`; "Finalizar Devolução" com `loading`; "Iniciar Devolução" com `loading` só na linha em andamento.

- [ ] **Step 1: Testes** (MSW; `URL.createObjectURL = vi.fn(() => 'blob:x')` e `URL.revokeObjectURL = vi.fn()` no `beforeEach`):
  - `item sem vínculo exige escolher do cadastro e mostra desligados` → "Gerar termo e iniciar" sem escolha → "Selecione o colaborador."; a busca em `/api/colaboradores` não tem `situacao` (Review Focus: desligado pode devolver).
  - `item vinculado envia colaborador_id e, com RevSign, não baixa o termo` → corpo do POST `/api/loans/9/return/initiate` igual a `{ colaborador_id: 7 }`; toast "Termo de devolução enviado para assinatura em maria@gmail.com."; nenhuma requisição a `/api/documents/files/...`.
  - `assinatura manual baixa o termo e abre a conferência física` → GET `/api/documents/files/termos_devolucao/termo.docx` feito; diálogo "Conferência Física & Termo de Devolução" visível.
  - `tabelas usam a paginação padrão` → `limit=25` e "Mostrando 1–1 de 1".
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/ReturnPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(devolucao): colaborador do cadastro e termo de devolução pelo RevSign`.

### Task 14: Frontend — tela de Termos

**Files:**
- Modify: `frontend/src/pages/TermsPage.tsx`
- Create: `frontend/src/pages/TermsPage.test.tsx`

**Interfaces:**
- Consumes: `getTermosDoItem`, `TermoVigenteCard` (Task 10).
- Produces: as duas abas usam `Pagination` + `usePageSize('termos')` (página 1-based por aba, volta a 1 quando a busca ou o tamanho mudam, corrige a última página), no lugar do paginador próprio de 10; `upper(item.revenda)` na lista e no visualizador; no visualizador, abaixo do cabeçalho do documento, seção "Assinatura eletrônica" com `TermoVigenteCard` do item selecionado (query `['item-termos', id, 'vigente']` = `getTermosDoItem(id, { limit: 1, offset: 0 })`).

- [ ] **Step 1: Testes** (MSW): `mostra a assinatura eletrônica do termo selecionado` (item Pendente id 5; vigente `manual` com motivo) → "Assinatura eletrônica", pílula "Assinatura manual" e o motivo; `paginação padrão de 25` → a busca `status=Pendente` leva `limit=25` e aparece "Mostrando 1–1 de 1".
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/TermsPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run`, `npx tsc --noEmit` e `npm run build` → verdes.
- [ ] **Step 5: Commit** `feat(termos): paginação padrão e assinatura eletrônica na tela de Termos`.

### Task 15: Verificação final da fase

**Files:** nenhum novo (só correções que a verificação apontar, cada uma com teste).

- [ ] **Step 1: Backend inteiro:** `RBT -q -m "not mudanca_esperada"` → todos passam (linha de base da F0 + novos da F3); `RBT -q tests/test_migrations.py` → ida e volta da `0300` ok.
- [ ] **Step 2: Frontend inteiro:** em `frontend/`, `npx vitest run`, `npx tsc --noEmit`, `npm run build` → verdes.
- [ ] **Step 3: Conferir o spec F3 item a item** (anotar no relatório da fase o teste que prova cada um):
  1. Colaborador do cadastro no empréstimo e na devolução, e-mail obrigatório e validado, campos antigos gravados → Tasks 6, 12, 13.
  2. `termo_solicitacoes` e termo vigente (tipo, versão do modelo, status, data da assinatura, quem solicitou) → Tasks 2, 7, 10.
  3. Envio pelo RevSign com `REVSIGN_ENABLED=true`; `manual` com integração desligada ou falha; fluxo atual intacto → Tasks 4, 5, 6.
  4. Webhook autenticado pelo segredo + consulta periódica; ao assinar, PDF em `termos_assinados`/`termos_devolucao_assinados` e a mesma confirmação do upload manual → Tasks 5, 8, 9.
  5. "Solicitar assinatura" a qualquer momento, com agendamento; nova cancela a pendente → Tasks 4, 7, 9, 10.
  6. Disparo automático em todo empréstimo e devolução → Task 6.
  7. Histórico por equipamento e por colaborador → Tasks 7, 10, 11.
  8. `RevSignClient` isolado, com falso e HTTP, premissas do contrato na docstring → Task 1.
  - Aceite: empréstimo gera solicitação e o webhook simulado conclui e grava o PDF (`test_webhook_assinado_conclui_emprestimo_e_grava_pdf`); envio agendado no tique certo (`test_envio_agendado_sai_no_tique_certo`); troca cancela a pendente (`test_devolucao_cancela_uso_pendente`, `test_nova_solicitacao_cancela_pendente_anterior_tambem_no_revsign`); integração desligada com os testes atuais verdes (Step 1).
  - Matriz: 3.1, 3.4 (exigência no envio), 3.5, 6.1, 6.2, 6.4, 6.5, 6.6.
  - Review Focus: os cinco testes citados existem e passam.
- [ ] **Step 4: Relatório da fase** (no PR/mensagem ao coordenador, sem arquivo novo): dependência `httpx==0.28.1` e justificativa; linhas acrescentadas em arquivos compartilhados (`main.py`, `dependencies.py`, `jobs.py`, `.env.example`, `requirements.txt`, `routers/history.py`); a troca de `toEqual([...])` por `expect.arrayContaining([...])` nos testes de abas da F0; premissas do contrato do RevSign pendentes de confirmação.
- [ ] **Step 5: Commit** (se houve correção) `test(termos): verificação final da fase 3`.
