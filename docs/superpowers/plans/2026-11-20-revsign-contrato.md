# RevSign — contrato da integração — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Com `REVSIGN_ENABLED=true`, o Patrimônio envia o termo (PDF com o marcador `{{assinatura}}`) pela API `revalle/patrimonio/v1` do RevSign e recebe o retorno pelo webhook HMAC, sem mudar o fluxo de termos.

**Architecture:** Reescrever `HttpRevSignClient` para o contrato, mantendo a interface `RevSignClient`; converter o `.docx` gerado para PDF com LibreOffice headless; pôr o marcador nos modelos; `termos.enviar` passa `tipo` e CPF.

**Tech Stack:** FastAPI, httpx (`MockTransport` nos testes), python-docx, LibreOffice headless (`soffice`), pytest.

**Spec:** `docs/superpowers/specs/2026-11-20-revsign-contrato-design.md` · Contrato: `docs/integracoes/revsign-contrato-patrimonio.md` (seções 2–8 são a referência exata de campos, rotas, estados, erros e webhook).

## Global Constraints
- A interface pública de `app/services/revsign_client.py` continua: `RevSignClient` (`criar_documento`, `consultar`, `cancelar`, `reenviar`, `baixar_assinado`, `validar_webhook`), `NovoDocumento`, `DocumentoCriado`, `StatusDocumento`, `EventoWebhook`, `RevSignError`, `WebhookInvalido`, `WebhookMalformado`, `FakeRevSignClient`, `usar_cliente`, `get_revsign_client`.
- `REVSIGN_BASE_URL` é a base completa (ex.: `http://agente_revalle_revsign:3000/api/revalle/patrimonio/v1`); o cliente só acrescenta `/termos…`.
- Termo `uso` vai com `kind=ENTREGA`; `entrega` → `ENTREGA`; `devolucao` → `DEVOLUCAO`.
- `message` do envio: `Revalle · Patrimônio: assine o termo do equipamento sob sua responsabilidade.`
- Nenhuma mensagem de erro ou log contém o token nem o segredo.
- Mensagens ao usuário em português.
- Teste de backend: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh C:/Users/enzo.jz/dev/patrimonio/revsign revsign <args pytest>` (MySQL de teste e imagem `stock-system-backend-test:dev` já no ar; a imagem ganha LibreOffice e poppler na Task 2). Não criar nem remover containers, imagens ou volumes, exceto o rebuild da imagem de teste descrito na Task 2.
- Commits terminam com `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus
1. Webhook com timestamp velho (> 300 s) ou do futuro, ou assinatura sem o prefixo `sha256=`: 401 — teste na Task 1.
2. `409 DUPLICATE_EXTERNAL_ID` depois de uma falha de rede na criação: tratar como sucesso com o `termoId` devolvido, sem criar outro — teste na Task 1.
3. LibreOffice travado ou ausente: o termo cai em `manual` com motivo claro em até 60 s, sem derrubar a requisição — teste na Task 2 e Task 4.
4. Colaborador sem CPF válido ou fora da base do RH: `manual` com o motivo da spec — teste na Task 4.
5. Termo impresso a partir do `.docx` manual: o marcador não aparece (fonte branca) e não desloca o layout — conferência visual na Task 3.

---

### Task 1: Cliente HTTP e webhook no contrato

**Files:**
- Modify: `backend/app/services/revsign_client.py`
- Test: `backend/tests/test_revsign_client.py` (reescrever os testes do HTTP e do webhook), `backend/tests/test_revsign_webhook.py` (só o jeito de assinar as requisições), qualquer teste que construa `NovoDocumento`

**Interfaces:**
- Produces:
  - `NovoDocumento` ganha `tipo: Literal["ENTREGA", "DEVOLUCAO"]` e `signatario_cpf: str` (só dígitos); `arquivo` passa a ser PDF e `nome_arquivo` termina em `.pdf`.
  - `RevSignError(mensagem: str, codigo: str | None = None)` — `codigo` é o `error.code` do contrato quando houver; `str(e)` é uma mensagem em português pronta para o técnico.
  - `assinar_webhook(segredo: str, corpo: bytes, timestamp: int) -> dict[str, str]` (módulo-level, usado pelos testes e pelo fake): devolve `{"X-RevSign-Timestamp": str(ts), "X-RevSign-Signature": "sha256=" + hex}`.
  - `EventoWebhook(tipo=event, documento_id=termoId, referencia_externa=externalId)`.
- Mapas fixos:
  - Status: `PENDING`→pendente, `SIGNED`→assinado (`assinado_em` de `signer.signedAt` ou `completedAt`), `REJECTED`→recusado (`motivo` = `signer.rejectionReason`), `EXPIRED`→expirado, `CANCELLED`→cancelado.
  - Erro → mensagem: `UNAUTHORIZED` "Chave da integração com o RevSign recusada."; `NOT_CONFIGURED` "Integração desligada no RevSign."; `INVALID_CPF` "CPF do colaborador recusado pelo RevSign."; `COLABORADOR_NOT_FOUND` "Colaborador não está na base do RH do RevSign."; `HR_BASE_UNAVAILABLE` "Base do RH do RevSign fora do ar."; `MARKER_NOT_FOUND` "O termo não tem o marcador de assinatura."; `INVALID_FILE` "O RevSign recusou o PDF do termo."; `FILE_TOO_LARGE` "PDF do termo maior que 10 MB."; `RESEND_TOO_SOON` "Aguarde 10 minutos entre reenvios."; `NOT_PENDING` "O termo já não está pendente no RevSign."; `RATE_LIMITED` "RevSign ocupado; tente de novo em instantes."; outros/sem corpo: "RevSign respondeu <status>.".

- [ ] **Step 1: Testes que falham** (`MockTransport`, sem rede):
  - `test_criar_envia_multipart_do_contrato` → `POST {BASE}/termos`, Bearer, campos `externalId`, `kind`, `title`, `signerName`, `signerEmail`, `signerCpf`, `message` e `file` (`application/pdf`); 201 `{"termoId": "envelope_x"}` → `DocumentoCriado("envelope_x")`.
  - `test_criar_duplicado_reaproveita_termo` → 409 `{"error":{"code":"DUPLICATE_EXTERNAL_ID","termoId":"envelope_y"}}` → `DocumentoCriado("envelope_y")`.
  - `test_criar_erro_vira_mensagem` (parametrizado com 3 códigos do mapa) → `RevSignError` com a mensagem e o `codigo`.
  - `test_consultar_mapeia_status` (5 estados).
  - `test_cancelar_envia_motivo_e_aceita_404`; `test_cancelar_nao_pendente_levanta` (409 `NOT_PENDING`).
  - `test_reenviar_429_mensagem` e `test_reenviar_ok` (`POST /termos/{id}/reenviar`).
  - `test_baixar_pdf` (`GET /termos/{id}/pdf`), mantendo os testes atuais de PDF inválido e grande.
  - Mantidos: erro de rede sem token na mensagem; sem configuração levanta.
  - `test_validar_webhook` parametrizado: assinatura certa → evento; segredo errado, sem cabeçalhos, sem `sha256=`, timestamp 301 s atrás e 301 s à frente → `WebhookInvalido`; corpo não JSON ou sem `event` → `WebhookMalformado`; cabeçalhos em minúsculas aceitos.
- [ ] **Step 2:** Rodar `… tests/test_revsign_client.py -q` → FAIL. **Step 3:** implementar; o validador compartilhado (`_ValidadorWebhook`) passa a conferir o HMAC com `hmac.compare_digest` e a janela de 300 s. **Step 4:** atualizar `test_revsign_webhook.py` para assinar com `assinar_webhook` e corpo no formato do contrato; atualizar quem cria `NovoDocumento` nos testes; rodar `tests/test_revsign_client.py tests/test_revsign_webhook.py tests/test_termos_servico.py tests/test_loan_colaborador.py -q` → PASS.
- [ ] **Step 5:** Commit `feat(revsign): cliente e webhook no contrato da integração do Patrimônio`.

### Task 2: Conversão do termo para PDF

**Files:**
- Create: `backend/app/services/termo_pdf.py`
- Modify: `backend/Dockerfile` (`apt-get install -y --no-install-recommends libreoffice-writer-nogui fonts-dejavu fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea && rm -rf /var/lib/apt/lists/*`, antes do `pip install`), `C:/Users/enzo.jz/dev/patrimonio/_tools/img/Dockerfile` (mesmos pacotes + `poppler-utils`; reconstruir com `docker build -t stock-system-backend-test:dev C:/Users/enzo.jz/dev/patrimonio/_tools/img` — mesma tag, sem criar outra)
- Test: `backend/tests/test_termo_pdf.py`

**Interfaces:**
- Produces: `docx_para_pdf(conteudo: bytes, *, timeout: int = 60) -> bytes`; `class ConversaoPdfFalhou(Exception)` com mensagem em português ("LibreOffice não está instalado no servidor.", "A conversão do termo para PDF demorou demais.", "Não foi possível converter o termo para PDF.").
- Implementação: diretório temporário por chamada (`tempfile.TemporaryDirectory`), `soffice --headless --norestore -env:UserInstallation=file:///<tmp>/perfil --convert-to pdf --outdir <tmp> <tmp>/termo.docx` via `subprocess.run(..., timeout=timeout, capture_output=True)`; binário por `shutil.which("soffice") or shutil.which("libreoffice")`; resultado precisa começar com `%PDF`.

- [ ] **Step 1: Testes que falham:**
  - `test_converte_modelo_real` (marcado `skipif(shutil.which("soffice") is None)`) → converte `modelos/termo.docx` e o resultado começa com `%PDF`.
  - `test_sem_libreoffice` → `monkeypatch` de `shutil.which` devolvendo `None` → `ConversaoPdfFalhou` "LibreOffice não está instalado no servidor."
  - `test_timeout` → `monkeypatch` de `subprocess.run` levantando `TimeoutExpired` → "A conversão do termo para PDF demorou demais."
  - `test_saida_que_nao_e_pdf` → fake `subprocess.run` que grava lixo no `outdir` → "Não foi possível converter o termo para PDF."
- [ ] **Step 2:** Reconstruir a imagem de teste (comando acima) e rodar `… tests/test_termo_pdf.py -q` → FAIL; implementar; PASS (o teste real **não** pode ficar pulado na imagem reconstruída).
- [ ] **Step 3:** `docker build -t stock-system-backend:verificacao backend` para provar que o Dockerfile de produção constrói, e `docker run --rm stock-system-backend:verificacao soffice --version`; depois `docker rmi stock-system-backend:verificacao`.
- [ ] **Step 4:** Commit `feat(termos): converte o termo para PDF com LibreOffice`.

### Task 3: Marcador `{{assinatura}}` nos modelos

**Files:**
- Modify: os 15 `.docx` de `backend/modelos/` (incluindo `modelos/devolucao/`)
- Create: `backend/scripts/marcar_assinatura.py` (o script que fez a alteração, para refazer se um modelo for trocado; idempotente)
- Possibly modify: `backend/app/db/inventory_manager_db.py` (só se o preenchimento atual apagar `{{…}}` desconhecidos)
- Test: `backend/tests/test_modelos_assinatura.py`

Regras:
- Cada modelo tem **exatamente um** `{{assinatura}}`, num run com cor branca (`RGBColor(0xFF,0xFF,0xFF)`), tamanho 8 pt, no lugar da assinatura do colaborador: devolução, na linha "Assinatura do Colaborador:"; `termo.docx`, acima de "Assinatura do empregado"; modelos por unidade, acima de "Assinatura" (o bloco com cidade e `{{data_hoje}}`). O campo do RevSign ocupa 180 × 54 pt a partir do canto superior esquerdo do marcador: a área não pode cobrir texto do termo (pode cobrir a linha de assinatura em branco).
- O texto visível do termo não muda.

- [ ] **Step 1: Testes que falham:**
  - `test_cada_modelo_tem_um_marcador_branco` (parametrizado nos 15 arquivos) → exatamente um run com `{{assinatura}}`, cor branca.
  - `test_termo_preenchido_mantem_o_marcador` → `render_loan_term` e `render_return_term` de um item de teste (fixtures do conftest) produzem `.docx` que ainda contém `{{assinatura}}`.
  - `test_pdf_tem_o_marcador` (`skipif` sem `soffice`/`pdftotext`) → `pdftotext` do PDF convertido de cada modelo contém `{{assinatura}}`.
- [ ] **Step 2:** FAIL; escrever `scripts/marcar_assinatura.py` e rodá-lo nos modelos; PASS.
- [ ] **Step 3: Conferência visual (obrigatória):** para cada modelo, converter para PDF, `pdftoppm -png -r 60` da página do marcador, desenhar com Pillow (já instalado? se não, só `pdftoppm`) um retângulo de 180 × 54 pt na posição do marcador (`pdftotext -bbox` dá a posição) e olhar as imagens com o Read tool. Registrar no relatório, por modelo, "área livre" ou o ajuste feito. Apagar as imagens.
- [ ] **Step 4:** Commit `feat(termos): marcador de assinatura do RevSign nos modelos`.

### Task 4: Envio com PDF, tipo e CPF

**Files:**
- Modify: `backend/app/services/termos.py` (`_gerar_documento`/`enviar`/`reenviar`), `backend/README.md` ou `README.md` (tabela de variáveis: novos valores de `REVSIGN_BASE_URL`/`REVSIGN_API_TOKEN`/`REVSIGN_WEBHOOK_SECRET`, URL interna do webhook, LibreOffice), `backend/.env.example`
- Test: `backend/tests/test_termos_servico.py` (casos novos), `backend/tests/test_revsign_webhook.py` se precisar

**Interfaces:**
- Consumes: Task 1 (`NovoDocumento.tipo`, `.signatario_cpf`, `RevSignError.codigo`), Task 2 (`docx_para_pdf`, `ConversaoPdfFalhou`).
- O `.docx` continua guardado em `documento_key` (download manual); só o que vai ao RevSign é o PDF, com `nome_arquivo` trocando `.docx` por `.pdf`.
- Motivos novos (gravados com `_cortar`): sem CPF válido (`cpf_valido` de `app/schemas/validators.py` ou o validador que o projeto já usa) → "Colaborador sem CPF válido: assinatura manual."; `ConversaoPdfFalhou` → f"{mensagem} Assinatura manual."; `RevSignError` → f"Falha ao enviar ao RevSign: {e} Assinatura manual." (sem duplicar pontuação).

- [ ] **Step 1: Testes que falham** (com `revsign_fake` existente):
  - `test_envio_usa_pdf_tipo_e_cpf` → empréstimo de colaborador com CPF → `revsign_fake.criados[0]` tem `tipo="ENTREGA"`, `signatario_cpf` só dígitos, `nome_arquivo` termina em `.pdf`, `arquivo` começa com `%PDF` (monkeypatch de `docx_para_pdf` devolvendo `b"%PDF-1.4 teste"` para não depender do LibreOffice).
  - `test_termo_de_uso_vai_como_entrega` e `test_devolucao_vai_como_devolucao`.
  - `test_sem_cpf_fica_manual`; `test_conversao_falhou_fica_manual` (fake levanta `ConversaoPdfFalhou("A conversão do termo para PDF demorou demais.")`) → motivo começa com essa frase; `test_colaborador_fora_da_base_fica_manual` (fake com `falhar_em` levantando `RevSignError("Colaborador não está na base do RH do RevSign.", "COLABORADOR_NOT_FOUND")`).
  - `test_reenviar_cedo_demais_mensagem` → `TermoErro` "Não foi possível reenviar: Aguarde 10 minutos entre reenvios."
- [ ] **Step 2:** FAIL; implementar; PASS. **Step 3:** suíte completa `… revsign -q -m "not mudanca_esperada"` → verde.
- [ ] **Step 4:** Commit `feat(termos): envio ao RevSign com PDF, tipo e CPF do colaborador`.
