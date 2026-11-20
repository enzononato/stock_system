# Coletor de dados da máquina — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** O técnico abre `/register/coletor` na máquina, cola um comando no Win+R e o cadastro do equipamento aparece preenchido com os dados do Windows, para conferir e salvar.

**Architecture:** Uma coleta (`coletas_maquina`) nasce com um código de uso único. Rotas públicas, autorizadas só pelo código, servem o script PowerShell e recebem o JSON bruto. O backend normaliza para os campos do formulário e procura correspondências. A página do técnico consulta o andamento e abre `/register` ou `/edit/:id` já preenchido. Ao salvar, sai da conta.

**Tech Stack:** FastAPI + PyMySQL (`get_cursor`), slowapi, migrações SQL aditivas; PowerShell 5.1; React 18 + TS + TanStack Query + Vitest/MSW.

**Spec:** `docs/superpowers/specs/2026-11-19-coletor-maquina-design.md`

**Nota:** o spec pedia o teste numa máquina real da empresa como primeira tarefa. O responsável preferiu construir a tela já ("pode fazer a tela sem problema, depois a gente altera"); esse teste passa a ser o primeiro passo depois do deploy, e a Task 3 faz a verificação possível aqui (o script rodando no PowerShell deste PC, sem envio).

## Global Constraints
- Mensagens ao usuário em português; nada de inglês na tela.
- Migração `1000_coletas_maquina.sql` + `.down.sql`, aditiva, sem charset/collation fixos, idempotente (`CREATE TABLE IF NOT EXISTS` / `DROP TABLE IF EXISTS`).
- Código: `secrets.token_urlsafe(16)`; banco guarda só `hashlib.sha256(codigo).hexdigest()`; validade **15 minutos**; **um** envio de dados por código.
- Comando exato (≤ 259 caracteres com `https://patrimonio.revalle.com.br`):
  `powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12';irm '<APP_PUBLIC_URL>/api/coletor/<codigo>'|iex"`
- Rotas públicas com `@limiter.limit(settings.COLETOR_RATE_LIMIT)`, padrão `"30/minute"`; corpo do envio até 64 KB.
- Rotas do técnico: `gestor_or_tecnico` (`app/dependencies.py`). Andamento só para quem criou ou Gestor (senão 404).
- Nada é gravado em `items` pela coleta; só o salvamento normal do formulário grava.
- Formatos: CPU `i5-1345U`, RAM `16GB`, armazenamento `512 SSD` (vários: `512 SSD + 1000 HDD`), sistema `Windows 11 Pro`, domínio `Sim`/`Não`, MAC `AA:BB:CC:DD:EE:FF`, AnyDesk só dígitos, licença = Product ID.
- Tamanho máximo de cada campo normalizado = coluna de destino: brand/model/identificador/host/cpu/sistema/licenca 100, ram/storage/anydesk/dominio 50, endereco_fisico 150.
- Teste de backend: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh C:/Users/enzo.jz/dev/patrimonio/coletor coletor <args pytest>` (MySQL de teste `stock_sys_mysql_test` em 127.0.0.1:3307 e imagem `stock-system-backend-test:dev` já estão no ar; não criar nem remover containers, imagens ou volumes). Se uma dependência nova for adicionada a `requirements*.txt`, avisar no relatório da task em vez de reconstruir a imagem.. Frontend: `cd frontend && npx vitest run <arquivo>` e `npx tsc --noEmit`.
- Commits terminam com `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus
1. Código vencido ou já usado chamado de novo pelo técnico (rodou o comando duas vezes): o script deve dizer "Este código venceu…" sem erro vermelho, e o envio repetido deve dar 410 sem sobrescrever os dados — teste na Task 4.
2. Máquina virtual ou montada (serial genérico "To be filled by O.E.M.", sem gateway, sem AnyDesk, fabricante "System manufacturer"): campos vazios, nada quebra, MAC cai no Ethernet físico — teste na Task 2.
3. Correspondência com equipamento baixado: não pode oferecer "Atualizar este cadastro" — teste na Task 6.
4. Corpo malicioso (strings enormes, HTML, tipos errados, JSON acima de 64 KB): 422/413, nada gravado; strings normalizadas cortadas no tamanho da coluna — teste na Task 4 e Task 2.
5. Logout que falha depois de salvar (rede caiu): a tela final aparece mesmo assim, com o aviso "Não foi possível sair da conta automaticamente. Saia pelo menu antes de entregar a máquina." — teste na Task 7.

---

### Task 1: Tabela e camada de dados das coletas

**Files:**
- Create: `backend/migrations/1000_coletas_maquina.sql`, `backend/migrations/1000_coletas_maquina.down.sql`
- Create: `backend/app/db/coleta_db.py`
- Test: `backend/tests/test_coleta_db.py`

**Interfaces:**
- Produces (`app/db/coleta_db.py`, class `ColetaDBManager`, todos os horários em UTC naive como o resto do banco):
  - `criar(codigo_hash: str, criado_por_id: int, expira_em: datetime) -> int`
  - `por_hash(codigo_hash: str) -> dict | None` (linha completa)
  - `por_id(coleta_id: int) -> dict | None` (com `dados` já desserializado de JSON, ou `None`)
  - `registrar_dados(codigo_hash: str, dados: dict, ip: str) -> bool` — um único `UPDATE … WHERE codigo_hash=%s AND recebido_em IS NULL AND expira_em > UTC_TIMESTAMP()`; devolve se uma linha mudou.
  - `apagar_antigas(dias: int = 7) -> int`
  - `correspondencias(serial: str | None, macs: list[str], anydesk: str | None) -> list[dict]` — itens com `UPPER(TRIM(identificador))` igual, ou `endereco_fisico`/`mac` normalizados (sem `:`/`-`, maiúsculo) em `macs`, ou `REGEXP_REPLACE(anydesk,'[^0-9]','')` igual; cada dict: `id, tipo, brand, model, status, revenda, ativo (bool de is_active), motivos (list[str] em {"serial","mac","anydesk"})`; ordenado por ativo desc, id desc. Entradas vazias não casam com campos vazios.

- [ ] **Step 1:** Migração: `coletas_maquina (id INT AUTO_INCREMENT PK, codigo_hash CHAR(64) NOT NULL UNIQUE, criado_por_id INT NOT NULL, criado_em DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, expira_em DATETIME NOT NULL, recebido_em DATETIME NULL, recebido_ip VARCHAR(45) NULL, dados JSON NULL, INDEX idx_coletas_criado_em (criado_em), CONSTRAINT fk_coletas_usuario FOREIGN KEY (criado_por_id) REFERENCES usuarios(id) ON DELETE CASCADE)`. Down: `DROP TABLE IF EXISTS coletas_maquina`.
- [ ] **Step 2: Testes que falham** em `test_coleta_db.py`:
  - `test_migracao_1000_ida_e_volta` → `ida_e_volta_migracao(1000)` (de `tests/fixtures_auth.py`) e a tabela existe depois.
  - `test_registrar_dados_uma_vez_so` → primeiro `registrar_dados` True; segundo False; `por_id` mantém os dados do primeiro.
  - `test_registrar_dados_vencido` → `expira_em` no passado → False.
  - `test_correspondencias_por_serial_mac_anydesk` → itens criados com `criar_item` (conftest) com `identificador="ABC123 "`, `endereco_fisico="aa-bb-cc-dd-ee-ff"`, `anydesk="1 234 567 890"`; consultas com `"abc123"`, `["AABBCCDDEEFF"]`, `"1234567890"` acham o item com os motivos certos; um único item que casa por dois critérios aparece uma vez com dois motivos.
  - `test_correspondencias_ignora_vazios` → item com `identificador=""` não casa com `serial=None`/`""`.
  - `test_correspondencias_item_baixado` → item com `is_active=0` (via `executar_sql`) vem com `ativo=False`.
  - `test_apagar_antigas` → coleta com `criado_em` de 8 dias atrás some; a de hoje fica.
- [ ] **Step 3:** Rodar: `… run-backend-tests.sh … coletor tests/test_coleta_db.py -q` → FAIL.
- [ ] **Step 4:** Implementar `ColetaDBManager` com `get_cursor` (padrão de `app/db/email_confirmacao_db.py`).
- [ ] **Step 5:** Rodar de novo → PASS. Commit `feat(coletor): tabela e camada de dados das coletas`.

### Task 2: Normalização da coleta bruta

**Files:**
- Create: `backend/app/schemas/coletor.py`, `backend/app/services/coletor_normalizacao.py`
- Test: `backend/tests/test_coletor_normalizacao.py` (unitário, sem banco)

**Interfaces:**
- Produces (`app/schemas/coletor.py`):
  - `Disco(BaseModel)`: `tipo: str | None` (≤20), `bytes: int | None` (≥0), `barramento: str | None` (≤20)
  - `Adaptador(BaseModel)`: `nome: str | None` (≤200), `descricao: str | None` (≤200), `mac: str | None` (≤40), `tipo: Literal["ethernet","wifi","outro"] = "outro"`, `ativo: bool = False`, `gateway: bool = False`
  - `ColetaBruta(BaseModel, extra="ignore")`: `versao_coletor: str` (≤20), `fabricante, modelo, versao_produto, serial, host, dominio, cpu, sistema, product_id, anydesk: str | None` (cada ≤200), `chassi: list[int] = []` (≤10 itens), `no_dominio: bool | None`, `ram_bytes: int | None`, `discos: list[Disco] = []` (≤10), `adaptadores: list[Adaptador] = []` (≤20)
  - `CamposColeta = dict[str, str]` com as chaves `brand, model, identificador, dominio, host, endereco_fisico, cpu, ram, storage, sistema, licenca, anydesk` (só as que tiverem valor)
- Produces (`app/services/coletor_normalizacao.py`):
  - `normalizar(bruta: ColetaBruta) -> ColetaNormalizada` (dataclass: `campos: dict[str, str]`, `tipo_sugerido: Literal["Notebook","Desktop"] | None`, `macs: list[str]` — todos os MACs físicos normalizados `AABBCCDDEEFF`, para a busca)
  - Funções internas testáveis: `marca(fabricante) -> str`, `cpu_curta(nome) -> str`, `ram(bytes) -> str`, `armazenamento(discos) -> str`, `sistema(caption) -> str`, `serial(valor) -> str`, `mac_gravado(adaptadores) -> str`, `tipo_pelo_chassi(chassi) -> str | None`.

Regras não óbvias (o resto vem do spec, tabela "O que o coletor lê"):
- `marca`: tira sufixos `Inc.`, `Inc`, `Ltda`, `Ltda.`, `Tecnologia`, `Corporation`, `Corp.`, `Co., Ltd.`, `Computer`, vírgulas soltas; mapa fixo `{"DELL": "Dell", "LENOVO": "Lenovo", "HP": "HP", "HEWLETT-PACKARD": "HP", "ASUSTEK": "ASUS", "ACER": "Acer", "SAMSUNG ELECTRONICS": "Samsung", "POSITIVO": "Positivo"}` por prefixo maiúsculo; genéricos (`"System manufacturer"`, `"To be filled by O.E.M."`, `"Default string"`) → vazio; demais em `str.title()`.
- `model`: Lenovo usa `versao_produto` quando não vazio; os demais usam `modelo`; genéricos → vazio.
- `cpu_curta`: tira `(R)`, `(TM)`, `CPU`, `Processor`, `@ …GHz`, `with Radeon …`, `\d+th Gen`, espaços duplos; Intel `Core(TM) i5-1345U` → `i5-1345U`; `Core Ultra 5 125U` → `Ultra 5 125U`; AMD `Ryzen 5 5500U` → `Ryzen 5 5500U`; outros → nome limpo.
- `ram`: arredonda para o GB inteiro mais próximo (`17179869184` → `16GB`; `8463519744` → `8GB`).
- `armazenamento`: ignora `barramento` USB; GiB → tamanho comercial mais próximo em `[120, 128, 240, 250, 256, 480, 500, 512, 960, 1000, 2000]` (tolerância 12%); fora da lista, GB inteiro; tipo `SSD`/`HDD` (`Unspecified` com barramento NVMe → SSD; senão sem tipo); vários discos unidos por ` + `.
- `serial`: genéricos (`To be filled by O.E.M.`, `Default string`, `System Serial Number`, `0`, `None`, só zeros/espaços) → vazio; senão `strip()`.
- `mac_gravado`: adaptador `ativo and gateway`; senão primeiro `ethernet` ativo; senão primeiro `wifi` ativo; senão vazio. Formato `AA:BB:CC:DD:EE:FF`. MAC com tamanho inválido → ignorado.
- `tipo_pelo_chassi`: {8, 9, 10, 14, 30, 31, 32} → Notebook; {3, 4, 5, 6, 7, 13, 15, 16, 35, 36} → Desktop; senão None.
- `dominio`: `no_dominio` True → "Sim", False → "Não", None → ausente.
- Todo valor final passa por corte no tamanho da coluna (Global Constraints) e tem caracteres de controle e `<`/`>` removidos.

- [ ] **Step 1: Testes que falham**, cada um com valores reais:
  - `test_dell_latitude` (fabricante `"Dell Inc."`, modelo `"Latitude 3420"`, CPU `"11th Gen Intel(R) Core(TM) i5-1135G7 @ 2.40GHz"`, RAM `17179869184`, disco NVMe 512110190592 SSD, chassi `[10]`) → `brand="Dell"`, `model="Latitude 3420"`, `cpu="i5-1135G7"`, `ram="16GB"`, `storage="512 SSD"`, `tipo_sugerido="Notebook"`.
  - `test_lenovo_usa_nome_comercial` (`"LENOVO"`, modelo `"21JRS00X00"`, versao_produto `"ThinkPad E14 Gen 5"`) → `brand="Lenovo"`, `model="ThinkPad E14 Gen 5"`.
  - `test_daten_e_desktop` (`"Daten Tecnologia Ltda"`, `"ESTELAR"`, chassi `[3]`, discos `[SSD 240057409536, HDD 1000204886016]`) → `brand="Daten"`, `storage="240 SSD + 1000 HDD"`, `tipo_sugerido="Desktop"`.
  - `test_hp_core_ultra` (`"HP"`, CPU `"Intel(R) Core(TM) Ultra 5 125U"`) → `cpu="Ultra 5 125U"`.
  - `test_amd_ryzen` (`"AMD Ryzen 5 5500U with Radeon Graphics"`) → `"Ryzen 5 5500U"`.
  - `test_maquina_generica_sem_rede` (fabricante `"System manufacturer"`, serial `"To be filled by O.E.M."`, sem adaptadores ativos, chassi `[1]`) → sem `brand`, sem `identificador`, sem `endereco_fisico`, `tipo_sugerido is None`.
  - `test_mac_prefere_gateway_depois_ethernet` → wifi ativo com gateway vence ethernet ativo sem gateway; sem gateway, ethernet ativo vence wifi ativo; `macs` contém os dois normalizados.
  - `test_sistema_e_licenca` (`"Microsoft Windows 11 Pro"`, `"00355-62588-12345-AAOEM"`) → `sistema="Windows 11 Pro"`, `licenca="00355-62588-12345-AAOEM"`.
  - `test_anydesk_so_digitos` (`"1 234 567 890"`) → `"1234567890"`; `test_dominio_sim_nao`.
  - `test_corta_no_tamanho_da_coluna_e_remove_html` (modelo com 300 chars contendo `<script>`) → `len(model)==100` e sem `<`/`>`.
  - `test_schema_recusa_listas_enormes` → `ColetaBruta(adaptadores=[…21 itens])` levanta `ValidationError`.
- [ ] **Step 2:** Rodar `… tests/test_coletor_normalizacao.py -q` → FAIL.
- [ ] **Step 3:** Implementar schemas e normalização.
- [ ] **Step 4:** Rodar → PASS. Commit `feat(coletor): normalização dos dados coletados para o formulário`.

### Task 3: Script PowerShell e montagem do comando

**Files:**
- Create: `backend/app/templates/coletor/coletor.ps1`, `backend/app/services/coletor_script.py`
- Test: `backend/tests/test_coletor_script.py`

**Interfaces:**
- Produces (`app/services/coletor_script.py`):
  - `montar_comando(base_url: str, codigo: str) -> str` (texto exato do Global Constraints; `base_url` sem `/` final)
  - `montar_script(base_url: str, codigo: str) -> str` (lê o template e substitui `__URL_ENVIO__` por `f"{base_url}/api/coletor/{codigo}/dados"` e `__VERSAO__` por `VERSAO_COLETOR = "1"`; recusa `codigo` fora de `^[A-Za-z0-9_-]{16,64}$` com `ValueError`)
  - `script_vencido() -> str` (script que escreve "Este código venceu. Gere um novo na página do Patrimônio." e `Read-Host "Aperte Enter para fechar"`)
  - `montar_cmd(base_url: str, codigo: str) -> str` (`@echo off` + `montar_comando(...)`, com finais de linha CRLF)

O template (`coletor.ps1`, PowerShell 5.1, sem administrador, sem gravar arquivo):
- `$ErrorActionPreference='Stop'`; função `Ler { param($bloco) try { & $bloco } catch { $null } }` envolvendo cada leitura.
- Lê as fontes da tabela do spec (CIM: `Win32_ComputerSystem`, `Win32_ComputerSystemProduct`, `Win32_BIOS`, `Win32_SystemEnclosure`, `Win32_Processor`, `Win32_PhysicalMemory`, `Win32_OperatingSystem`; `Get-PhysicalDisk`; `Get-NetAdapter -Physical` + `Get-NetIPConfiguration` para `gateway`; AnyDesk em `C:\ProgramData\AnyDesk\system.conf` linha `ad.anynet.id=`, depois `& "${env:ProgramFiles(x86)}\AnyDesk\AnyDesk.exe" --get-id` e `"$env:ProgramFiles\AnyDesk\AnyDesk.exe"`), monta o objeto com as chaves de `ColetaBruta` (adaptador `tipo`: `wifi` se `PhysicalMediaType` contém `802.11` ou `NdisPhysicalMedium` = 9, `ethernet` se `802.3`, senão `outro`).
- Escreve "Coletando os dados desta máquina para o Patrimônio…"; se `$env:COLETOR_SO_MOSTRAR -eq '1'`, imprime o JSON (`ConvertTo-Json -Depth 4`) e para; senão `Invoke-RestMethod -Method Post -Uri __URL_ENVIO__ -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($json))`.
- Sucesso: "Pronto! Volte ao navegador." e `Start-Sleep 5`. Falha: mostra `$_.ErrorDetails.Message` (ou a mensagem da exceção) e `Read-Host "Aperte Enter para fechar"`.

- [ ] **Step 1: Testes que falham:**
  - `test_comando_cabe_no_win_r` → `len(montar_comando("https://patrimonio.revalle.com.br", "A"*22)) <= 259` e o texto é exatamente o do Global Constraints.
  - `test_script_tem_url_e_codigo` → contém `https://x/api/coletor/<codigo>/dados`, não contém `__URL_ENVIO__` nem `__VERSAO__`.
  - `test_script_recusa_codigo_estranho` → `"abc'; rm"` levanta `ValueError`.
  - `test_cmd_tem_crlf_e_comando`.
  - `test_template_cobre_todas_as_chaves` → cada nome de campo de `ColetaBruta` aparece no template.
- [ ] **Step 2:** Rodar `… tests/test_coletor_script.py -q` → FAIL. **Step 3:** implementar. **Step 4:** PASS.
- [ ] **Step 5: Verificação real no Windows (obrigatória):** renderizar com `python -c` um script com `base_url="http://127.0.0.1:1"` e código de teste, salvar no scratchpad, e rodar no **PowerShell tool** desta sessão (não via `powershell.exe` filho, que o ambiente bloqueia) com `$env:COLETOR_SO_MOSTRAR='1'; Get-Content <arquivo> -Raw | iex`. Esperado: JSON com `fabricante`, `serial`, `discos`, `adaptadores` preenchidos; validar esse JSON com `ColetaBruta.model_validate_json` e `normalizar` (no container de teste ou com o Python do host) e registrar a saída normalizada no relatório da task. Apagar o arquivo depois.
- [ ] **Step 6:** Commit `feat(coletor): script PowerShell do coletor e comando do Win+R`.

### Task 4: Rotas do coletor

**Files:**
- Create: `backend/app/routers/coletor.py`
- Modify: `backend/app/main.py` (import e `include_router`), `backend/app/core/config.py` (`COLETOR_RATE_LIMIT: str = "30/minute"`, `COLETOR_VALIDADE_MINUTOS: int = 15`), `backend/tests/conftest.py` (`os.environ["COLETOR_RATE_LIMIT"] = "10000/minute"` junto do `LOGIN_RATE_LIMIT`), `backend/app/jobs.py` (job `limpar_coletas`, `diariamente_as="03:30"`, chama `ColetaDBManager().apagar_antigas(7)`), `backend/.env.example` e `README.md` (as duas variáveis novas, uma linha cada na tabela existente).
- Test: `backend/tests/test_coletor_api.py`

**Interfaces:**
- Consumes: Task 1 (`ColetaDBManager`), Task 2 (`ColetaBruta`, `normalizar`), Task 3 (`montar_comando`, `montar_script`, `montar_cmd`, `script_vencido`).
- Produces (HTTP):
  - `POST /api/coletor` (gestor_or_tecnico) → 201 `{"id": int, "comando": str, "url_arquivo": "/api/coletor/<codigo>/coletor.cmd", "expira_em": ISO-8601 UTC com "Z"}`
  - `GET /api/coletor/{codigo}` (público) → 200 `text/plain; charset=utf-8`, script válido ou `script_vencido()` (código inexistente, vencido ou já recebido).
  - `GET /api/coletor/{codigo}/coletor.cmd` (público) → 200 `application/octet-stream`, `Content-Disposition: attachment; filename="coletor-patrimonio.cmd"`; vencido → 410 `{"detail": "Este código venceu. Gere um novo na página do Patrimônio."}`.
  - `POST /api/coletor/{codigo}/dados` (público) → 200 `{"ok": true}`; 410 (mesma mensagem) se vencido/já usado/inexistente; 413 `"Dados grandes demais."` acima de 65536 bytes (ler o corpo em pedaços como `integracoes_revsign.py`); 422 se não validar em `ColetaBruta`. Grava `bruta.model_dump()` e o IP de `get_client_ip`.
  - `GET /api/coletor/{id}/andamento` (gestor_or_tecnico; outro Técnico → 404 `"Coleta não encontrada."`) → `{"status": "aguardando"|"recebido"|"vencido", "expira_em": str, "campos"?: dict, "tipo_sugerido"?: str|null, "correspondencias"?: list}`; "recebido" normaliza na hora a partir de `dados` e chama `correspondencias(campos.get("identificador"), normalizada.macs, campos.get("anydesk"))`.
- Base da URL: `settings.APP_PUBLIC_URL.rstrip("/")`.

- [ ] **Step 1: Testes que falham** (`client`, `client_tecnico`, `client_gestor`, `client_aprendiz` do conftest):
  - `test_criar_coleta_tecnico` → 201, `comando` contém `/api/coletor/` e o código; o código **não** está no banco em texto (só o hash).
  - `test_aprendiz_nao_cria` → 403.
  - `test_script_publico_e_vencido` → GET com o código devolve script com `/dados`; com código inventado devolve o texto "Este código venceu".
  - `test_cmd_attachment`.
  - `test_envio_ok_e_andamento` → POST dados (payload Dell da Task 2) → 200; andamento → `recebido`, `campos["brand"]=="Dell"`, `tipo_sugerido=="Notebook"`.
  - `test_envio_repetido_410_mantem_primeiro` → segundo POST com outro fabricante → 410; andamento continua `Dell`; GET do script depois do envio → texto de vencido.
  - `test_envio_vencido_410` (forçar `expira_em` no passado via `executar_sql`).
  - `test_envio_grande_413` (corpo de 70 000 bytes) e `test_envio_invalido_422` (`ram_bytes="muito"`).
  - `test_andamento_de_outro_tecnico_404` (segundo técnico criado com `criar_usuario`) e `test_gestor_ve_andamento_de_qualquer_um`.
  - `test_andamento_com_correspondencia` → item existente com `endereco_fisico` igual ao MAC do payload → `correspondencias[0]["motivos"]==["mac"]`, `ativo is True`.
  - `test_job_limpar_coletas_registrado` → depois de `registrar_jobs()`, `"limpar_coletas" in scheduler.jobs_registrados()`.
- [ ] **Step 2:** Rodar `… tests/test_coletor_api.py -q` → FAIL. **Step 3:** implementar. **Step 4:** PASS.
- [ ] **Step 5:** Suíte completa do backend: `… run-backend-tests.sh … coletor -q -m "not mudanca_esperada"` → tudo verde.
- [ ] **Step 6:** Commit `feat(coletor): rotas públicas e do técnico, limpeza diária`.

### Task 5: Campo de número de série em Notebook e Desktop

**Files:**
- Modify: `frontend/src/components/equipment/TypeSpecificFields.tsx` (em `Notebook` e `Desktop`, depois de `host`: `{ key: 'identificador', label: 'Nº de Série' }`)
- Test: o teste existente do componente (`TypeSpecificFields.test.tsx`, se existir; senão `RegisterItemPage.test.tsx`)

- [ ] **Step 1:** Teste: com tipo Notebook, o campo "Nº de Série" aparece; idem Desktop. FAIL.
- [ ] **Step 2:** Implementar; rodar o teste e `npx tsc --noEmit` → PASS.
- [ ] **Step 3:** Commit `feat(cadastro): número de série em Notebook e Desktop`.

### Task 6: Página "Cadastrar esta máquina"

**Files:**
- Create: `frontend/src/api/coletor.ts`, `frontend/src/features/coletor/ColetorPage.tsx`, `frontend/src/features/coletor/estadoColetor.ts`
- Modify: `frontend/src/App.tsx` (rota `register/coletor` dentro de `RequireRole roles={['Gestor','Técnico']}`, antes de `register`), `frontend/src/components/layout/navigation.ts` (item "Cadastrar esta máquina" → `/register/coletor`, `roles: ['Gestor', 'Técnico']`, logo depois de "Cadastro"), `frontend/src/pages/RegisterItemPage.tsx` (botão "Preencher com os dados desta máquina" no cabeçalho, só em `mode==='create'`, navega para `/register/coletor`)
- Test: `frontend/src/features/coletor/ColetorPage.test.tsx`

**Interfaces:**
- Produces (`api/coletor.ts`): `criarColeta(): Promise<ColetaCriada>` (`{id, comando, url_arquivo, expira_em}`), `andamentoColeta(id: number): Promise<AndamentoColeta>` (`{status, expira_em, campos?, tipo_sugerido?, correspondencias?: Correspondencia[]}`), `Correspondencia = {id, tipo, brand, model, status, revenda, ativo, motivos}`.
- Produces (`estadoColetor.ts`): `type EstadoColetor = { origem: 'coletor'; coletaId: number; campos: Record<string, string>; tipoSugerido: string | null }` — passado como `navigate(destino, { state: estado })`; e `lerEstadoColetor(state: unknown): EstadoColetor | null` (valida a forma).

Comportamento:
- Ao abrir, chama `criarColeta` uma vez; mostra o comando num campo só leitura, botão **Copiar comando** (`navigator.clipboard.writeText`; sem clipboard, seleciona o texto e mostra "Selecione e copie com Ctrl+C"), as instruções "Aperte **Windows + R**, cole com **Ctrl+V** e dê **Enter**", "⏳ Aguardando os dados da máquina…", e o link secundário **Baixar o coletor** (`url_arquivo`).
- `useQuery` de `andamentoColeta` com `refetchInterval: 2000` enquanto `aguardando`.
- `vencido` → "O código venceu." + botão **Gerar novo código** (cria outra coleta).
- `recebido` sem correspondências ativas → `navigate('/register', { state })`.
- `recebido` com correspondências → lista cada uma ("#123 Notebook Dell Latitude 3420 · RV-JZ · mesmo MAC"); ativa: botão **Atualizar este cadastro** → `navigate('/edit/123', { state })`; baixada (`ativo=false`): texto "Consta como baixada", link "Ver registro" (`/items/123`), sem botão de atualizar; sempre **Cadastrar como novo mesmo assim** → `/register`.
- Botão **Sair deste computador** → `logout()` do `useAuth` e `navigate('/login')`.
- Textos de motivo: serial → "mesmo número de série", mac → "mesmo MAC", anydesk → "mesmo AnyDesk".

- [ ] **Step 1: Testes que falham** (MSW):
  - `mostra o comando e copia` → clicar em Copiar chama `clipboard.writeText` com o comando.
  - `sem correspondência navega para o cadastro com os campos` → andamento `recebido` → rota `/register` recebe `state.campos.brand === 'Dell'`.
  - `correspondência ativa oferece atualizar` → clicar leva a `/edit/123` com o estado.
  - `correspondência baixada não oferece atualizar` → não existe botão "Atualizar este cadastro"; existe "Cadastrar como novo mesmo assim".
  - `código vencido oferece gerar novo` → clicar chama `POST /api/coletor` de novo.
  - `lerEstadoColetor` recusa estado sem `origem: 'coletor'`.
- [ ] **Step 2:** `npx vitest run src/features/coletor` → FAIL. **Step 3:** implementar. **Step 4:** PASS + `npx tsc --noEmit`.
- [ ] **Step 5:** Commit `feat(coletor): página para coletar os dados da máquina`.

### Task 7: Preenchimento do cadastro, destaque e saída automática

**Files:**
- Modify: `frontend/src/pages/RegisterItemPage.tsx`, `frontend/src/App.tsx` (rota **pública** `/coletor/concluido`)
- Create: `frontend/src/features/coletor/ColetorConcluidoPage.tsx`
- Test: `frontend/src/pages/RegisterItemPage.test.tsx` (casos novos), `frontend/src/features/coletor/ColetorConcluidoPage.test.tsx`

**Interfaces:**
- Consumes: `lerEstadoColetor` e `EstadoColetor` (Task 6); `logout` de `useAuth`.
- Produces: rota `/coletor/concluido?id=<itemId>&saiu=<0|1>`.

Comportamento:
- `RegisterItemPage` lê `lerEstadoColetor(useLocation().state)`.
  - Criação: aplica `tipoSugerido` (se estiver em `equipmentTypes`) e `campos` (`brand`, `model` nos estados próprios; o resto em `specificFields`) uma vez, no primeiro render.
  - Edição: aplica `campos` **depois** da hidratação do item; guarda os valores anteriores para destacar.
- Campo preenchido pelo coletor mostra a marca "do coletor"; em edição, campo cujo valor mudou mostra também "antes: <valor antigo>" (ou "antes: vazio"). Um aviso no topo: "Dados preenchidos pelo coletor. Confira antes de salvar."
- Depois de salvar com sucesso **e** com estado do coletor: `await logout()` dentro de `try`; navega para `/coletor/concluido?id=<id>&saiu=1` (ou `saiu=0` se o logout lançou). Fora do fluxo do coletor, o comportamento atual não muda. Se o anexo da NF falhar (`erroNf`), segue o fluxo atual, sem logout.
- `ColetorConcluidoPage` (sem login): "Máquina cadastrada (#<id>)."; com `saiu=1`, "Você saiu da sua conta neste computador."; com `saiu=0`, "Não foi possível sair da conta automaticamente. Saia pelo menu antes de entregar a máquina."; e o lembrete "Entregue ao colaborador e imprima a etiqueta pelo seu computador (Empréstimo e Etiquetas)."

- [ ] **Step 1: Testes que falham:**
  - `preenche o cadastro com os dados do coletor` → render com `initialEntries=[{pathname:'/register', state:{origem:'coletor', coletaId:1, campos:{brand:'Dell', model:'Latitude 3420', host:'NB-001', identificador:'ABC123'}, tipoSugerido:'Notebook'}}]` → os campos mostram esses valores e o aviso aparece.
  - `edição destaca o que mudou` → item com `host:'NB-OLD'`, estado com `host:'NB-NEW'` → campo mostra `NB-NEW` e o texto "antes: NB-OLD".
  - `salvar pelo coletor sai da conta` → `POST /api/items` ok → `POST /api/auth/logout` chamado → rota `/coletor/concluido?id=…&saiu=1`.
  - `logout falhou ainda mostra a tela final` → logout 500 → `saiu=0`.
  - `fora do coletor não sai da conta` → salvar sem estado não chama logout.
  - `ColetorConcluidoPage` mostra os textos de `saiu=1` e `saiu=0`.
- [ ] **Step 2:** `npx vitest run src/pages/RegisterItemPage.test.tsx src/features/coletor` → FAIL. **Step 3:** implementar. **Step 4:** PASS.
- [ ] **Step 5:** Frontend completo: `npx tsc --noEmit && npm run test:run && npm run build` → tudo verde.
- [ ] **Step 6:** Commit `feat(coletor): cadastro preenchido pelo coletor e saída automática da conta`.
