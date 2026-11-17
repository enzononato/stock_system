# Fase 7 — Painel, alertas e exportações — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fechar o ciclo: painel inicial com indicadores calculados no servidor e alertas calculados na hora, exportação Excel/PDF nas telas operacionais, as fontes de termos, linha telefônica e inventário na linha do tempo do equipamento, a transferência registrada quando um empréstimo muda a unidade do item e o README de operação para o deploy.

**Architecture:** Módulos novos e isolados: `app/db/painel_db.py` (todas as consultas do painel e dos alertas, SQL cru), `app/services/painel.py` (regras e textos dos alertas, limiares lidos de `app.core.config.settings` na hora da chamada), `app/routers/painel.py` (`/api/painel`), `app/services/exportacoes_telas.py` (colunas e linhas de cada tela) + `app/routers/exportacoes.py` (`/api/exportacoes/*`, usando `app/services/exports.py` da F0) e `app/services/linha_do_tempo_fontes_f7.py` (três fontes registradas com `registrar_fonte`). As únicas mudanças em arquivos de outras fases são: uma linha em `MODULOS_DE_FONTES` (F4, previsto), o `issue()` do `InventoryDBManager` (Ruling 23) e botões de exportação acrescentados nas telas. No frontend, `api/painel.ts`, `features/painel/*`, `features/exportacoes/api.ts` e `components/exportacao/BotoesExportar.tsx`; a `DashboardPage` (dona: F7) é reescrita.

**Tech Stack:** FastAPI 0.115 + PyMySQL/DBUtils (SQL cru via `get_cursor`) + MySQL 8; `openpyxl`/`fpdf2` (já instalados pela F0); React 18 + TS + Tailwind + TanStack Query 5 + Vitest/MSW. Nenhuma dependência nova.

**Spec:** `docs/superpowers/specs/2026-11-12-patrimonio-melhorias-design.md` — seção 6 "F7", matriz da seção 3 (linhas 7.2, 12.3, 12.4), seção 4.3 (variáveis `ALERTA_*`), seções 7–9. Contexto comum e rulings: `C:/Users/enzo.jz/dev/patrimonio/_tools/contexto-planejamento.md` (as "Regras acrescentadas pelo controlador" no fim valem acima deste plano). Interfaces consumidas: planos F0, F3, F4, F5, F6, F8 em `docs/superpowers/plans/`.

**Ambiente:** a F7 roda sobre a integração com TODAS as fases já mescladas (F0–F6, F8). Worktree `WT=C:/Users/enzo.jz/dev/patrimonio/f7-painel`, branch `fase/f7-painel` criada a partir de `feat/patrimonio-melhorias` depois do merge da F2 (ajuste se o coordenador usar outro nome). Backend: `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh $WT f7 <args do pytest>` (abreviado `RBT <args>`). Frontend: em `$WT/frontend`, `npx vitest run <arquivo>`, `npx tsc --noEmit`, `npm run build`. Antes da Task 1, conferir `RBT -q` e `npx vitest run` verdes (linha de base da integração).

## Global Constraints

- Spec F7.1: "`GET /painel/resumo`: equipamentos por unidade e por status, em manutenção, termos pendentes, linhas ativas e custo mensal; números calculados no servidor (sem o limite de 500 do cliente)."
- Spec F7.2: "`GET /painel/alertas`, calculado na hora: garantia vencendo (≤ `ALERTA_GARANTIA_DIAS`), termo pendente (> `ALERTA_TERMO_PENDENTE_DIAS`), devolução pendente (`Pendente Devolução` > `ALERTA_DEVOLUCAO_PENDENTE_DIAS`), manutenção parada (> `ALERTA_MANUTENCAO_DIAS`), colaborador desligado ainda com equipamento. Cada alerta leva ao registro." Padrões: `ALERTA_GARANTIA_DIAS` (30), `ALERTA_TERMO_PENDENTE_DIAS` (3), `ALERTA_DEVOLUCAO_PENDENTE_DIAS` (7), `ALERTA_MANUTENCAO_DIAS` (15) — já existem em `app/core/config.py`; nenhuma variável nova.
- Spec seção 2: "Alertas — Só no painel (sem e-mail)". Nenhum job, nenhuma tabela de alertas.
- Spec F7.4: "Exportação Excel e PDF em: Estoque (com os filtros da tela), Histórico, Manutenções, Linhas, Inventário, Termos pendentes e Relatório mensal (o CSV atual continua)." Sempre por `app/services/exports.py` (`Coluna`, `gerar_excel`, `gerar_pdf`, `resposta_arquivo`); unidade/revenda com `Coluna(..., maiusculas=True)`; CPF nunca sai por inteiro em Excel/PDF.
- Aceite F7: "números do painel batem com consultas diretas; cada alerta aparece com o dado de teste certo e some quando resolvido; arquivos exportados abrem e respeitam os filtros."
- Migrações: faixa `0700–0799`; **esta fase não cria migração** (tudo é consulta sobre tabelas existentes).
- 403 com as mensagens padrão: `Acesso restrito ao Gestor.` / `Acesso restrito a Gestor ou Técnico.`
- Código que lê limiares usa `from app.core import config` e `config.settings.ALERTA_*` no momento da chamada (o `tests/test_auth.py` recarrega o módulo de configuração).
- Arquivos compartilhados (`backend/app/main.py`, `backend/app/dependencies.py`, `frontend/src/test/handlers.ts`): só acrescentar linhas. Não alterar `StatusPill`, `Pagination`, `Button`, `StatBlock`, `lib/status.ts`, registros de abas. Nas telas de outras fases (StockPage, HistoryPage, ManutencoesPage, LinhasPage, TermsPage, ReportPage) só se ACRESCENTA o botão de exportação e o import; nada é reorganizado.
- Telas: paleta pelos tokens (nenhuma cor fora deles; nada de `bg-emerald-*`), `StatusPill`, `Button loading`, `upper()` para unidade, texto azul só `text-brand-text`, utilizável em 390 px.
- Textos, mensagens e commits em português; commits `tipo(escopo): descrição` + linha em branco + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Decisões fixadas por este plano

| Tema | Decisão |
|---|---|
| Permissões do painel | `GET /api/painel/resumo`: qualquer usuário autenticado (só contagens; o Jovem Aprendiz já vê o estoque e as linhas com custo). `GET /api/painel/alertas`: Gestor e Técnico. Justificativa: alertas são pendências que só Gestor/Técnico resolvem (devolução, termo, manutenção, linhas e baixa são rotas deles) e trazem nome de colaborador desligado (minimização LGPD); a `DashboardPage` atual já esconde a "Atenção Operacional" do Jovem Aprendiz. O painel do Aprendiz mostra só os indicadores e não chama `/alertas`. |
| Permissões de exportação | Estoque, Histórico, Manutenções, Linhas e Termos pendentes: `gestor_or_tecnico` (mesma regra das exportações da F5/F6). Relatório mensal Excel/PDF: `gestor_only` (mesma regra do CSV atual). O Jovem Aprendiz não vê botão de exportar. |
| Rotas de exportação | Router próprio `/api/exportacoes/{estoque,historico,manutencoes,linhas,termos-pendentes,relatorio-mensal}?formato=xlsx|pdf&<filtros da tela>`; `formato: Literal["xlsx","pdf"]` obrigatório (inválido → 422). Não mexe nos routers das outras fases nem na ordem de `/api/items/{item_id}`. Inventário continua em `GET /api/inventarios/{id}/relatorio/exportar` (F6, só conferido). |
| Limite de linhas | `LIMITE_LINHAS = {"xlsx": 50_000, "pdf": 5_000}`. Acima → 400 `f"A exportação em {'PDF' if pdf else 'Excel'} está limitada a {limite:,} linhas e esta consulta tem {total:,}. Refine os filtros" + (" ou exporte em Excel." if pdf else ".")`, com separador de milhar `.` (ex.: `5.000`). A consulta pede `limite + 1` linhas e usa o `total` devolvido. |
| Termos pendentes | Solicitações de `termo_solicitacoes` com status em `PENDENTES = ("agendado","enviado","manual","erro")` (F3) de itens ativos **mais** os itens legados `Pendente` sem nenhuma solicitação pendente (empréstimos anteriores ao deploy da F3), estes com `status="sem_solicitacao"`, `tipo="entrega"` e data = última `history` `Empréstimo` não estornada. O indicador do painel `termos.pendentes` conta só as solicitações; os legados vão em `termos.sem_solicitacao`. |
| Idade de cada alerta | Termo: data de referência `agendado → agendado_para`, `enviado → COALESCE(enviado_em, solicitado_em)`, demais → `solicitado_em`, comparada com `agora_local()` (F3 grava em horário de Brasília); legado comparado com `datetime.now()`. Devolução: última `history` `Devolução` não estornada (sem registro → entra no alerta com `desde=None`). Manutenção parada: último movimento `GREATEST(aberta_em, COALESCE(enviada_em, aberta_em), COALESCE(em_reparo_em, aberta_em))` das ordens em `STATUS_EM_ANDAMENTO` (F4). Todos: alerta quando `agora - desde > timedelta(days=N)` (estrito). Garantia: itens ativos com `0 <= (garantia_ate - hoje).days <= ALERTA_GARANTIA_DIAS` (garantia já vencida não alerta). |
| Desligado com equipamento | Mesma regra de `AcessoDBManager.equipamentos_com_colaborador` (F8): itens `is_active=1` com status em `STATUS_COM_COLABORADOR` vinculados por `colaborador_id`, ou, sem vínculo, com o CPF (só dígitos) igual ao do colaborador `situacao='desligado'`. Feito numa consulta só (UNION), não um loop por colaborador. Um alerta por item. |
| Alerta extra (F5) | `linha_aparelho_baixado`: linha `status='ativa'` com `item_id` de item `is_active=0`. |
| Ruling 23 | `issue()` compara `items.revenda` atual com a nova (`strip().casefold()`, igual à F4); se mudou, na MESMA transação: `UPDATE items` também grava `local_id=NULL` (o local antigo é de outra unidade) e `INSERT transferencias (item_id, de_unidade, para_unidade, de_local_id, para_local_id, motivo, transferido_por, transferido_em)` com `para_local_id=NULL`, `motivo="Empréstimo para colaborador de outra unidade"`, `transferido_por=logged_user`, `transferido_em=dt_issue`. Sem linha `history` extra (a fonte de transferências já mostra o evento). O estorno do empréstimo não restaura `revenda` (comportamento atual), então a transferência continua verdadeira. |
| Linha do tempo | Fontes novas num módulo da F7 acrescentado a `MODULOS_DE_FONTES`; eventos `edicao` de linha telefônica ficam fora (são dados da linha, não do aparelho); e-mail do destinatário do termo nunca entra no texto. |
| Painel do Gestor/Técnico | Remove as consultas `listItemsPaginated({limit})` e "Últimas movimentações"; ficam indicadores, alertas, equipamentos por unidade, gráfico de movimentação e fluxos rápidos. |

## Review Focus

- Banco de produção com dados antigos: item com `revenda` NULL/vazia, empréstimo `Pendente` anterior à F3 (sem solicitação), `Pendente Devolução` sem linha `Devolução` no histórico, item legado sem `colaborador_id` só com CPF mascarado (`529.982.247-25`) de um desligado → resumo agrupa em "Sem unidade", os alertas aparecem (legado e `desde=None`) sem erro. Testes nas Tasks 3 e 4.
- Exportação grande demais (PDF com mais de 5.000 linhas) → 400 com mensagem clara, e o frontend mostra essa mensagem mesmo com a resposta vindo como `Blob` (não "Não foi possível concluir a operação."). Testes nas Tasks 5 e 7.
- Mudar `ALERTA_*` no `.env` muda o painel sem reiniciar código em cache: limiar lido na hora (monkeypatch em `config.settings`) → alerta aparece/some. Teste na Task 4.
- Alerta resolvido some na próxima leitura: devolver, assinar, retornar a manutenção, desvincular a linha → nenhum resíduo (não há tabela de alertas). Teste na Task 4.
- Jovem Aprendiz no painel: indicadores carregam, nenhuma chamada a `/api/painel/alertas` (sem 403 no console), nenhum botão de exportar nas telas. Testes nas Tasks 8 e 9.

## Mapa de arquivos

Backend (criar): `app/db/painel_db.py`, `app/schemas/painel.py`, `app/services/painel.py`, `app/routers/painel.py`, `app/services/exportacoes_telas.py`, `app/routers/exportacoes.py`, `app/services/linha_do_tempo_fontes_f7.py`; testes `tests/test_emprestimo_transferencia.py`, `tests/test_linha_do_tempo_f7.py`, `tests/test_painel_resumo.py`, `tests/test_painel_alertas.py`, `tests/test_exportacoes.py`, `tests/test_exportacoes_termos_relatorio.py`.
Backend (alterar): `app/db/inventory_manager_db.py` (só `issue`), `app/services/linha_do_tempo.py` (uma linha em `MODULOS_DE_FONTES`), `app/dependencies.py` (acréscimo), `app/main.py` (imports, `include_router`, rótulo `formato`).
Frontend (criar): `src/api/painel.ts`, `src/features/painel/{alertas.ts,IndicadoresPainel.tsx,EquipamentosPorUnidade.tsx,ListaAlertas.tsx}`, `src/features/exportacoes/api.ts` (+ `.test.ts`), `src/components/exportacao/BotoesExportar.tsx` (+ `.test.tsx`), `src/pages/DashboardPage.test.tsx`.
Frontend (alterar): `src/pages/DashboardPage.tsx` (reescrita), e só acréscimos em `StockPage.tsx`, `HistoryPage.tsx`, `ManutencoesPage.tsx`, `LinhasPage.tsx`, `TermsPage.tsx`, `ReportPage.tsx` e nos testes delas.
Docs: `README.md`.

---

### Task 1: Ruling 23 — empréstimo que muda a unidade registra transferência

**Files:**
- Modify: `backend/app/db/inventory_manager_db.py` (só o método `issue`)
- Test: `backend/tests/test_emprestimo_transferencia.py`

**Interfaces:**
- Consumes: tabela `transferencias` (F4, 0404); fonte `fonte_transferencias` (F4).
- Produces: `MOTIVO_TRANSFERENCIA_EMPRESTIMO = "Empréstimo para colaborador de outra unidade"` (constante de módulo em `inventory_manager_db.py`). Assinatura, mensagens e retorno de `issue(...)` não mudam.

- [ ] **Step 1: Testes** (`pytestmark = pytest.mark.integration`; item via `criar_item()` → revenda `"Revalle Juazeiro"`; `hoje = datetime.now().strftime("%d/%m/%Y")`; chamada `inv_manager.issue(item["id"], "Fulano de Tal", "52998224725", "101 - Puxada", "Analista", "TI", <revenda>, hoje, "gestor.teste")`):
  - `test_emprestimo_para_outra_unidade_grava_transferencia`: item com `local_id` de um local de Juazeiro (`executar_sql("UPDATE items SET local_id=%s ...")` com local criado via `client_gestor.post("/api/locais", ...)` como em `test_transferencias.py`); revenda `"Revalle Petrolina"` → `ok is True`; exatamente uma linha em `transferencias` com `de_unidade == "Revalle Juazeiro"`, `para_unidade == "Revalle Petrolina"`, `de_local_id == <local>`, `para_local_id is None`, `motivo == "Empréstimo para colaborador de outra unidade"`, `transferido_por == "gestor.teste"`, `transferido_em.date() == date.today()`; item `revenda == "Revalle Petrolina"`, `local_id is None`.
  - `test_emprestimo_na_mesma_unidade_nao_grava`: revenda `"REVALLE JUAZEIRO "` → nenhuma linha em `transferencias`; `local_id` mantido.
  - `test_falha_na_transferencia_desfaz_o_emprestimo`: `executar_sql("RENAME TABLE transferencias TO transferencias_f7_tmp")`; `issue(..., "Revalle Petrolina", ...)` → `(False, "Erro ao processar a operação. Tente novamente.")`; item continua `Disponível` com a revenda antiga e sem linha `Empréstimo` nova; `finally`: `RENAME TABLE transferencias_f7_tmp TO transferencias`.
  - `test_linha_do_tempo_mostra_a_transferencia_do_emprestimo`: depois do empréstimo para Petrolina, `client_gestor.get(f"/api/items/{id}/linha-do-tempo")` tem um evento `tipo == "transferencia"`, `titulo == "Transferência Revalle Juazeiro → Revalle Petrolina"`, `descricao` começando com `"Empréstimo para colaborador de outra unidade"`.
  - `test_estorno_do_emprestimo_mantem_a_transferencia`: estornar a linha `Empréstimo` (`inv_manager.reverse_history_entry(<id>, "gestor.teste")`) → a linha de `transferencias` continua e `revenda == "Revalle Petrolina"`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_emprestimo_transferencia.py -q` → FAIL (nenhuma transferência gravada).
- [ ] **Step 3: Implementar** em `issue`: o `find(pid)` já traz `revenda` e `local_id`; calcular `muda = (item["revenda"] or "").strip().casefold() != (revenda or "").strip().casefold()`; dentro do mesmo `with get_cursor(...)`, o `UPDATE items` acrescenta `local_id=NULL` só quando `muda` (dois textos de SQL ou `local_id=%s` com o valor atual quando não muda) e, depois do `INSERT history`, o `INSERT transferencias` descrito em "Decisões". Nenhuma mensagem existente muda.
- [ ] **Step 4: Rodar** `RBT tests/test_emprestimo_transferencia.py tests/test_loan_workflow.py tests/test_loan_colaborador.py tests/test_reversal.py tests/test_transferencias.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(emprestimo): transferência registrada quando o empréstimo muda a unidade do equipamento`.

### Task 2: Fontes da linha do tempo — termos, linha telefônica e inventário

**Files:**
- Create: `backend/app/services/linha_do_tempo_fontes_f7.py`, `backend/tests/test_linha_do_tempo_f7.py`
- Modify: `backend/app/services/linha_do_tempo.py` (só a tupla: `MODULOS_DE_FONTES = ("app.services.linha_do_tempo_fontes", "app.services.linha_do_tempo_fontes_f7")` e o comentário acima dela passa a dizer que a F7 já acrescentou o seu módulo)

**Interfaces:**
- Consumes: `Evento`, `registrar_fonte` (F4); `TermoDBManager` (`listar_por_item(item_id, limit=, offset=)`, F3) ou SQL próprio em `termo_solicitacoes`; `ROTULO_TIPO` (`app.services.termos`, F3); `LinhaDBManager().eventos_do_aparelho(item_id)` (F5); `formatar_numero_linha` (`app.schemas.linhas`, F5); `InventarioDBManager().historico_do_item(item_id)` (F6); `ROTULOS_CATEGORIA` (`app.services.inventario`, F6).
- Produces (cada uma com `@registrar_fonte`, nesta ordem de registro):
  - `fonte_termos(item_id) -> list[Evento]` — `origem="termo"`, `referencia_id=termo.id`; por solicitação do item (`SELECT ... FROM termo_solicitacoes t LEFT JOIN colaboradores c ON c.id=t.colaborador_id WHERE t.item_id=%s ORDER BY t.id`), com `r = ROTULO_TIPO[tipo]`:
    - `solicitado_em` → `("termo_solicitado", f"Termo de {r} solicitado")`, `usuario=solicitado_por`, `descricao`: com `agendado_para` → `f"Agendado para {dd/mm/aaaa HH:MM}"`; status `manual` → `motivo`; senão `None`.
    - `enviado_em` → `("termo_enviado", f"Termo de {r} enviado para assinatura")`, `descricao=f"Destinatário: {colaborador_nome}"` quando houver.
    - `assinado_em` → `("termo_assinado", f"Termo de {r} assinado")`, `descricao=colaborador_nome`.
    - `cancelado_em` → `("termo_cancelado", "Solicitação de termo cancelada")`, `descricao=motivo`.
    - status `recusado` / `expirado` / `erro` → evento em `atualizado_em`: `("termo_recusado", f"Termo de {r} recusado")`, `("termo_expirado", f"Termo de {r} expirado")`, `("termo_erro", "Assinatura recebida, mas a confirmação falhou")`, `descricao=motivo`.
  - `fonte_linha_telefonica(item_id) -> list[Evento]` — `origem="linha"`, `tipo=f"linha_{e['tipo']}"`, `referencia_id=e['id']`, `usuario=por_usuario`, a partir de `eventos_do_aparelho(item_id)` sem os de tipo `edicao`; `num = formatar_numero_linha(linha_numero)`; título: `troca_aparelho` com `item_id_anterior == item_id` → `f"Linha {num} saiu deste aparelho"`, `troca_aparelho` com `item_id == item_id` → `f"Linha {num} passou para este aparelho"`, demais → `f"Linha {num}: {ROTULO_EVENTO_LINHA[tipo]}"` com `ROTULO_EVENTO_LINHA = {"ativacao": "Ativação", "troca_aparelho": "Troca de aparelho", "troca_usuario": "Troca de usuário", "portabilidade": "Portabilidade", "suspensao": "Suspensão", "reativacao": "Reativação", "cancelamento": "Cancelamento"}` (mesmos rótulos de `frontend/src/features/linhas/eventos.ts`); `descricao=f"{de_valor or '—'} → {para_valor or '—'}"` quando algum dos dois existe, senão `observacao`.
  - `fonte_inventario(item_id) -> list[Evento]` — `origem="inventario"`, `referencia_id=inventario_id`, `usuario=por`, de `historico_do_item(item_id)`; `evento == "leitura"` → `("inventario_leitura", f"Inventário #{inventario_id} ({unidade_nome.upper()}): {ROTULOS_CATEGORIA[resultado]}")`, `descricao=f"Local onde foi encontrado: {local_encontrado_nome}"` ou `None`; `nao_encontrado` → `("inventario_nao_encontrado", f"Inventário #{inventario_id} ({unidade_nome.upper()}): Não encontrado")`. Linhas com `data` nula são ignoradas.

- [ ] **Step 1: Testes** (`tests/test_linha_do_tempo_f7.py`, integration; importar `criar_colaborador`, `unidade_juazeiro` de `tests.fixtures_termos`; `criar_aparelho`, `dados_linha`, `nova_linha` de `tests.fixtures_linhas` — **não** importar o `criar_colaborador` de `fixtures_linhas`, que tem o mesmo nome):
  - `test_fontes_da_f7_registradas`: `carregar_fontes()`; `fontes_registradas()` contém `"fonte_termos"`, `"fonte_linha_telefonica"`, `"fonte_inventario"` depois das quatro da F4.
  - `test_termo_solicitado_enviado_e_assinado`: `TermoDBManager().criar(item_id=..., tipo="entrega", status="manual", solicitado_por="gestor.teste", solicitado_em=datetime(2026, 1, 1, 10, 0), colaborador_id=c["id"], destinatario_email="maria.souza@gmail.com")` e `atualizar(id, status="assinado", enviado_em=datetime(2026, 1, 1, 10, 5), assinado_em=datetime(2026, 1, 2, 9, 0))` → `montar_linha_do_tempo(item)` traz, nessa ordem, `termo_solicitado` ("Termo de entrega solicitado", `usuario == "gestor.teste"`), `termo_enviado` ("Termo de entrega enviado para assinatura", `descricao == "Destinatário: Maria Souza"`), `termo_assinado` ("Termo de entrega assinado"); nenhum `descricao`/`titulo` contém `"@"`.
  - `test_termo_cancelado_mostra_motivo` (`status="cancelado"`, `cancelado_em`, `motivo="Substituída pela solicitação #9."`) → evento `termo_cancelado` com essa descrição.
  - `test_linha_que_passou_pelo_aparelho`: mesma montagem de `tests/test_linhas_eventos.py::test_eventos_do_aparelho_inclui_linha_que_saiu` (linha criada no aparelho A e movida para B) → linha do tempo de A tem `linha_ativacao` (`"Linha (74) 99999-0001: Ativação"` — use o número da fixture) e `linha_troca_aparelho` com título `"... saiu deste aparelho"`; a de B tem `"... passou para este aparelho"`; uma edição de plano da linha (PUT só `plano`) não gera evento `linha_edicao`.
  - `test_inventario_lido_e_nao_encontrado`: mesma montagem de `tests/test_inventario_relatorio.py::test_historico_do_item` → eventos `inventario_leitura` e `inventario_nao_encontrado`, nessa ordem, com títulos começando por `f"Inventário #{id1} ("` e `f"Inventário #{id2} ("` e terminando em `"Conferido"` / `"Não encontrado"`.
  - `test_api_traz_fontes_da_f7_sem_erro`: `GET /api/items/{id}/linha-do-tempo` (gestor) → 200, `fontes_com_erro == []`, há evento com `origem == "termo"`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_linha_do_tempo_f7.py -q` → FAIL (módulo inexistente).
- [ ] **Step 3: Implementar** o módulo e a linha em `MODULOS_DE_FONTES`.
- [ ] **Step 4: Rodar** `RBT tests/test_linha_do_tempo_f7.py tests/test_linha_do_tempo.py -q` → PASS (o teste da F4 que conta os tipos de um item sem termo/linha/inventário continua igual).
- [ ] **Step 5: Commit** `feat(ciclo-vida): termos, linha telefônica e inventário na linha do tempo do equipamento`.

### Task 3: `GET /api/painel/resumo`

**Files:**
- Create: `backend/app/db/painel_db.py`, `backend/app/schemas/painel.py`, `backend/app/routers/painel.py`, `backend/tests/test_painel_resumo.py`
- Modify: `backend/app/dependencies.py` (import + `get_painel_db`), `backend/app/main.py` (`from app.routers import painel` e `app.include_router(painel.router)` depois do último `include_router`)

**Interfaces:**
- Consumes: `ITEM_STATUSES` (`app.core.config`, F4); `STATUS_EM_ANDAMENTO` (`app.db.manutencao_db`); `PENDENTES` (`app.db.termo_db`); `LinhaDBManager.relatorio_custo()` (F5: `{"unidades": [{unidade_id, unidade_nome, quantidade, custo_total}], "total_linhas", "custo_total", "gerado_em"}`).
- Produces:
  - `PainelDBManager.contagem_itens() -> list[dict]` — `SELECT NULLIF(TRIM(revenda), '') AS unidade, status, COUNT(*) AS n FROM items WHERE is_active=1 GROUP BY 1, 2`.
  - `PainelDBManager.ordens_em_andamento() -> int`; `PainelDBManager.termos_por_status() -> dict[str, int]` (só `PENDENTES`, itens ativos, as 4 chaves sempre presentes); `PainelDBManager.legados_sem_solicitacao() -> int` (itens ativos `Pendente` sem solicitação em `PENDENTES`).
  - `PainelDBManager.resumo() -> dict` com o formato abaixo; `por_unidade` ordenado por nome (`casefold`), "sem unidade" (`unidade: None`) por último; `por_status` sempre com as 5 chaves de `ITEM_STATUSES` (status fora delas só conta no `total`).
  - `get_painel_db() -> PainelDBManager` (`@lru_cache`).
  - Schemas (`app/schemas/painel.py`): `EquipamentosPorUnidade(unidade: str | None, total: int, por_status: dict[str, int])`; `ResumoEquipamentos(total: int, por_status: dict[str, int], por_unidade: list[EquipamentosPorUnidade])`; `ResumoManutencao(equipamentos_em_manutencao: int, ordens_em_andamento: int)`; `ResumoTermos(pendentes: int, por_status: dict[str, int], sem_solicitacao: int)`; `LinhasPorUnidade(unidade_nome: str, quantidade: int, custo_total: Decimal)`; `ResumoLinhas(ativas: int, custo_mensal: Decimal, por_unidade: list[LinhasPorUnidade])`; `ResumoPainel(gerado_em: datetime, equipamentos: ResumoEquipamentos, manutencao: ResumoManutencao, termos: ResumoTermos, linhas: ResumoLinhas)`. Decimais serializados como `"49.90"` (mesmo padrão da F5).
  - Router `APIRouter(prefix="/api/painel", tags=["painel"])`: `GET /resumo` (`get_current_user`) → `ResumoPainel`. `equipamentos_em_manutencao = por_status["Em Manutenção"]`.

- [ ] **Step 1: Testes** (`tests/test_painel_resumo.py`, integration):
  - `test_resumo_vazio`: banco limpo → `equipamentos.total == 0`, `por_status == {"Disponível": 0, "Indisponível": 0, "Pendente": 0, "Pendente Devolução": 0, "Em Manutenção": 0}`, `por_unidade == []`, `termos == {"pendentes": 0, "por_status": {"agendado": 0, "enviado": 0, "manual": 0, "erro": 0}, "sem_solicitacao": 0}`, `linhas.ativas == 0`, `linhas.custo_mensal == "0.00"`.
  - `test_resumo_bate_com_consultas_diretas` (aceite): 3 itens Juazeiro (1 `Disponível`, 1 `Indisponível` e 1 `Em Manutenção` via `executar_sql`), 2 Petrolina (`Pendente`, `Pendente Devolução`), 1 com `revenda=NULL` e 1 baixado (`is_active=0`); o teste compara cada contagem com `SELECT revenda, status, COUNT(*) FROM items WHERE is_active=1 GROUP BY revenda, status` feito no próprio teste, e afirma `total == 6`; `por_unidade[0]["unidade"] == "Revalle Juazeiro"` com `total == 3`; último `unidade is None` com `total == 1` (Review Focus); `manutencao.equipamentos_em_manutencao == 1`.
  - `test_resumo_conta_ordens_e_termos`: uma manutenção aberta (`POST /api/manutencoes`) e uma cancelada → `ordens_em_andamento == 1`; duas solicitações (`manual`, `enviado`) e uma `assinado` via `TermoDBManager().criar` → `termos.pendentes == 2`, `por_status["manual"] == 1`; item legado `Pendente` sem solicitação (fixture `item_pendente`) → `sem_solicitacao == 1`.
  - `test_resumo_linhas_usa_relatorio_de_custo`: duas linhas ativas (custos `49.90` e `30.10`) e uma suspensa → `linhas.ativas == 2`, `custo_mensal == "80.00"`, `por_unidade[0]["quantidade"] == 2`.
  - `test_resumo_liberado_para_todos_os_papeis`: `client_aprendiz` e `client_tecnico` → 200; sem token → 401.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_painel_resumo.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `RBT tests/test_painel_resumo.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(painel): resumo de indicadores calculado no servidor`.

### Task 4: `GET /api/painel/alertas`

**Files:**
- Create: `backend/app/services/painel.py`, `backend/tests/test_painel_alertas.py`
- Modify: `backend/app/db/painel_db.py`, `backend/app/schemas/painel.py`, `backend/app/routers/painel.py`

**Interfaces:**
- Consumes: Task 3; `agora_local()` (`app.core.tempo`, F3); `STATUS_COM_COLABORADOR`, `descrever_item` (`app.db.acesso_db`, F8); `ROTULOS_STATUS` (`app.db.manutencao_db`); `ROTULO_TIPO` (`app.services.termos`); `formatar_numero_linha` (F5).
- Produces (`app/db/painel_db.py`, consultas cruas, sem limiar — o serviço filtra):
  - `termos_pendentes(*, busca: str | None = None, limit: int | None = None) -> tuple[list[dict], int]` — chaves `termo_id (None no legado), item_id, item_tipo, item_brand, item_model, item_identificador, item_revenda, colaborador_id, colaborador_nome (c.nome, senão i.assigned_to), tipo, status ("sem_solicitacao" no legado), solicitado_por, solicitado_em, enviado_em, agendado_para, motivo, referencia (datetime da regra de idade), legado: bool`; ordem `referencia` crescente; `busca` com `LIKE` em `i.brand`, `i.model`, `i.identificador`, `c.nome`, `i.assigned_to`; total = solicitações + legados. Também usada pela exportação (Task 6).
  - `garantias_entre(inicio: date, fim: date) -> list[dict]` (itens `is_active=1` com `garantia_ate BETWEEN inicio AND fim`; o serviço passa `hoje` e `hoje + timedelta(days=N)`): `id, tipo, brand, model, identificador, revenda, garantia_ate`.
  - `devolucoes_pendentes() -> list[dict]` (itens ativos `Pendente Devolução` com `desde` = `MAX(h.data_operacao)` das `Devolução` não estornadas, `LEFT JOIN`, `NULL` se não houver) com `assigned_to`, `colaborador_nome`.
  - `manutencoes_em_andamento() -> list[dict]` (`m.id, m.item_id, m.status, m.assistencia, ultimo_movimento` pela expressão `GREATEST(...)`, campos do item).
  - `desligados_com_equipamento() -> list[dict]` — `UNION` de (a) `JOIN colaboradores c ON c.id = i.colaborador_id` e (b) `i.colaborador_id IS NULL AND c.cpf = REPLACE(REPLACE(REPLACE(REPLACE(i.cpf,'.',''),'-',''),' ',''),'/','')`, ambos com `c.situacao='desligado'`, `i.is_active=1`, `i.status IN STATUS_COM_COLABORADOR`: `item_id, tipo, brand, model, identificador, status, revenda, colaborador_id, colaborador_nome`.
  - `linhas_em_aparelho_baixado() -> list[dict]`: `linha_id, numero, unidade_nome, item_id, tipo, brand, model, identificador`.
- Produces (`app/services/painel.py`):
  - `TipoAlerta = Literal["colaborador_desligado", "devolucao_pendente", "termo_pendente", "manutencao_parada", "linha_aparelho_baixado", "garantia_vencendo"]`; `SEVERIDADE = {"colaborador_desligado": "alta", "devolucao_pendente": "alta", "termo_pendente": "media", "manutencao_parada": "media", "linha_aparelho_baixado": "media", "garantia_vencendo": "baixa"}`.
  - `@dataclass(frozen=True) Limites(garantia_dias: int, termo_pendente_dias: int, devolucao_pendente_dias: int, manutencao_dias: int)`; `limites_atuais() -> Limites` (lê `config.settings.ALERTA_*` na chamada).
  - `calcular_alertas(db: PainelDBManager, *, agora: datetime | None = None, agora_termos: datetime | None = None, limites: Limites | None = None) -> dict` → `{"gerado_em", "limites", "total", "por_tipo" (6 chaves sempre), "alertas"}`; cada alerta `{"chave": f"{tipo}:{id}", "tipo", "severidade", "titulo", "descricao", "link", "item_id", "colaborador_id", "desde", "dias"}`; `dias = (agora - desde).days` (garantia: `(garantia_ate - hoje).days`); ordem: severidade (`alta`, `media`, `baixa`), `desde is None` primeiro, `dias` decrescente, `chave`. `equip = f"#{item_id} {descrever_item(linha) or 'Equipamento'}"`. Textos exatos:
    - `garantia_vencendo`: título `"Garantia vence hoje"` (dias 0) ou `f"Garantia vence em {dias} dia(s)"`; descrição `f"{equip} · {revenda.upper() or 'SEM UNIDADE'}"`; link `f"/items/{id}"`.
    - `termo_pendente`: título `f"Termo de {ROTULO_TIPO[tipo]} pendente há {dias} dias"`; legado `f"Termo de entrega sem assinatura há {dias} dias"`; descrição `f"{equip} · {colaborador_nome or 'sem colaborador'} · {ROTULO_STATUS_TERMO[status]}"` com `ROTULO_STATUS_TERMO = {"agendado": "Agendado", "enviado": "Enviado", "manual": "Assinatura manual", "erro": "Erro", "sem_solicitacao": "Sem solicitação"}`; link `f"/items/{item_id}?aba=termos"`; chave `f"termo_pendente:{termo_id}"` ou `f"termo_pendente:item{item_id}"` no legado.
    - `devolucao_pendente`: título `f"Devolução pendente há {dias} dias"` ou `"Devolução pendente sem data registrada"`; descrição `f"{equip} · {colaborador_nome or assigned_to or 'sem colaborador'}"`; link `f"/items/{id}"`.
    - `manutencao_parada`: título `f"Manutenção parada há {dias} dias"`; descrição `f"{equip} · {ROTULOS_STATUS[status]} · {assistencia or 'sem assistência'}"`; link `f"/items/{item_id}?aba=manutencoes"`; chave pelo id da manutenção.
    - `colaborador_desligado`: título `f"{colaborador_nome} (desligado) ainda está com equipamento"`; descrição `f"{equip} · {status}"`; link `f"/items/{item_id}"`; `colaborador_id` preenchido.
    - `linha_aparelho_baixado`: título `f"Linha {formatar_numero_linha(numero)} ativa em aparelho baixado"`; descrição `f"{equip} · {unidade_nome.upper()}"`; link `f"/linhas?busca={numero}"`; chave pelo id da linha.
  - Schemas: `LimitesAlerta`, `Alerta` (`desde: date | datetime | None`, `dias: int | None`), `AlertasPainel(gerado_em: datetime, limites: LimitesAlerta, total: int, por_tipo: dict[str, int], alertas: list[Alerta])`.
  - `GET /api/painel/alertas` (`gestor_or_tecnico`) → `AlertasPainel` (chama `calcular_alertas(db)`).

- [ ] **Step 1: Testes** (`tests/test_painel_alertas.py`, integration; helper `_alertas(client, tipo)` = `[a for a in client.get("/api/painel/alertas").json()["alertas"] if a["tipo"] == tipo]`; datas recuadas com `executar_sql`):
  - `test_garantia_vencendo_no_limite`: itens com `garantia_ate` = hoje, hoje+30, hoje+31 e ontem → só os dois primeiros; títulos `"Garantia vence hoje"` e `"Garantia vence em 30 dia(s)"`; `link == f"/items/{id}"`.
  - `test_termo_pendente_respeita_o_limiar_e_some_quando_assinado`: solicitação `manual` com `solicitado_em = agora_local() - timedelta(days=3, hours=1)` → um alerta `"Termo de entrega pendente há 3 dias"`, `link == f"/items/{item}?aba=termos"`; outra com `- timedelta(days=2, hours=23)` → sem alerta; `TermoDBManager().atualizar(id, status="assinado", assinado_em=...)` → o alerta some (aceite).
  - `test_termo_agendado_conta_da_data_agendada`: `agendado` com `solicitado_em` há 10 dias e `agendado_para` amanhã → sem alerta.
  - `test_termo_legado_sem_solicitacao` (Review Focus): `item_pendente` com a `history` `Empréstimo` recuada 4 dias → título `"Termo de entrega sem assinatura há 4 dias"`, `chave == f"termo_pendente:item{id}"`.
  - `test_devolucao_pendente`: `forcar_devolucao_iniciada(item)` e a linha `Devolução` recuada 8 dias → `"Devolução pendente há 8 dias"`; com 6 dias → nada; item `Pendente Devolução` forçado por SQL sem linha `Devolução` → `"Devolução pendente sem data registrada"`, `desde is None` (Review Focus).
  - `test_manutencao_parada_usa_ultimo_movimento`: ordem aberta há 20 dias e enviada há 5 → sem alerta; outra aberta há 16 dias sem movimento → `"Manutenção parada há 16 dias"`, link `?aba=manutencoes`; `POST /api/manutencoes/{id}/cancelar` → some.
  - `test_colaborador_desligado_por_vinculo_e_por_cpf_legado` (Review Focus): colaborador `desligado` (`criar_colaborador(situacao="desligado", cpf="52998224725")`); item A `Indisponível` com `colaborador_id`; item B legado `Indisponível` sem `colaborador_id` e `cpf="529.982.247-25"`; item C `Disponível` com `colaborador_id` esquecido → alertas para A e B, não C; título `"Maria Souza (desligado) ainda está com equipamento"`; o conjunto de `item_id` é igual a `{e["id"] for e in AcessoDBManager().equipamentos_com_colaborador(colab_id)}`.
  - `test_linha_ativa_em_aparelho_baixado`: linha ativa no aparelho (`nova_linha`), aparelho com `is_active=0` via SQL → um alerta com `link == f"/linhas?busca={numero}"`; linha suspensa → nada.
  - `test_limiares_lidos_na_hora` (Review Focus): `monkeypatch.setattr(config.settings, "ALERTA_GARANTIA_DIAS", 5)` (com `from app.core import config`) → item com garantia em 10 dias não alerta; com `20` alerta; a resposta traz `limites["garantia_dias"] == 20`.
  - `test_ordem_e_contagens`: com um alerta de cada tipo → `alertas[0]["severidade"] == "alta"`, `por_tipo` tem as 6 chaves e `total == sum(por_tipo.values())`.
  - `test_alertas_so_gestor_e_tecnico`: `client_aprendiz` → 403 `"Acesso restrito a Gestor ou Técnico."`; técnico → 200.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_painel_alertas.py -q` → FAIL.
- [ ] **Step 3: Implementar** as consultas, o serviço e a rota.
- [ ] **Step 4: Rodar** `RBT tests/test_painel_alertas.py tests/test_painel_resumo.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(painel): alertas de garantia, termo, devolução, manutenção, desligado e linha em aparelho baixado`.

### Task 5: Exportações de Estoque, Histórico, Manutenções e Linhas

**Files:**
- Create: `backend/app/services/exportacoes_telas.py`, `backend/app/routers/exportacoes.py`, `backend/tests/test_exportacoes.py`
- Modify: `backend/app/main.py` (`from app.routers import exportacoes`, `app.include_router(exportacoes.router)`, `_ROTULOS_DE_CAMPO` ganha `"formato": "Formato"`)

**Interfaces:**
- Consumes: `exports.Coluna`, `gerar_excel`, `gerar_pdf`, `resposta_arquivo` (F0); `InventoryDBManager.list_items(..., limit=, offset=0)` e `montar_order_by_items` (F1); `InventoryDBManager.list_history(search=, item_id=, limit=)`; `ManutencaoDBManager.listar(...)` + `ROTULOS_STATUS` (F4); `LinhaDBManager.listar(...)` + `formatar_numero_linha` (F5); `ColaboradorDBManager.obter`, `get_unidade_db` (nome da unidade para o filtro de linhas).
- Produces (`app/services/exportacoes_telas.py`):
  - `LIMITE_LINHAS = {"xlsx": 50_000, "pdf": 5_000}`; `class ExportacaoGrande(Exception)` com `detail`; `checar_limite(total: int, formato: str) -> None` (mensagem da tabela "Decisões").
  - `@dataclass Planilha(titulo: str, colunas_excel: list[Coluna], colunas_pdf: list[Coluna], linhas: list[dict], filtros: dict[str, str], nome_base: str, paisagem: bool = True)`; `gerar(planilha: Planilha, formato: Literal["xlsx", "pdf"]) -> Response`.
  - `ROTULO_ORDENACAO = {"id": "Patrimônio", "tipo": "Tipo", "equipamento": "Equipamento", "status": "Status", "colaborador": "Colaborador", "setor": "Setor", "cargo": "Cargo", "unidade": "Unidade", "cadastro": "Cadastro"}`; `data_hora_br(v: datetime | None) -> str | None` (`"dd/mm/aaaa HH:MM"`).
  - `planilha_estoque(inv, *, tipo, status, revenda, search, colaborador_id, sort, direcao, formato, colaborador_nome) -> Planilha` — `list_items(..., limit=LIMITE_LINHAS[formato] + 1, offset=0)` na MESMA ordem da tela; colunas Excel: `Patrimônio (id, numero)`, `Tipo`, `Marca`, `Modelo`, `Identificador`, `Status`, `Colaborador` (`colaborador_nome` ou `assigned_to`), `Setor`, `Cargo`, `Superior`, `Unidade` (`revenda`, maiúsculas), `Local`, `Cadastro` (data), `Garantia até` (data), `Valor de compra` (moeda); PDF: `Patrimônio`, `Tipo`, `Equipamento` (`brand model`), `Identificador`, `Status`, `Colaborador`, `Setor`, `Unidade`, `Local`. Filtros no cabeçalho: `Tipo`, `Status`, `Unidade` (maiúsculas), `Colaborador`, `Busca` (só os preenchidos) e sempre `Ordenação` = `f"{ROTULO_ORDENACAO[sort]} ({'crescente' if direcao == 'asc' else 'decrescente'})"`. Título `"Estoque de equipamentos"`, `nome_base="estoque"`. Sem CPF.
  - `planilha_historico(inv, *, search, item_id, formato) -> Planilha` — colunas `ID`, `Data` (`data_hora_br`), `Operação`, `Operador`, `Tipo`, `Marca`, `Modelo`, `Identificador`, `Usuário`, `Setor`, `Unidade` (maiúsculas), `Detalhes` (só Excel); sem CPF; título `"Histórico de operações"`, `nome_base="historico"`.
  - `planilha_manutencoes(db, *, status, item_id, revenda, tipo, de, ate, busca, formato) -> Planilha` — colunas `Nº` (chave `id`), `Aberta em`, `Equipamento` (`#id tipo marca modelo`), `Identificador`, `Unidade`, `Defeito`, `Assistência`, `Status` (`ROTULOS_STATUS`), `Previsão` (data), `Retorno` (`data_hora_br(retornada_em)`), `Custo` (moeda), `Custo acumulado do equipamento` (moeda, só Excel); última linha `{"id": "TOTAL", "custo": custo_total}` com a soma de todas as linhas filtradas; status `em_andamento` aparece no filtro como `"Em andamento"`. Título `"Manutenções"`, `nome_base="manutencoes"`.
  - `planilha_linhas(db, *, busca, status, unidade_id, unidade_nome, operadora, formato) -> Planilha` — colunas `Número` (`formatar_numero_linha`), `Operadora`, `Plano`, `Custo mensal` (moeda), `Status` (`Ativa`/`Suspensa`/`Cancelada`), `Unidade` (maiúsculas), `Aparelho` (`descrever_item` + `" (baixado)"` quando `item_ativo` é falso), `Colaborador` (+ `" (desligado)"`), `ICCID` (só Excel); última linha `{"numero": "Total mensal (linhas ativas)", "custo_mensal": soma das ativas exportadas}`. Título `"Linhas telefônicas"`, `nome_base="linhas"`.
- Produces (`app/routers/exportacoes.py`, `APIRouter(prefix="/api/exportacoes", tags=["exportacoes"])`, todos `gestor_or_tecnico`, `formato: Literal["xlsx", "pdf"] = Query(...)`, `ExportacaoGrande` → 400 com o `detail`):
  - `GET /estoque` (`tipo`, `status`, `revenda`, `search`, `colaborador_id`, `sort="id"`, `dir="desc"` com `alias="dir"`; ordenação inválida → 422 com a mensagem de `montar_order_by_items`).
  - `GET /historico` (`search`, `item_id`).
  - `GET /manutencoes` (mesmos filtros e tipos do `GET /api/manutencoes`, sem `limit`/`offset`).
  - `GET /linhas` (`busca`, `status`, `unidade_id`, `operadora`).

- [ ] **Step 1: Testes** (`tests/test_exportacoes.py`, integration; helper `_planilha(resp) -> list[list]` = `openpyxl.load_workbook(BytesIO(resp.content)).active.iter_rows(min_row=5, values_only=True)`; linha 3 da planilha tem os filtros):
  - `test_estoque_excel_respeita_filtros_e_ordem` (aceite): 3 itens Juazeiro (2 `Disponível`, 1 `Indisponível`) e 1 Petrolina `Disponível`; `GET /api/exportacoes/estoque?formato=xlsx&status=Disponível&revenda=Revalle Juazeiro&sort=id&dir=asc` → 200, `content-type` de xlsx, `content-disposition` contém `estoque_`; os ids da 1ª coluna == ids de `GET /api/items?status=Disponível&revenda=Revalle Juazeiro&sort=id&dir=asc` na mesma ordem; a coluna Unidade traz `"REVALLE JUAZEIRO"`; `A3` contém `"Status: Disponível"`, `"Unidade: REVALLE JUAZEIRO"` e `"Ordenação: Patrimônio (crescente)"`.
  - `test_estoque_ordenado_por_colaborador`: `sort=colaborador&dir=desc` → mesma ordem da API de listagem.
  - `test_estoque_pdf`: `formato=pdf` → `resp.content[:4] == b"%PDF"`.
  - `test_estoque_sem_cpf`: item emprestado com CPF `52998224725` → nenhuma célula contém `"52998224725"` nem `"529.982.247-25"`.
  - `test_formato_invalido_422` (`formato=csv`) e `test_ordenacao_invalida_422` (`sort=senha`) → `detail == "Ordenação inválida."`.
  - `test_pdf_acima_do_limite_400` (Review Focus): `monkeypatch.setitem(exportacoes_telas.LIMITE_LINHAS, "pdf", 2)` com 3 itens → 400, `detail == "A exportação em PDF está limitada a 2 linhas e esta consulta tem 3. Refine os filtros ou exporte em Excel."`.
  - `test_historico_excel_com_busca`: duas operações de itens diferentes; `search=<identificador de um>` → só as linhas desse item; `A3` contém `"Busca: <identificador>"`.
  - `test_manutencoes_excel_com_total`: mesmos dados de `test_custo_acumulado_por_equipamento` (F4: A com `150.00` e `89.90`, B com `1000.00`); `item_id=A` → 2 linhas + `TOTAL` com custo `239.9`; sem filtro → `TOTAL` `1239.9`.
  - `test_linhas_excel_com_filtro_de_status`: uma ativa (`49.90`) e uma suspensa → `status=ativa` → 1 linha + total `49.9`; aparelho baixado mostra `"(baixado)"`.
  - `test_exportacoes_exigem_gestor_ou_tecnico`: `client_aprendiz` nas quatro rotas → 403 `"Acesso restrito a Gestor ou Técnico."`.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_exportacoes.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `RBT tests/test_exportacoes.py tests/test_exports.py -q` → PASS.
- [ ] **Step 5: Commit** `feat(exportacoes): Excel e PDF do estoque, histórico, manutenções e linhas com os filtros da tela`.

### Task 6: Exportações de Termos pendentes e Relatório mensal; conferir Inventário

**Files:**
- Create: `backend/tests/test_exportacoes_termos_relatorio.py`
- Modify: `backend/app/services/exportacoes_telas.py`, `backend/app/routers/exportacoes.py`

**Interfaces:**
- Consumes: `PainelDBManager.termos_pendentes(busca=, limit=)` (Task 4); `ROTULO_STATUS_TERMO` (Task 4), `ROTULO_TIPO` (F3); `InventoryDBManager.generate_monthly_report(ano, mes)`; `_mascarar_cpf` não é público — usar uma função local `mascarar_cpf(v) -> str | None` em `exportacoes_telas.py` com a mesma regra (`***.456.789-**`; formato inesperado → `"***"`).
- Produces:
  - `planilha_termos_pendentes(db, *, busca, formato) -> Planilha` — colunas `Solicitação` (`termo_id` ou `"—"`), `Equipamento` (`#id tipo marca modelo`), `Identificador`, `Unidade` (maiúsculas), `Colaborador`, `Termo` (`"Entrega"`/`"Devolução"`/`"Uso"`), `Status` (`ROTULO_STATUS_TERMO`), `Solicitado por`, `Solicitado em`, `Enviado em`, `Agendado para` (os três com `data_hora_br`), `Dias pendente` (número, `(agora - referencia).days`), `Motivo` (só Excel). Sem e-mail e sem CPF. Título `"Termos pendentes de assinatura"`, `nome_base="termos_pendentes"`.
  - `planilha_relatorio_mensal(inv, *, ano, mes, formato) -> Planilha` — mesmas chaves e cabeçalhos do `HEADERS_PT` de `routers/reports.py` (copiar o dicionário para `exportacoes_telas.py`; não editar `reports.py`), com `cpf` passando por `mascarar_cpf` e `revenda` maiúscula; datas com `data_hora_br`; PDF só com `Operação`, `Data Empréstimo`, `Data Devolução`, `Usuário`, `Cargo`, `Setor`, `Centro de Custo`, `Revenda`, `Tipo`, `Marca`, `Modelo`, `Identificador`, `Operador`. Título `f"Relatório mensal {mes:02d}/{ano}"`, filtro `{"Período": f"{mes:02d}/{ano}"}`, `nome_base=f"relatorio_mensal_{ano}_{mes:02d}"`.
  - `GET /api/exportacoes/termos-pendentes` (`gestor_or_tecnico`; `busca`, `formato`).
  - `GET /api/exportacoes/relatorio-mensal` (`gestor_only`; `year: int | None`, `month: int | None` com padrão no mês atual, `month` fora de 1–12 → 400 `"Mês inválido."`, `formato`). O `GET /api/reports/monthly/export` (CSV) não muda.

- [ ] **Step 1: Testes** (`tests/test_exportacoes_termos_relatorio.py`, integration):
  - `test_termos_pendentes_excel`: solicitações `manual` e `enviado` (com `destinatario_email="maria.souza@gmail.com"`), uma `assinado` e um legado `item_pendente` → 3 linhas; nenhuma célula contém `"@"`; a do legado tem `Solicitação == "—"` e `Status == "Sem solicitação"`; ordenadas da mais antiga para a mais nova.
  - `test_termos_pendentes_busca`: `busca=<modelo de um item>` → só as linhas dele; `A3` contém `"Busca: "`.
  - `test_termos_pendentes_pdf` (`%PDF`) e `test_termos_pendentes_aprendiz_403`.
  - `test_relatorio_mensal_excel_mascara_cpf`: empréstimo com CPF `52998224725` no mês atual → a coluna `CPF` traz `"***.982.247-**"`, nenhuma célula com `"52998224725"`; revenda em maiúsculas; o CSV (`GET /api/reports/monthly/export`) continua trazendo o CPF como hoje (comportamento inalterado).
  - `test_relatorio_mensal_pdf` (`%PDF`); `test_relatorio_mensal_mes_invalido_400` (`month=13`); `test_relatorio_mensal_so_gestor` (técnico → 403 `"Acesso restrito ao Gestor."`).
  - Inventário (F6, só conferência): rodar `tests/test_inventario_relatorio.py -k exportar` no Step 4 — nada a implementar.
- [ ] **Step 2: Rodar e ver falhar:** `RBT tests/test_exportacoes_termos_relatorio.py -q` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `RBT tests/test_exportacoes_termos_relatorio.py tests/test_monthly_report.py -q` e `RBT tests/test_inventario_relatorio.py -k exportar -q` → PASS; depois a suíte inteira `RBT -q` → verde.
- [ ] **Step 5: Commit** `feat(exportacoes): termos pendentes e relatório mensal em Excel e PDF`.

### Task 7: Frontend — clientes do painel e das exportações e o componente de botões

**Files:**
- Create: `frontend/src/api/painel.ts`, `frontend/src/features/painel/alertas.ts`, `frontend/src/features/exportacoes/api.ts`, `frontend/src/features/exportacoes/api.test.ts`, `frontend/src/components/exportacao/BotoesExportar.tsx`, `frontend/src/components/exportacao/BotoesExportar.test.tsx`

**Interfaces:**
- Consumes: contratos das Tasks 3–6; `downloadAuthenticated` (`api/client.ts`); `getErrorMessage`; `toast`; `Button` (`loading`); `EstoqueSort`, `SortDir` (`features/estoque/api.ts`); `ListManutencoesParams` (`api/manutencoes.ts`).
- Produces:
  - `api/painel.ts`: tipos espelhando os schemas (`ResumoPainel`, `AlertasPainel`, `Alerta`, `TipoAlerta`, `SeveridadeAlerta = 'alta' | 'media' | 'baixa'`; decimais como `string`, datas como `string`); `getResumoPainel(): Promise<ResumoPainel>` (`GET /painel/resumo`); `getAlertasPainel(): Promise<AlertasPainel>` (`GET /painel/alertas`).
  - `features/painel/alertas.ts`: `ORDEM_TIPOS_ALERTA: TipoAlerta[] = ['colaborador_desligado', 'devolucao_pendente', 'termo_pendente', 'manutencao_parada', 'linha_aparelho_baixado', 'garantia_vencendo']`; `ROTULO_TIPO_ALERTA = { colaborador_desligado: 'Colaborador desligado com equipamento', devolucao_pendente: 'Devolução pendente', termo_pendente: 'Termo pendente', manutencao_parada: 'Manutenção parada', linha_aparelho_baixado: 'Linha ativa em aparelho baixado', garantia_vencendo: 'Garantia vencendo' }`; `ROTULO_SEVERIDADE = { alta: 'Alta', media: 'Média', baixa: 'Baixa' }`; `TOM_SEVERIDADE: Record<SeveridadeAlerta, StatusTone> = { alta: 'danger', media: 'pending', baixa: 'neutral' }`; `agruparAlertas(alertas: Alerta[]): { tipo: TipoAlerta; alertas: Alerta[] }[]` (na ordem de `ORDEM_TIPOS_ALERTA`, só grupos não vazios).
  - `features/exportacoes/api.ts`: `type FormatoExportacao = 'xlsx' | 'pdf'`; `montarQuery(params: Record<string, string | number | null | undefined>): string` (pula `undefined`, `null` e `''`; `URLSearchParams`); `exportar(caminho: string, params: Record<string, string | number | null | undefined>, nomeBase: string, formato: FormatoExportacao): Promise<void>` → `downloadAuthenticated(`/exportacoes/${caminho}?${montarQuery({ ...params, formato })}`, `${nomeBase}.${formato}`)`; erro do axios cujo `response.data` é `Blob` → lê o texto, faz `JSON.parse` e regrava `error.response.data` antes de relançar (assim `getErrorMessage` mostra o `detail`); wrappers `exportarEstoque(p: { tipo?; status?; revenda?; search?; colaborador_id?; sort: EstoqueSort; dir: SortDir }, f)`, `exportarHistorico(p: { search? }, f)`, `exportarManutencoes(p: Omit<ListManutencoesParams, 'limit' | 'offset'>, f)`, `exportarLinhas(p: { busca?; status?; unidade_id?; operadora? }, f)`, `exportarTermosPendentes(p: { busca? }, f)`, `exportarRelatorioMensal(year: number, month: number, f)` com `nomeBase` `estoque`, `historico`, `manutencoes`, `linhas`, `termos_pendentes`, `` `relatorio_mensal_${year}_${String(month).padStart(2, '0')}` ``.
  - `BotoesExportar({ onExportar, disabled = false, rotulo = 'Exportar' }: { onExportar: (formato: FormatoExportacao) => Promise<void>; disabled?: boolean; rotulo?: string })` — `role="group"` com `aria-label={rotulo}`; dois `Button size="sm" variant="outline"` com textos `Excel` e `PDF`, `aria-label` `` `${rotulo} em Excel` `` / `` `${rotulo} em PDF` ``, ícone `Download` e `loading` independente por formato (os dois desabilitados enquanto um exporta); erro → `toast(getErrorMessage(err, 'Não foi possível exportar.'), 'error')`.

- [ ] **Step 1: Testes:**
  - `api.test.ts`: `montarQuery({ a: 'x', b: undefined, c: '', d: 0 })` → `'a=x&d=0'`; `exportarEstoque({ status: 'Disponível', sort: 'colaborador', dir: 'asc' }, 'xlsx')` faz `GET /api/exportacoes/estoque` com `status=Disponível`, `sort=colaborador`, `dir=asc`, `formato=xlsx` (MSW devolve um blob; `URL.createObjectURL`/`revokeObjectURL` com `vi.fn`); MSW devolvendo 400 `{ detail: 'A exportação em PDF está limitada a 5.000 linhas e esta consulta tem 7.000. Refine os filtros ou exporte em Excel.' }` → a promessa rejeita e `getErrorMessage(err)` devolve exatamente esse texto (Review Focus).
  - `BotoesExportar.test.tsx`: clicar `Exportar em PDF` chama `onExportar('pdf')` e mostra `aria-busy` no botão PDF até resolver; `onExportar` rejeitando com `new Error('Falhou')` → toast `Falhou`; `disabled` desabilita os dois.
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/features/exportacoes src/components/exportacao` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run src/features/exportacoes src/components/exportacao` e `npx tsc --noEmit` → verdes.
- [ ] **Step 5: Commit** `feat(exportacoes): cliente de exportação e botões Excel/PDF`.

### Task 8: Frontend — `DashboardPage` redesenhada

**Files:**
- Create: `frontend/src/features/painel/IndicadoresPainel.tsx`, `frontend/src/features/painel/EquipamentosPorUnidade.tsx`, `frontend/src/features/painel/ListaAlertas.tsx`, `frontend/src/pages/DashboardPage.test.tsx`
- Modify: `frontend/src/pages/DashboardPage.tsx` (reescrita; dona: F7)

**Interfaces:**
- Consumes: Task 7; `StatBlock`, `StatusPill`, `Button`, `formatMoeda`, `upper`, `formatDateTime`; `getLoansChart` (gráfico mantido); `useAuth().hasRole`; `useTheme` (mantido para o gráfico).
- Produces:
  - `IndicadoresPainel({ resumo }: { resumo: ResumoPainel })` — `grid grid-cols-2 gap-4 sm:grid-cols-4 xl:grid-cols-7` com sete `StatBlock` (rótulos exatos): `Equipamentos` (`total`, hint `ativos cadastrados`); `Disponíveis` (hint `` `${pct}% do total` ``, `pct` arredondado, 0 sem itens); `Em uso` (`Indisponível`); `Pendentes` (`Pendente + Pendente Devolução`, hint `` `${p} empréstimo(s) · ${pd} devolução(ões)` ``); `Em manutenção` (hint `` `${ordens} ordem(ns) em andamento` ``); `Termos pendentes` (hint `` `+${n} sem solicitação` `` quando `sem_solicitacao > 0`, senão `aguardando assinatura`); `Linhas ativas` (hint `` `${formatMoeda(custo_mensal)}/mês` ``).
  - `EquipamentosPorUnidade({ unidades }: { unidades: EquipamentosPorUnidade[] })` — tabela dentro de `overflow-x-auto` (390 px rola sem quebrar a página): `Unidade` (`upper(unidade)` ou `SEM UNIDADE`), `Total` e uma coluna por status na ordem `Disponível`, `Indisponível`, `Pendente`, `Pendente Devolução`, `Em Manutenção`, com o cabeçalho em `StatusPill`; vazio → `Nenhum equipamento cadastrado.`
  - `ListaAlertas({ dados }: { dados: AlertasPainel })` — título `Alertas` com o total; vazio → `Nenhum alerta no momento.`; grupos de `agruparAlertas` com `ROTULO_TIPO_ALERTA` e a contagem; cada alerta é um `Link` (`react-router`) para `alerta.link` com `titulo` em negrito, `descricao` e `StatusPill status={ROTULO_SEVERIDADE[s]} tone={TOM_SEVERIDADE[s]}`; até 5 por grupo e o botão `` `Mostrar todos (${n})` `` / `Mostrar menos`; rodapé `` `Limites: garantia em até ${g} dias, termo há mais de ${t} dias, devolução há mais de ${d} dias, manutenção parada há mais de ${m} dias.` ``.
  - `DashboardPage`: cabeçalho `Painel` (eyebrow `SISTEMA DE GESTÃO PATRIMONIAL`) com `` `Atualizado em ${formatDateTime(resumo.gerado_em)}` `` e botão `Atualizar` (`loading` enquanto refaz as consultas `['painel']`); `IndicadoresPainel`; grade `lg:grid-cols-12`: `ListaAlertas` (`lg:col-span-7`, só `hasRole('Gestor', 'Técnico')`; a consulta `['painel', 'alertas']` tem `enabled` só para esses papéis) + `EquipamentosPorUnidade` (`lg:col-span-5`, ou 12 sem alertas); abaixo, o gráfico de movimentação atual e os "Fluxos Rápidos" atuais (acrescentar `{ to: '/manutencoes', label: 'Manutenções', description: 'Ordens abertas e custo', icon: Wrench, roles: ['Gestor', 'Técnico'] }`). Saem: as cinco consultas `listItemsPaginated`, "Atenção Operacional", "Últimas Movimentações" e o ponto `bg-emerald-500`. Estados: carregando (`Carregando indicadores...`), erro com `Tentar novamente`.

- [ ] **Step 1: Testes** (`DashboardPage.test.tsx`; `vi.mock('@/contexts/AuthContext')` como em `ItemDetailPage.test.tsx`, com o papel variável por teste; MSW para `/api/painel/resumo`, `/api/painel/alertas`, `/api/reports/charts/loans`; `MemoryRouter`; como o MSW recusa requisição não mockada, qualquer chamada a `/api/items` ou `/api/history` falha o teste):
  - `mostra os indicadores do resumo`: resumo com `total: 812`, `Disponível: 400`, `Indisponível: 300`, `Pendente: 10`, `Pendente Devolução: 2`, `Em Manutenção: 100`, `ordens_em_andamento: 98`, `termos.pendentes: 7`, `sem_solicitacao: 3`, `linhas.ativas: 40`, `custo_mensal: '1996.00'` → textos `812`, `49% do total`, `12`, `10 empréstimo(s) · 2 devolução(ões)`, `98 ordem(ns) em andamento`, `+3 sem solicitação`, `R$ 1.996,00/mês` (números acima de 500 provam que não há limite no cliente).
  - `lista os alertas agrupados com link para o registro`: dois alertas (`colaborador_desligado` com `link: '/items/9'`, `garantia_vencendo`) → cabeçalhos `Colaborador desligado com equipamento` antes de `Garantia vencendo`; o link do primeiro tem `href="/items/9"`; pílula `Alta`.
  - `mostra até 5 por grupo e expande`: 7 alertas `termo_pendente` → 5 visíveis e `Mostrar todos (7)`; clicar → 7.
  - `sem alertas mostra a mensagem` → `Nenhum alerta no momento.`
  - `Jovem Aprendiz vê só os indicadores` (Review Focus): papel `Jovem Aprendiz` → indicadores visíveis, nenhum texto `Alertas`, e o handler de `/api/painel/alertas` não foi chamado (contador no handler == 0).
  - `unidade sem nome aparece como SEM UNIDADE` e nomes em caixa alta (`REVALLE JUAZEIRO`).
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/DashboardPage.test.tsx` → FAIL.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run`, `npx tsc --noEmit` → verdes; conferir no navegador (`npm run dev` com o backend local) em 390, 768 e 1440 px que nada estoura a largura (o teste-guarda de `text-primary` continua verde).
- [ ] **Step 5: Commit** `feat(painel): painel inicial com indicadores do servidor e lista de alertas`.

### Task 9: Frontend — botões de exportação nas telas

**Files:**
- Modify (só acréscimos: import + botão): `frontend/src/pages/StockPage.tsx`, `HistoryPage.tsx`, `ManutencoesPage.tsx`, `LinhasPage.tsx`, `TermsPage.tsx`, `ReportPage.tsx`
- Test (acrescentar um teste em cada): `StockPage.test.tsx`, `HistoryPage.test.tsx`, `ManutencoesPage.test.tsx`, `LinhasPage.test.tsx`, `TermsPage.test.tsx`, `ReportPage.test.tsx`

**Interfaces:**
- Consumes: `BotoesExportar` e os wrappers de `features/exportacoes/api.ts` (Task 7).
- Produces (o botão recebe exatamente os filtros que a consulta da tela já usa, nos mesmos nomes):
  - `StockPage`: em `PageHeader.actions`, antes de `Novo Ativo`, `{hasRole('Gestor', 'Técnico') && <BotoesExportar disabled={total === 0} onExportar={(f) => exportarEstoque({ tipo, status, revenda, colaborador_id, search: busca, sort: ordem.sort, dir: ordem.dir }, f)} />}` (os mesmos `!== 'all'`/`?? undefined` da `queryFn`).
  - `HistoryPage`: `actions` passa a ser um fragmento com o botão `Exportar CSV` atual (inalterado) + `<BotoesExportar disabled={(data?.total ?? 0) === 0} onExportar={(f) => exportarHistorico({ search: busca || undefined }, f)} />`.
  - `ManutencoesPage`: em `actions`, antes de `Nova manutenção`, `BotoesExportar` com `{ status, item_id: itemId, revenda, tipo, de, ate, busca: buscaAplicada }` (vazios como `undefined`), `disabled={periodoInvalido || total === 0}`.
  - `LinhasPage`: dentro do fragmento de `podeEscrever`, antes de `Relatório de custo`, `BotoesExportar` com `{ busca, status, unidade_id: unidadeId ? Number(unidadeId) : undefined, operadora }`.
  - `TermsPage`: no bloco dos contadores (`Pendentes de Assinatura` / `Vigentes / Assinados`), depois deles, `<BotoesExportar rotulo="Exportar termos pendentes" onExportar={(f) => exportarTermosPendentes({ busca: searchAplicado || undefined }, f)} />` (a rota já é só Gestor/Técnico).
  - `ReportPage`: ao lado do botão `CSV`, `{hasRole('Gestor') && <BotoesExportar disabled={report.length === 0} onExportar={(f) => exportarRelatorioMensal(queryParams.year, queryParams.month, f)} />}` (importar `useAuth`); o aviso abaixo continua igual (vale para os três formatos).

- [ ] **Step 1: Testes** (um por tela, MSW com o handler da rota de exportação registrando a URL; `URL.createObjectURL`/`revokeObjectURL` com `vi.fn`):
  - Estoque: filtrar por status `Disponível`, ordenar por `Colaborador`, clicar `Exportar em Excel` → URL com `status=Disponível`, `sort=colaborador`, `dir=asc`, `formato=xlsx`; com papel `Jovem Aprendiz` não há grupo `Exportar` (Review Focus).
  - Histórico: digitar a busca `SN-1`, esperar o debounce, `Exportar em PDF` → `search=SN-1&formato=pdf`; o botão `Exportar CSV` continua lá.
  - Manutenções: status `Em andamento` → `status=em_andamento`; com `?item_id=7` → `item_id=7`.
  - Linhas: status `Ativa` → `status=ativa`.
  - Termos: `Exportar termos pendentes em Excel` → `GET /api/exportacoes/termos-pendentes?formato=xlsx`.
  - Relatório: Gestor vê `Exportar em PDF` e a URL leva `year`/`month` gerados; Técnico não vê o grupo (o `CSV` continua).
- [ ] **Step 2: Rodar e ver falhar:** `npx vitest run src/pages/StockPage.test.tsx src/pages/HistoryPage.test.tsx src/pages/ManutencoesPage.test.tsx src/pages/LinhasPage.test.tsx src/pages/TermsPage.test.tsx src/pages/ReportPage.test.tsx` → FAIL nos testes novos.
- [ ] **Step 3: Implementar.**
- [ ] **Step 4: Rodar** `npx vitest run`, `npx tsc --noEmit`, `npm run build` → verdes (os testes antigos dessas telas passam sem alteração).
- [ ] **Step 5: Commit** `feat(exportacoes): botões Excel e PDF no estoque, histórico, manutenções, linhas, termos e relatório`.

### Task 10: README de operação e deploy

**Files:**
- Modify: `README.md` (raiz). `backend/.env.example` não muda (conferir apenas).

**Interfaces:**
- Consumes: `backend/.env.example` (F0–F3), `backend/migrations/README.md`, `scripts/migrate.py`, `scripts/reset_2fa.py` (F2), plano mestre.
- Produces — seções novas/atualizadas no `README.md` (português, sem valor real de segredo):
  - **"Variáveis de ambiente novas"**: tabela com cada variável acrescentada nesta rodada e para que serve: `APP_PUBLIC_URL`; `MAIL_BACKEND`, `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM`, `SMTP_VERIFY_TLS`; `COLABORADORES_DATABASE_URL`, `COLABORADORES_TABLE`, `COLABORADORES_SCHEMA`, `COLABORADORES_COLUMNS`, `COLABORADORES_SYNC_HORA`; `SCHEDULER_ENABLED`; `TOTP_ENCRYPTION_KEY` (gerar com Fernet; vazia = 2FA desligado; trocar a chave invalida os autenticadores cadastrados; com a chave definida, todo Gestor é levado a cadastrar o 2FA no primeiro login); `REVSIGN_ENABLED` (padrão `False` — fica desligado até o contrato do RevSign chegar; termos seguem o fluxo manual), `REVSIGN_BASE_URL`, `REVSIGN_API_TOKEN`, `REVSIGN_WEBHOOK_SECRET` (URL do webhook `<APP_PUBLIC_URL>/api/integrations/revsign/webhook`); `ALERTA_GARANTIA_DIAS` (30), `ALERTA_TERMO_PENDENTE_DIAS` (3), `ALERTA_DEVOLUCAO_PENDENTE_DIAS` (7), `ALERTA_MANUTENCAO_DIAS` (15).
  - **"Deploy no EasyPanel"**: o container do backend precisa estar na mesma rede do EasyPanel em que o host `agente_revalle_postgres` resolve (base de RH, só leitura; usar usuário de leitura); a câmera do inventário exige HTTPS (o proxy com TLS já atende; conferir antes de usar no celular); SMTP `mail.revalle.com.br:465` com a senha só no `.env` do servidor; recomendação de `TZ=America/Sao_Paulo` no container do backend (os termos gravam em horário de Brasília; as demais tabelas usam o relógio do servidor).
  - **"Migrações"**: rodam sozinhas na subida do backend (com trava para os 2 workers) e por linha de comando dentro do container: `python -m scripts.migrate` (aplica), `python -m scripts.migrate pendentes` (lista), `python -m scripts.migrate reverter` (reverte a última); exemplo `docker compose exec backend python -m scripts.migrate pendentes`.
  - **"2FA: Gestor sem acesso"**: `docker compose exec backend python -m scripts.reset_2fa <usuario>` zera o 2FA do usuário e grava na auditoria.
  - **"Antes de publicar (checklist)"**, em caixas de seleção: (1) backup completo do MySQL (`mysqldump --single-transaction --routines`); (2) rodar as migrações numa CÓPIA restaurada do backup (`python -m scripts.migrate` apontando para a cópia) e conferir o sistema nela; (3) `SHOW CREATE TABLE items` em produção: a coluna `status` deve ser `enum('Disponível','Indisponível','Pendente','Pendente Devolução') DEFAULT 'Disponível'` — a 0402 redefine a coluna com essa base mais `'Em Manutenção'`; se a definição for diferente (ex.: `NOT NULL` ou outro padrão), ajuste a 0402 antes do deploy; (4) agendar janela de manutenção: 0004 e 0401 adicionam colunas com chave estrangeira e reconstroem a tabela `items` (bloqueia escrita durante a cópia); (5) preencher o `.env` com as variáveis novas; (6) confirmar a rede do EasyPanel e o HTTPS; (7) depois da subida: `python -m scripts.migrate pendentes` vazio, sincronizar colaboradores pela tela e conferir o painel.
  - Seção "Endpoints da API": acrescentar `/api/painel/resumo`, `/api/painel/alertas`, `/api/exportacoes/*` (uma linha cada) — só acréscimo.

- [ ] **Step 1: Conferência que falha antes:** `cd $WT && for v in $(grep -oE '^[A-Z_]+=' backend/.env.example | tr -d '=' | grep -E 'APP_PUBLIC|MAIL_|SMTP_|COLABORADORES_|SCHEDULER|TOTP|REVSIGN|ALERTA_'); do grep -q "$v" README.md || echo "faltando: $v"; done` → lista variáveis faltando.
- [ ] **Step 2: Escrever as seções.**
- [ ] **Step 3: Rodar a conferência de novo** → nenhuma linha `faltando:`; `grep -n "reset_2fa\|scripts.migrate\|agente_revalle_postgres\|HTTPS\|SHOW CREATE TABLE items\|0004\|0401" README.md` mostra todas as ocorrências.
- [ ] **Step 4: Commit** `docs(readme): variáveis novas, deploy no EasyPanel, migrações, 2FA e checklist antes de publicar`.

### Task 11: Verificação final da F7 e da matriz inteira

**Files:** nenhum novo (corrigir na tarefa dona o que a verificação apontar).

- [ ] **Step 1: Backend inteiro:** `RBT -q` → tudo verde (os marcados `mudanca_esperada` seguem as regras de `docs/TESTES.md`).
- [ ] **Step 2: Migrações do zero:** `bash C:/Users/enzo.jz/dev/patrimonio/_tools/run-backend-tests.sh $WT f7zero tests/test_migrations.py tests/test_migracoes_f4.py -q` → verde; depois `docker exec stock_sys_mysql_test mysql -uroot -pstock_test_root_pw -e "DROP DATABASE stock_sys_test_f7zero"`. Confirmar que não existe arquivo `backend/migrations/07*.sql`.
- [ ] **Step 3: Frontend inteiro:** em `$WT/frontend`: `npx vitest run`, `npx tsc --noEmit`, `npm run build` → verdes; `grep -rn "text-primary\|bg-emerald" src/pages/DashboardPage.tsx src/features/painel src/components/exportacao` vazio.
- [ ] **Step 4: Spec F7 item a item** (anotar no relatório o teste que prova cada um):
  - [ ] F7.1 resumo no servidor, por unidade e status (incl. `Em Manutenção`), manutenção, termos pendentes, linhas ativas e custo → Task 3 (`test_resumo_bate_com_consultas_diretas`).
  - [ ] F7.2 alertas calculados na hora, os cinco do spec + linha em aparelho baixado, cada um com link, somem quando resolvidos → Task 4.
  - [ ] F7.3 painel redesenhado → Task 8.
  - [ ] F7.4 exportação Excel/PDF: Estoque com filtros e ordenação (Task 5), Histórico (5), Manutenções (5), Linhas (5), Inventário (F6, conferido na Task 6), Termos pendentes (6), Relatório mensal com o CSV mantido (6); botões nas telas (Task 9).
  - [ ] Ruling 17 (fontes da linha do tempo) → Task 2; Ruling 23 (empréstimo muda unidade) → Task 1; README → Task 10.
- [ ] **Step 5: Matriz da seção 3 inteira** — para cada linha "entra", localizar o teste com `grep -rln "<nome do teste>" backend/tests frontend/src` e registrar fase e evidência no relatório:

  | Item | Fase | Evidência a conferir |
  |---|---|---|
  | 1.1 Esqueci minha senha | F2 | `tests/test_redefinir_senha.py` |
  | 1.2 2FA | F2 | `tests/test_mfa_login.py`, `test_mfa_cadastro_gestor.py`, `test_zerar_2fa.py` |
  | 2.1 / 2.2 paginação 25/50/100 | F0+F1 | `Pagination.test.tsx`, `src/test/listagens.guard.test.ts` |
  | 2.3 / 7.9 status em pílula e negrito | F0+F1 | `StatusPill.test.tsx` |
  | 2.4 Setor e Cargo no estoque | F1 | `StockPage.test.tsx`, `tests/test_items_ordenacao.py` |
  | 3.1 / 3.5 RevSign e termo por e-mail | F3 | `tests/test_revsign_webhook.py::test_webhook_assinado_conclui_emprestimo_e_grava_pdf` |
  | 3.4 telefone/e-mail validados e exigidos | F0+F3 | `tests/test_colaboradores.py`, `tests/test_loan_colaborador.py::test_colaborador_sem_email_400` |
  | 4.1 / 4.2 manutenção com custo | F4 | `tests/test_manutencoes.py::test_custo_acumulado_por_equipamento` |
  | 6.1 / 6.2 / 6.4 / 6.5 / 6.6 termos | F3 | `test_termos_api.py`, `test_termos_jobs.py::test_envio_agendado_sai_no_tique_certo`, `test_loan_colaborador.py::test_devolucao_cancela_uso_pendente` |
  | 7.1 / 7.5 / 7.6 / 7.7 visual, responsivo, paleta | F0+F1 | `tokens.test.ts`, conferência em 390/768/1440 px |
  | 7.2 painel com indicadores | F7 | Tasks 3 e 8 |
  | 7.3 busca, filtros e ordenação | F1 | `tests/test_items_ordenacao.py`, `SortableHeader.test.tsx` |
  | 7.4 detalhe com histórico completo | F0+F1/3/4/5/8+F7 | `ItemDetailPage.test.tsx`, `tests/test_linha_do_tempo.py`, `tests/test_linha_do_tempo_f7.py` |
  | 7.8 carregamento nos botões | F0+F1 | `button.test.tsx` |
  | 7.10 unidade/revenda em caixa alta | F1+F7 | `test_export_csv_revenda_em_caixa_alta`, `test_estoque_excel_respeita_filtros_e_ordem` |
  | 7.11 ir para a página | F0 | `Pagination.test.tsx` |
  | 8.1 / 8.2 / 8.5 / 8.7 CPF, base de RH, cadastro | F0 | `tests/test_colaboradores.py`, `tests/test_colaboradores_source.py` |
  | 8.4 localização | F0+F4 | `tests/test_locais.py`, `tests/test_items_compra.py::test_criar_com_local_da_unidade` |
  | 8.6 superior imediato | F0+F1 | `tests/test_colaboradores.py`, aba Dados |
  | 10.1 / 10.2 NF, valor, garantia, anexo | F4 | `tests/test_items_compra.py` |
  | 10.3 transferência | F4+F7 | `tests/test_transferencias.py`, `tests/test_emprestimo_transferencia.py` |
  | 10.4 linha do tempo | F4+F7 | `tests/test_linha_do_tempo.py`, `tests/test_linha_do_tempo_f7.py` |
  | 10.5 baixa com apagamento | F4 | `tests/test_baixa_apagamento.py` |
  | 11.1–11.4 etiquetas e inventário | F6 | `tests/test_etiquetas.py`, `tests/test_inventario_relatorio.py` |
  | 12.2 log de alterações | F0 | `tests/test_audit.py`, `tests/test_audit_fluxos.py` |
  | 12.3 alertas | F7 | `tests/test_painel_alertas.py` |
  | 12.4 Excel e PDF | F0+F6+F7 | `tests/test_exports.py`, `tests/test_exportacoes.py`, `tests/test_exportacoes_termos_relatorio.py` |
  | 13.1–13.4 linhas | F5 | `tests/test_linhas*.py` |
  | 14.1–14.12 acessos | F8 | `tests/test_acesso*.py` |

  Itens "Fora" (3.2, 3.3, 3.6, 5.x, 6.3, 8.3, 12.1 por unidade, 14.13) continuam fora: confirmar que nada foi implementado para eles. Item sem evidência = pendência no relatório (não inventar teste nesta etapa).
- [ ] **Step 6: Limpeza:** nenhum container, imagem ou database de teste deixado (o script usa `docker run --rm`; databases `stock_sys_test_f7*` removidos); listar no relatório o que foi removido.
- [ ] **Step 7: Relatório da fase** (mensagem ao coordenador, sem arquivo novo): linhas acrescentadas em arquivos compartilhados (`main.py`, `dependencies.py`, `linha_do_tempo.py`), a mudança no `issue()` (Ruling 23, inclusive `local_id=NULL`), as telas de outras fases que ganharam botão, a decisão de permissões (alertas sem Jovem Aprendiz; Excel/PDF do relatório mensal só Gestor e com CPF mascarado, CSV inalterado), e o risco de fuso abaixo.
- [ ] **Step 8: Commit** (só se houve correção) `fix(painel): ajustes da verificação final da F7`.

---

## Notas e riscos

- **Fuso:** `termo_solicitacoes` grava em horário de Brasília (`agora_local`, F3) e as outras tabelas em `datetime.now()` do servidor. Se o container rodar em UTC, a linha do tempo pode mostrar eventos de termo com 3 h de deslocamento em relação aos demais; os alertas não são afetados (cada fonte compara com o seu próprio relógio). Recomendação ao Enzo: `TZ=America/Sao_Paulo` no container do backend (anotar no README, seção de deploy, como recomendação).
- `issue()` é da F3; a mudança é a do Ruling 23 e mais nada. Se a F3 mudar a assinatura depois, a Task 1 segue o `issue` vigente.
- Exportar o histórico inteiro pode passar de 50.000 linhas em produção: a mensagem orienta a filtrar; o CSV da página atual continua disponível.
