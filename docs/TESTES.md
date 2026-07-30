# Testes de caracterização (Módulos 7a/7b)

Esta suíte **não testa (só) o comportamento desejado** do sistema — ela nasceu (Módulo
7a) como rede de segurança para uma refatoração ampla do backend, testando o
comportamento **atual** de propósito. A refatoração já foi mesclada. Este documento
cobre também o trabalho do **Módulo 7b**: a triagem das falhas que a suíte passou a
acusar depois do merge (separando mudança intencional de regressão), a correção de
alguns bugs de teste encontrados no caminho, e a cobertura de comportamento novo que
não existia quando a suíte original foi escrita.

> ## 🚨 Regra absoluta: nunca escrever no MySQL de produção
>
> O host `72.61.53.20` / banco `stock_sys_db` é **produção**. Uma versão anterior
> desta suíte já gravou lá por acidente (criou usuários indevidos) por causa de um
> singleton (`app.core.config.settings`) construído antes da hora. A suíte se protege
> com **duas camadas**, ambas em `backend/tests/conftest.py`:
>
> 1. **`pytest_configure(config)`** (hook do pytest, roda ANTES de qualquer coleta de
>    teste): lê `TEST_DB_HOST`/`TEST_DB_NAME` do ambiente e, se não baterem com os
>    marcadores conhecidos de produção, **sobrescreve `os.environ["DB_HOST"]` etc.
>    imediatamente** — antes que qualquer `import app.*` possa materializar
>    `Settings()` a partir do `backend/.env` real. Sem isso, um teste `unit` que
>    importa `app.core.config` no nível de módulo (antes de qualquer fixture rodar)
>    já bastaria para "vazar" a configuração de produção para o resto do processo.
> 2. **`_assegurar_app_apontando_para_teste(cfg)`**, chamada no fim da fixture
>    `_test_schema` (a primeira vez que qualquer coisa toca o banco de verdade):
>    confere se `app.core.config.settings.DB_HOST/DB_NAME` realmente é o banco de
>    teste. Se não for — sinal de que algum import aconteceu cedo demais e o
>    singleton já foi construído com outro valor — **aborta a sessão inteira**
>    (`pytest.exit`, saída 1) em vez de deixar a suíte gravar no lugar errado.
>
> Além disso, `TEST_DB_HOST`/`TEST_DB_NAME` iguais aos marcadores de produção
> (`72.61.53.20` / `stock_sys_db`, ou o que estiver num `backend/.env` real presente
> na máquina) também abortam a sessão (`pytest.exit`) antes mesmo de tentar conectar.
> **Nunca** rode esta suíte com credenciais de produção em `TEST_DB_*` — use sempre o
> MySQL efêmero de `docker-compose.test.yml`.

## Sumário

- [Como subir o banco de teste](#como-subir-o-banco-de-teste)
- [Variáveis de ambiente](#variáveis-de-ambiente)
- [Como rodar a suíte](#como-rodar-a-suíte)
- [Como interpretar uma falha depois de uma refatoração](#como-interpretar-uma-falha-depois-de-uma-refatoração)
- [Mudanças de comportamento deste ciclo (Módulo 7b)](#mudanças-de-comportamento-deste-ciclo-módulo-7b)
- [Bugs corrigidos com teste de regressão](#bugs-corrigidos-com-teste-de-regressão)
- [Bug de produto encontrado neste ciclo (reportado, não corrigido)](#bug-de-produto-encontrado-neste-ciclo-reportado-não-corrigido)
- [Cobertura nova adicionada neste ciclo (T2)](#cobertura-nova-adicionada-neste-ciclo-t2)
- [Lista de BUG CONHECIDO (fora de escopo desta rodada)](#lista-de-bug-conhecido-fora-de-escopo-desta-rodada)
- [Reclassificado: comportamento esperado, não bug](#reclassificado-comportamento-esperado-não-bug)

## Como subir o banco de teste

O banco de teste é um MySQL 8 efêmero (dados em `tmpfs`, tudo é perdido ao parar o
container), definido em `docker-compose.test.yml` na raiz do repositório — **não** é o
mesmo arquivo/serviço do `docker-compose.yml` de produção.

```bash
# Subir (porta padrão 3307, para não colidir com um MySQL local em 3306)
docker compose -f docker-compose.test.yml up -d

# Esperar ficar saudável
docker compose -f docker-compose.test.yml ps

# Derrubar (e descartar todos os dados)
docker compose -f docker-compose.test.yml down
```

Se a porta 3307 já estiver em uso, sobrescreva antes de subir:

```bash
TEST_DB_PORT=3308 docker compose -f docker-compose.test.yml up -d
```

## Variáveis de ambiente

A suíte **se recusa a rodar** sem essas variáveis explícitas — isso é proposital: o
`backend/.env` real da máquina de desenvolvimento aponta para o MySQL de **produção**
(`72.61.53.20` / `stock_sys_db`), e nenhuma fixture pode arriscar escrever lá.

| Variável            | Obrigatória | Default (se ausente)        | Observação |
|---------------------|:-----------:|------------------------------|------------|
| `TEST_DB_HOST`      | **sim**     | —                             | ex.: `127.0.0.1` |
| `TEST_DB_NAME`      | **sim**     | —                             | ex.: `stock_sys_test` |
| `TEST_DB_PORT`      | não         | `3307`                        | deve casar com o `docker-compose.test.yml` |
| `TEST_DB_USER`      | não         | `stock_test`                  | deve casar com `MYSQL_USER` do compose |
| `TEST_DB_PASSWORD`  | não         | `stock_test_pw`               | deve casar com `MYSQL_PASSWORD` do compose |
| `JWT_SECRET`        | **sim**     | —                             | não é específico da suíte: `app/core/config.py::Settings.JWT_SECRET` é obrigatório e sem default (mínimo 32 caracteres) para a aplicação inteira, mesmo fora de teste. Sem ele, a coleta falha na importação de `app.core.config` com `ValidationError: JWT_SECRET Field required`, mesmo antes do guard de banco entrar em ação. Gere um valor descartável com `python -c "import secrets; print(secrets.token_urlsafe(48))"` — qualquer string de 32+ caracteres serve para os testes. |

Comportamento do guard (`backend/tests/conftest.py::_ambiente_de_teste` e
`pytest_configure`, ver também o alerta destacado no topo deste documento):

- **`TEST_DB_HOST`/`TEST_DB_NAME` não definidos** → todos os testes de integração
  **pulam** (`pytest.skip`) com uma mensagem explicando o que configurar. Nada falha,
  nada dá erro — é o estado padrão deste repositório (CI sem Docker, `pytest
  --collect-only`, etc.).
- **`TEST_DB_HOST`/`TEST_DB_NAME` coincidem com valores conhecidos de produção**
  (`72.61.53.20` / `stock_sys_db`, ou o que estiver num `backend/.env` real presente na
  máquina) → a sessão inteira é **abortada** (`pytest.exit`, código de saída 1), bem
  mais agressivo que um simples skip. Isso é intencional: é a última linha de defesa
  contra rodar os testes contra produção por engano.
- **Configuração válida, mas o MySQL não está acessível** (Docker não subiu, porta
  errada, credenciais erradas) → também pula, com a mensagem de conexão do PyMySQL
  incluída.

Exemplo de uso local (bash):

```bash
export TEST_DB_HOST=127.0.0.1
export TEST_DB_PORT=3307
export TEST_DB_NAME=stock_sys_test
export TEST_DB_USER=stock_test
export TEST_DB_PASSWORD=stock_test_pw
export JWT_SECRET=$(python -c "import secrets; print(secrets.token_urlsafe(48))")
```

## Como rodar a suíte

Sempre a partir de `backend/`, com o Python do venv do projeto:

```bash
cd backend

# Só coletar (deve sempre funcionar, com ou sem banco/variáveis configuradas — desde
# que JWT_SECRET esteja definido, ver tabela acima)
python -m pytest --collect-only

# Só os testes unitários (não tocam banco nem HTTP, sempre rodam)
python -m pytest -m unit

# Só os testes de integração (precisam do banco de teste no ar)
python -m pytest -m integration

# Tudo
python -m pytest

# Com cobertura (pytest-cov)
python -m pytest --cov=app --cov-report=term-missing -m integration
```

`backend/pytest.ini` registra os markers `unit`, `integration` e `mudanca_esperada`
(`--strict-markers` está ligado: usar um marker não registrado é erro de coleta, de
propósito). **Neste momento (pós Módulo 7b) nenhum teste usa `mudanca_esperada`** — as
duas mudanças que estavam marcadas assim (itens nº 3 e nº 4 do ciclo anterior) já
chegaram e foram promovidas para
["Bugs corrigidos com teste de regressão"](#bugs-corrigidos-com-teste-de-regressão). O
marker continua registrado em `pytest.ini` para uso em ciclos futuros.

## Como interpretar uma falha depois de uma refatoração

1. **Leia a mensagem de asserção inteira antes de mexer em qualquer código.** Os testes
   fixam strings exatas em português e comparam status HTTP exatos. Uma falha
   geralmente mostra *exatamente* o que mudou: a mensagem, o status code, ou o valor de
   um campo.
2. **Verifique se a falha é numa área que a mudança deveria tocar.** Se um teste de
   `test_monthly_report.py` falhar depois de uma mudança que só deveria afetar
   `issue()`, isso é suspeito e vale investigar antes de "só atualizar o teste".
3. **Nunca "conserte" um teste de caracterização só trocando o valor esperado para o
   que o código passou a devolver**, sem entender por quê. Se a mudança for
   intencional, atualize o teste e deixe um comentário dizendo qual mudança de
   comportamento motivou e por que ela é correta — não apague o histórico
   silenciosamente. Foi assim que os itens 1, 3, 4, 5 e 6 (parcial) desta lista
   viraram a seção "Bugs corrigidos com teste de regressão".
4. **Se a falha for causada por um bug real do código de produção** (não uma mudança
   de contrato intencional), **não a mascare**: marque o teste com
   `@pytest.mark.xfail(reason="...", strict=True)` explicando o bug, e reporte — ver
   ["Bug de produto encontrado neste ciclo"](#bug-de-produto-encontrado-neste-ciclo-reportado-não-corrigido)
   para o padrão adotado. `strict=True` faz a suíte AVISAR (o teste vira "XPASS", que
   conta como falha) se um dia o bug for corrigido sem que ninguém atualize o teste.
5. **Nem toda falha é bug de produto ou mudança de contrato — às vezes é bug do
   PRÓPRIO teste** (fixture com data hardcoded que ficou no passado, dois fixtures que
   sem querer mutam o mesmo registro, um CPF de exemplo que nunca foi válido). Ver
   ["Mudanças de comportamento deste ciclo"](#mudanças-de-comportamento-deste-ciclo-módulo-7b)
   para exemplos concretos encontrados nesta rodada — todos com comentário explicando
   a causa raiz no próprio teste corrigido.
6. **Testes que dependem de `backend/modelos/*.docx`** (geração de termo de empréstimo
   e de devolução) fazem `pytest.skip` quando os arquivos não existem nesta máquina.
   Neste worktree (Módulo 7b) os arquivos **existem** (ao contrário do worktree do
   Módulo 7a, que os escreveu sem eles) — por isso `TestCaminhoFeliz::
   test_fluxo_completo_disponivel_ate_devolvido` roda o fluxo completo de verdade
   aqui, em vez de pular na parte da devolução.

## Mudanças de comportamento deste ciclo (Módulo 7b)

A refatoração ampla mesclada antes deste módulo mudou vários contratos de API de
propósito. Cada uma abaixo tem o antes/depois e os testes que passaram a afirmar o
comportamento novo.

### 1. Envelope de paginação em GET /api/items e GET /api/history

- **Antes:** `GET /api/items` e `GET /api/history` respondiam um array JSON cru
  (`[{...}, {...}]`).
- **Depois:** respondem `{"items": [...], "total": N}` (schema `Paginated`, ver
  `app/schemas/common.py`) e aceitam `limit`/`offset` na querystring (`limit` rejeitado
  com 422 acima de `settings.MAX_PAGE_SIZE`). Do lado da camada de dados,
  `InventoryDBManager.list_items()`/`list_history()` passaram a devolver a tupla
  `(linhas, total)` em vez de só a lista.
- **Testes atualizados:** `test_items.py::TestListarEBuscarItem` (4 testes, agora
  leem `resp.json()["items"]`), mais a nova classe `TestPaginacaoDeItems` (limit,
  offset, offset além do fim, limit acima do máximo). `test_reversal.py::_history_id`
  e `test_rbac.py::test_estornar_historico_somente_gestor` desempacotam a tupla de
  `inv_manager.list_history()`. Cobertura da paginação/busca de `GET /api/history` em
  si ganhou um arquivo dedicado: `test_history.py`.

### 2. Validações de domínio migraram dos routers para os schemas Pydantic (400 → 422)

- **Antes:** CPF (dígito verificador), nota fiscal (9 dígitos), MAC, IP e datas
  `dd/mm/aaaa` eram validados dentro dos routers/`InventoryDBManager`, devolvendo
  `HTTPException(400, "mensagem em português")`.
- **Depois:** a validação vive nos schemas (`app/schemas/validators.py` +
  `field_validator` em `ItemCreate`/`ItemUpdate`/`LoanRequest`), o FastAPI responde
  422. Um handler global (`app/main.py::erro_de_validacao`, decorado com
  `@app.exception_handler(RequestValidationError)`) reduz a lista padrão de erros do
  Pydantic a uma frase única em português citando o **rótulo amigável do campo**
  (`_ROTULOS_DE_CAMPO`), porque o frontend lê `response.data.detail` esperando uma
  string — com a lista crua do Pydantic ele mostraria "[object Object]". A lista
  original continua disponível em `errors`, para quem depura ou consome a API
  programaticamente.
- **Testes atualizados:** `test_items.py::test_criar_item_com_data_em_formato_invalido`,
  `test_loan_workflow.py` (7 testes — a maioria só precisava de um CPF de teste
  *de fato válido*: o fixture usava `"11122233344"`, que nunca teve dígitos
  verificadores corretos; a validação nova simplesmente passou a notar isso).
  Cobertura nova do próprio contrato do handler (detail é string, contém o rótulo,
  `errors` continua lista): ver os testes de `test_items.py` e
  `test_loan_workflow.py` citados acima, que checam `isinstance(detail, str)` e
  `isinstance(errors, list)` explicitamente.

### 3. Estorno passou a exigir reconfirmação de senha (T6)

- **Antes:** `POST /api/history/{id}/reverse` não tinha corpo — só checava o role
  (Gestor).
- **Depois:** exige `{"password": "..."}` no corpo (`app/schemas/history.py::
  ReverseRequest`); sem o campo (ou vazio) é 422. Isso existe porque estornar é uma
  operação destrutiva sobre o histórico (T6): o operador logado precisa reconfirmar a
  própria senha, não só ter o role certo.
- **Testes atualizados:** todos os testes de `test_reversal.py` que iam por HTTP
  passaram a chamar `InventoryDBManager.reverse_history_entry()` diretamente para
  caracterizar a MÁQUINA DE ESTADOS do estorno — porque a checagem de senha em si
  está quebrada neste momento (ver
  ["Bug de produto encontrado"](#bug-de-produto-encontrado-neste-ciclo-reportado-não-corrigido)
  abaixo). A parte do contrato HTTP que não depende do código com bug (422 sem
  `password`/com `password` vazio) está coberta em
  `test_reversal.py::TestEstornoComSenhaSchemaHttp`.

### 4. `confirm_return()` agora remove o vínculo em `equipment_peripherals` (era mudança esperada nº 4/item 16)

- **Antes:** ao confirmar uma devolução, o status de cada periférico vinculado virava
  `'Disponível'`, mas a linha em `equipment_peripherals` nunca era apagada — o
  periférico continuava aparecendo como vinculado ao equipamento devolvido em
  `GET /api/items/{id}/peripherals`.
- **Depois:** `InventoryDBManager.confirm_return()` também executa
  `DELETE FROM equipment_peripherals WHERE equipment_id=%s` e registra
  `'Desvínculo Periférico'` no histórico para cada periférico que estava vinculado.
- **Teste:** `test_peripherals.py::TestEfeitoDoFluxoDeEmprestimoSobrePerifericos::
  test_confirmar_devolucao_libera_periferico_e_remove_o_vinculo` (renomeado de
  `..._mas_nao_remove_o_vinculo`; era marcado `@pytest.mark.mudanca_esperada`, o marker
  foi removido). Movido para "Bugs corrigidos com teste de regressão" (nº 4) abaixo.

### 5. Mensagem técnica ao gerar termo de devolução com arquivo ausente (era mudança esperada nº 3)

- **Antes:** para uma revenda válida (presente em `TERMO_DEVOLUCAO_MODELOS`), a
  checagem amigável só disparava se a revenda fosse desconhecida; um `.docx` ausente
  para revenda válida caía direto em `Document(modelo_path)`, cuja exceção virava
  `"Erro ao gerar documento: {e}"` crua.
- **Depois:** `generate_return_term_bytes()` chama `os.path.exists(modelo_path)`
  incondicionalmente, então a mensagem amigável (`"Modelo de termo de devolução não
  encontrado para {revenda}."`) dispara também para revenda válida com arquivo
  ausente.
- **Teste:** `test_loan_workflow.py::TestIniciarDevolucaoTransicoesInvalidas::
  test_iniciar_devolucao_com_arquivo_de_modelo_ausente_retorna_mensagem_amigavel`
  (renomeado de `..._retorna_erro_tecnico_nao_amigavel`; era
  `@pytest.mark.mudanca_esperada`, marker removido). Como os `.docx` **existem** neste
  worktree (ao contrário do Módulo 7a), o teste usa `monkeypatch` para forçar,
  deterministicamente, um caminho inexistente para uma revenda válida — não depende
  mais de o ambiente ter ou não os arquivos reais. Movido para "Bugs corrigidos com
  teste de regressão" (nº 3) abaixo.

## Bugs corrigidos com teste de regressão

### nº 1 — [CRÍTICO, CORRIGIDO] Autenticação por JWT quebrada pelo tipo do claim "sub"

**Histórico:** este módulo encontrou que `app/routers/auth.py::login()` montava o
claim `"sub"` do token com `user["id"]` — um `int` nativo do MySQL — enquanto
`python-jose` valida no decode, por padrão (`verify_sub=True`), que `"sub"` seja uma
*string*, rejeitando qualquer outro tipo com
`JWTClaimsError("Subject must be a string.")`. `app/core/security.decode_token()`
capturava isso como um `JWTError` genérico e devolvia 401
`"Token inválido ou expirado."`. Na prática: o login funcionava e devolvia um
`access_token`, mas **esse mesmo token era rejeitado em qualquer rota protegida**.

**Testes:** `test_auth.py::TestLogin::
test_regressao_sub_do_jwt_emitido_pelo_login_deve_ser_string`,
`test_auth.py::TestMe::test_token_de_login_real_autentica_em_rotas_protegidas`,
`test_auth.py::TestRefresh::test_refresh_com_cookie_de_login_real_emite_novo_access_token`,
`test_unit_seguranca.py::TestTokens::test_decode_token_rejeita_sub_nao_string_regressao_bug_1`.

### nº 2 — [CORRIGIDO] `link_peripheral_to_equipment()` não tratava o UNIQUE de forma amigável

**Antes:** vincular o mesmo par (equipamento, periférico) duas vezes violava o
`UNIQUE(equipment_id, peripheral_id)` de `equipment_peripherals`, e o usuário recebia a
exceção do PyMySQL praticamente crua, só prefixada com `"Erro ao vincular: "` (ex.:
`"Erro ao vincular: (1062, \"Duplicate entry...\")"`).

**Depois:** `link_peripheral_to_equipment()` captura `pymysql.MySQLError` de forma
genérica e devolve `"Erro ao vincular periférico."` — sem vazar a exceção do driver.
Não chega a ser uma mensagem específica de "já vinculado" (ainda não distingue esse
UNIQUE do resto dos erros de banco, ao contrário de `add_peripheral()` com o
identificador duplicado), mas o vazamento técnico, que era o bug real, acabou.

**Teste:** `test_peripherals.py::TestVincularEDesvincular::
test_vincular_o_mesmo_par_duas_vezes_retorna_mensagem_amigavel` (renomeado; a
asserção antiga usava `.startswith("Erro ao vincular:")`).

### nº 3 — [CORRIGIDO] Mensagem técnica ao faltar o arquivo de modelo de devolução para revenda válida

Ver item 5 de ["Mudanças de comportamento deste ciclo"](#mudanças-de-comportamento-deste-ciclo-módulo-7b)
acima ("Mensagem técnica ao gerar termo de devolução com arquivo ausente") —
descrição completa e teste.

### nº 4 — [CORRIGIDO] `confirm_return()` não desfazia o vínculo em `equipment_peripherals`

Ver item 4 de ["Mudanças de comportamento deste ciclo"](#mudanças-de-comportamento-deste-ciclo-módulo-7b)
acima ("`confirm_return()` agora remove o vínculo em `equipment_peripherals`") —
descrição completa e teste.

### nº 5 — [CORRIGIDO PARCIALMENTE] `remove_user()` reportava sucesso mesmo com id inexistente

**Antes:** `UserDBManager.remove_user()` executa um `DELETE` sem checar antes se o id
existe nem olhar `cursor.rowcount` depois — um `DELETE` que não afeta nenhuma linha não
levanta exceção no MySQL, então a função sempre devolvia `(True, "...com sucesso.")`,
mesmo quando nada foi removido.

**Depois (parcial):** `UserDBManager.remove_user()` em si **não mudou**. Mas
`app/routers/users.py::remove_user()` (o router) passou a buscar o usuário-alvo com
`user_db.get_user_by_id(user_id)` **antes** de chamar `remove_user()` — como parte da
proteção de lockout (T5: Gestor não remove a si mesmo nem o último Gestor). Efeito
colateral correto: um id inexistente agora é barrado nessa checagem e devolve 404
antes mesmo de chegar no `DELETE` sem efeito. **`update_password()`/
`PUT /api/users/{id}/password` continuam com o mesmo bug, sem nenhuma checagem
equivalente** — ver a entrada correspondente em
["Lista de BUG CONHECIDO"](#lista-de-bug-conhecido-fora-de-escopo-desta-rodada) abaixo.

**Teste:** `test_users.py::TestRemoverUsuario::
test_remover_usuario_inexistente_retorna_404` (renomeado de
`..._tambem_retorna_sucesso`).

## Bug de produto encontrado neste ciclo (reportado, não corrigido)

> Por instrução explícita deste módulo, bugs de produto são **reportados, não
> corrigidos** — a correção é decisão de quem mantém `backend/app/**`. O teste fica
> **falhando de propósito**, marcado `@pytest.mark.xfail(strict=True)`.

### `POST /api/history/{id}/reverse` quebra com 500 para qualquer senha, certa ou errada

**Onde:** `app/routers/history.py::reverse_entry()` (feature nova "estorno com
senha", T6 — ver item 3 de "Mudanças de comportamento deste ciclo" acima).

**O bug:** o endpoint verifica a senha do operador logado assim:

```python
user = user_db.get_user_by_id(current_user.id)
if not user or not verify_password(body.password, user["password"]):
    ...
```

Mas `UserDBManager.get_user_by_id()` seleciona **de propósito** sem a coluna
`password` (`"SELECT id, username, role FROM usuarios WHERE id = %s"` — o próprio
docstring do método diz "sem o hash da senha"; é o mesmo método que
`users.py::remove_user()` usa para a checagem de lockout, que não precisa de senha). O
método certo para obter o hash seria `get_user_by_username(current_user.username)`,
usado corretamente em `auth.py::login()`.

**Consequência:** `user["password"]` sempre levanta `KeyError: 'password'`. Todo
`POST /api/history/{id}/reverse` autenticado como Gestor — com a senha certa OU errada
— quebra com 500 Internal Server Error em vez de 200/403. A feature "estorno com
senha" está, no estado atual do código, **completamente inoperante via HTTP**.

**Como foi encontrado:** ao atualizar `test_reversal.py` para enviar `password` no
corpo (mudança de contrato esperada, item 3 acima), os testes de caminho feliz
passaram a quebrar com um `KeyError` vindo de dentro do próprio endpoint, não com uma
falha de asserção — sinal de bug de produto, não de teste desatualizado.

**Testes:**
- `test_reversal.py::TestBugDeProdutoSenhaDoEstorno::
  test_senha_correta_deveria_estornar_mas_quebra_com_keyerror` (`xfail`, `strict=True`)
- `test_reversal.py::TestBugDeProdutoSenhaDoEstorno::
  test_senha_errada_deveria_dar_403_mas_quebra_com_keyerror` (`xfail`, `strict=True`)
- `test_rbac.py::TestRbacHistorico::test_estornar_historico_somente_gestor` também
  documenta o bug: o caminho positivo (Gestor) usa `pytest.raises(KeyError)` em vez de
  afirmar 200, com um comentário explicando que a exceção acontece DEPOIS da checagem
  de role (ou seja, o RBAC em si está correto — só a lógica de senha, mais abaixo, é
  que está quebrada).

**Correção sugerida (não aplicada aqui — fora da posse deste módulo):** trocar
`user_db.get_user_by_id(current_user.id)` por
`user_db.get_user_by_username(current_user.username)` em
`app/routers/history.py::reverse_entry()`.

**O que continua funcionando (não afetado pelo bug):** a validação do corpo pelo
schema `ReverseRequest` roda ANTES de `reverse_entry()` tocar em `user["password"]` —
por isso `password` ausente ou vazio ainda dá 422 normalmente (coberto em
`test_reversal.py::TestEstornoComSenhaSchemaHttp`), e toda a MÁQUINA DE ESTADOS do
estorno (o que cada tipo de operação desfaz) continua caracterizada e passando, só que
via `InventoryDBManager.reverse_history_entry()` diretamente em vez do endpoint HTTP
(ver item 3 de "Mudanças de comportamento deste ciclo").

## Cobertura nova adicionada neste ciclo (T2)

Comportamento que não existia (ou não tinha teste dedicado) quando a suíte original
foi escrita:

- **Revogação de refresh token** (`test_auth.py::TestRevogacaoDeRefreshToken`):
  logout revoga o `jti` do refresh token; um `/api/auth/refresh` com token revogado dá
  401 "Sessão encerrada. Faça login novamente."; a rotação (todo `/refresh`
  bem-sucedido) invalida o refresh token anterior (tentar reutilizá-lo também dá 401).
- **Rate limiting do login** (`test_auth.py::TestRateLimitDoLogin`): a 6ª tentativa de
  login numa janela de 1 minuto dá 429. Como o `conftest.py` eleva
  `LOGIN_RATE_LIMIT` para `"10000/minute"` para a suíte inteira (senão as próprias
  fixtures de autenticação tropeçariam no limite ao longo de uma sessão de testes), e
  o valor é "congelado" no decorator `@limiter.limit(...)` no momento em que
  `app/routers/auth.py` é importado (não é reavaliado a cada request), este teste
  usa um helper (`_cliente_com_login_rate_limit`) que reimporta a cadeia
  `app.core.config -> app.dependencies -> app.routers.auth -> app.main` com a env var
  já no valor baixo, monta um `TestClient` novo, e desfaz tudo ao final — sem afetar o
  app compartilhado (`_fastapi_app`, session-scoped) usado pelo resto da suíte.
- **Estorno com senha** (`test_reversal.py::TestEstornoComSenhaSchemaHttp` +
  `TestBugDeProdutoSenhaDoEstorno`): sem o campo `password` (ou vazio) dá 422 — isso
  funciona. Senha correta estorna e senha errada dá 403 sem estornar — isso **não
  funciona hoje** por causa do bug de produto descrito acima; os testes ficam
  documentados como `xfail`.
- **Proteção de lockout** (`test_users.py::TestProtecaoDeLockout`): Gestor não remove
  a si mesmo; com dois Gestores, remover um funciona; não remove o último Gestor —
  este último caso só é alcançável, na prática, com o token do Gestor logado já
  emitido mas o role dele rebaixado diretamente no banco (`executar_sql`, novo fixture
  de teste), porque removê-lo a si mesmo já seria barrado ANTES pela checagem de
  autorremoção (ver docstring do teste para a prova completa).
- **`DELETE /api/peripherals/{id}`** (`test_peripherals.py::TestInativarPeriferico`):
  inativa (soft-delete) periférico disponível; recusa com 400 se estiver "Em Uso"
  (nota: vincular já marca como "Em Uso", então esse é o ramo que dispara antes do
  ramo "vinculado a um equipamento" ter a chance de rodar); RBAC (Gestor ou Técnico,
  Jovem Aprendiz barrado).
- **Busca server-side no histórico** (`test_history.py`): `?search=` filtra por
  operador, usuário, operação, revenda, tipo, marca, modelo e identificador
  (case-insensitive); `total` reflete o filtro, não o total geral. Também cobre a
  paginação (`limit`/`offset`/`total`/offset além do fim/limit acima do máximo) do
  mesmo endpoint.
- **Upload** (`test_uploads.py`): arquivo acima de 20MB dá 413 (testado em
  `DELETE /api/items/{id}` e `POST /api/loans/{id}/confirm`, ambos usam a mesma
  constante `MAX_UPLOAD_SIZE`); nome de arquivo com `../` é sanitizado por
  `_safe_filename()` (via `os.path.basename`) e não escapa do diretório da categoria
  de storage; nome só com barras cai no fallback `"arquivo"`.
- **`/api/health`/`/api/health/db`** (`test_health.py`): `/api/health` responde 200
  sempre, sem tocar no banco e sem exigir autenticação; `/api/health/db` reporta 503
  com `{"status": "unavailable", "database": "unreachable", "detail": "..."}` quando o
  banco falha (simulado com `monkeypatch` em `app.main.get_inventory_db`, sem derrubar
  o banco de teste de verdade) e não derruba `/api/health`.

## Lista de BUG CONHECIDO (fora de escopo desta rodada)

Bugs reais, verificados diretamente contra o código de `backend/app/**`, sem correção
planejada nesta rodada. Os testes continuam afirmando o comportamento atual; nenhuma
ação é esperada além de manter os testes como estão.

- **nº 1 — `format_date()` não reformata datas que já vêm com hora embutida em
  string.** `app/db/utils.py::format_date()` só reconhece objetos `datetime`/`date`
  ou strings `"YYYY-MM-DD"`; uma string `"YYYY-MM-DD HH:MM:SS"` falha no `strptime`
  interno e é devolvida **sem alteração**. Afeta diretamente
  `InventoryDBManager.issue()`, que monta a mensagem de erro "data de empréstimo
  anterior ao cadastro" chamando `format_date(str(item['date_registered']))`.
  Testes: `test_unit_formatacao.py::TestFormatDate::test_string_iso_com_hora_nao_e_reformatada`,
  `test_loan_workflow.py::TestIniciarEmprestimoTransicoesInvalidas::test_data_de_emprestimo_anterior_ao_cadastro`.

- **nº 2 — `update_password()` reporta sucesso mesmo quando o id não existe.** Mesma
  classe de bug que afetava `remove_user()` (ver nº 5 em "Bugs corrigidos com teste de
  regressão" acima) — mas aqui **não foi corrigido**: `PUT /api/users/{id}/password`
  não tem uma checagem de existência equivalente à que `remove_user()` ganhou no
  router. `update_password()` não confere `cursor.rowcount`, então um `UPDATE` que não
  afeta nenhuma linha ainda assim devolve sucesso.
  Teste: `test_users.py::TestAtualizarSenha::test_atualizar_senha_de_usuario_inexistente_tambem_retorna_sucesso`.

Além da lista acima, há uma **inconsistência observada** que não foi classificada
como bug por não ter um "comportamento errado" claro, só uma falta de padronização:
`GET`/`PUT /api/items/{id}` devolvem `404 "Item não encontrado."` para um id
inexistente, enquanto `DELETE /api/items/{id}` devolve `400 "ID não encontrado."`
para a mesma situação. Ver
`test_items.py::TestRemoverItem::test_remover_item_inexistente_retorna_400_nao_404`.

## Reclassificado: comportamento esperado, não bug

### nº 1 — `format_cpf()` não valida dígito verificador

`format_cpf()` é um **formatador**, não um validador — seu contrato é "tem 11
dígitos? então pontua; senão, devolve como veio", sem checar dígito verificador. Isso
é esperado: a validação de CPF de verdade vive em `isValidCpf` no frontend e em
`app/schemas/validators.py::validar_cpf()` no backend (usado por `LoanRequest.cpf`),
não em `format_cpf()`. O teste documenta esse contrato:
`test_unit_formatacao.py::TestFormatCpf::test_sem_11_digitos_nao_formata_pois_nao_e_papel_desta_funcao_validar`.
