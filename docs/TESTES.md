# Testes de caracterização (Módulo 7a)

Esta suíte **não testa o comportamento desejado** do sistema — ela testa o
comportamento **atual**, como rede de segurança para a refatoração em andamento do
context manager de conexão em `backend/app/db/inventory_manager_db.py` (e das demais
mudanças paralelas). Depois de qualquer merge, rode a suíte de novo: testes que
passavam e passaram a falhar apontam exatamente o que mudou.

Nem todo teste caracteriza um comportamento que deve continuar existindo para sempre:

- Testes marcados `# BUG CONHECIDO` no código documentam bugs reais e **ainda fora de
  escopo** — devem continuar passando (afirmando o comportamento antigo) até alguém
  corrigir o bug de propósito.
- Testes marcados `@pytest.mark.mudanca_esperada` documentam bugs reais com
  **correção já planejada para este mesmo ciclo** por outro módulo — também afirmam o
  comportamento antigo de propósito, mas espera-se que **passem a falhar** assim que
  a correção for mesclada. Isso não é uma regressão: é o sinal de que a correção
  chegou. Ver [Mudanças esperadas neste ciclo](#mudanças-esperadas-neste-ciclo).
- Testes na seção [Bugs corrigidos com teste de regressão](#bugs-corrigidos-com-teste-de-regressão)
  já documentam um bug que **foi corrigido** antes mesmo do merge — eles afirmam o
  comportamento CORRETO e incluem um teste de regressão dedicado para impedir que o
  bug volte.

## Sumário

- [Como subir o banco de teste](#como-subir-o-banco-de-teste)
- [Variáveis de ambiente](#variáveis-de-ambiente)
- [Como rodar a suíte](#como-rodar-a-suíte)
- [Como interpretar uma falha depois de uma refatoração](#como-interpretar-uma-falha-depois-de-uma-refatoração)
- [Limitações conhecidas deste ambiente](#limitações-conhecidas-deste-ambiente)
- [Bugs corrigidos com teste de regressão](#bugs-corrigidos-com-teste-de-regressão)
- [Mudanças esperadas neste ciclo](#mudanças-esperadas-neste-ciclo)
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

Comportamento do guard (`backend/tests/conftest.py::_ambiente_de_teste`):

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
```

## Como rodar a suíte

Sempre a partir de `backend/`, com o Python do venv do projeto:

```bash
cd backend

# Só coletar (deve sempre funcionar, com ou sem banco/variáveis configuradas)
python -m pytest --collect-only

# Só os testes unitários (não tocam banco nem HTTP, sempre rodam)
python -m pytest -m unit

# Só os testes de integração (precisam do banco de teste no ar)
python -m pytest -m integration

# Tudo
python -m pytest

# Com cobertura (pytest-cov)
python -m pytest --cov=app --cov-report=term-missing -m integration

# Só os testes de "mudança esperada neste ciclo" (bugs 3 e 4 — ver seção dedicada);
# útil logo depois de mesclar o trabalho do Módulo 2, para checar rapidamente se a
# correção chegou
python -m pytest -m mudanca_esperada -v

# Tudo, EXCETO os testes de mudança esperada (para não poluir uma rodada de CI comum
# com falhas "esperadas" enquanto a correção do Módulo 2 ainda não chegou)
python -m pytest -m "not mudanca_esperada"
```

`backend/pytest.ini` registra os markers `unit`, `integration` e `mudanca_esperada`
(`--strict-markers` está ligado: usar um marker não registrado é erro de coleta, de
propósito).

## Como interpretar uma falha depois de uma refatoração

1. **Leia a mensagem de asserção inteira antes de mexer em qualquer código.** Os testes
   fixam strings exatas em português (`"Este item não está disponível para
   empréstimo."`, `"Apenas itens com status 'Pendente' podem ser confirmados."` etc.) e
   comparam status HTTP exatos. Uma falha geralmente mostra *exatamente* o que mudou:
   a mensagem, o status code, ou o valor de um campo.
2. **Verifique se a falha é numa área que a sua mudança deveria tocar.** Ex.: se a
   refatoração do connection manager mudou o *tratamento de exceções* dentro de
   `add_peripheral`/`link_peripheral_to_equipment`/etc., é esperado que
   `test_peripherals.py` ou `test_items.py` acusem — isso é a rede de segurança
   funcionando. Se um teste de `test_monthly_report.py` falhar depois de uma mudança
   que só deveria afetar `issue()`, isso é suspeito e vale investigar antes de "só
   atualizar o teste".
3. **Um teste `@pytest.mark.mudanca_esperada` que passa a falhar é o resultado
   esperado**, não uma regressão — ver [Mudanças esperadas neste ciclo](#mudanças-esperadas-neste-ciclo)
   para a lista exata e o que fazer com cada um quando isso acontecer.
4. **Um teste com `# BUG CONHECIDO` (sem `mudanca_esperada`) que passa a falhar de um
   jeito novo** (não simplesmente "parou de reproduzir o bug") pode indicar que o bug
   mudou de forma (ex.: a mensagem de erro crua mudou de texto porque a exceção agora
   vem de outro lugar) — vale conferir se é só isso ou se é sintoma de outra coisa.
5. **Um teste com `# BUG CONHECIDO` (sem `mudanca_esperada`) que passa a "passar como
   se o bug tivesse sumido"**, sem que ninguém tenha mexido de propósito naquele bug e
   sem que ele estivesse na lista de mudanças esperadas, é o sinal mais importante
   desta suíte: normalmente indica um efeito colateral não intencional da refatoração
   (bom ou ruim, mas não intencional) que vale entender antes de simplesmente apagar o
   teste.
6. **Nunca "conserte" um teste de caracterização só trocando o valor esperado para o
   que o código passou a devolver**, sem entender por quê. Se a mudança for
   intencional, atualize o teste e troque o comentário `# BUG CONHECIDO`/
   `mudanca_esperada` por uma nota indicando que foi corrigido em tal mudança — não
   apague o histórico silenciosamente (foi assim que o item 1 virou a seção
   "Bugs corrigidos com teste de regressão" abaixo).
7. **Testes que dependem de `backend/modelos/*.docx`** (geração de termo de empréstimo
   e de devolução) fazem `pytest.skip` quando os arquivos não existem nesta máquina —
   isso é esperado neste worktree específico (ver seção seguinte) e deve desaparecer
   assim que os templates estiverem presentes (depois do merge com o restante do
   trabalho em andamento).

## Limitações conhecidas deste ambiente

No momento em que esta suíte foi escrita, o worktree deste módulo foi criado a partir
de um commit **anterior** ao estado atual da branch de integração. Duas consequências
práticas, que não são bugs do sistema e não foram caracterizadas como tal:

- `backend/app/db/database_mysql.py` (que define `get_connection()`, usado por toda a
  camada de dados) ainda não existe neste worktree — é posse do módulo que está
  reescrevendo a camada de conexão. `backend/tests/conftest.py` importa esse módulo
  **só dentro do corpo de fixtures** (nunca no nível de arquivo) e transforma qualquer
  falha de import numa mensagem de skip clara, então a ausência do arquivo não quebra
  `--collect-only` nem faz a suíte "explodir" — ela só pula os testes de integração,
  com uma mensagem que cita o módulo em falta.
- `backend/modelos/**` (os `.docx` de termo de empréstimo/devolução) também não existe
  neste worktree. Os testes que dependem de gerar um documento de verdade
  (`test_loan_workflow.py::TestCaminhoFeliz` e o teste do "modelo ausente") checam a
  presença desses arquivos em tempo de execução (fixtures `termo_devolucao_disponivel`
  / `termo_emprestimo_disponivel`) e pulam a parte que não pode ser caracterizada sem
  eles, com uma mensagem explicando o motivo. Assim que os templates existirem (depois
  do merge), essas partes passam a rodar de verdade sem precisar tocar no teste.
- Por conta dos dois pontos acima **e** do Docker daemon não estar disponível durante
  este trabalho, não foi possível rodar a suíte de integração contra um banco MySQL
  real — ver a seção final do relatório de entrega para os detalhes do que *foi*
  validado (coleta, guard de segurança, skip por ausência de configuração, skip por
  banco inacessível).

## Bugs corrigidos com teste de regressão

### nº 1 — [CRÍTICO, CORRIGIDO] Autenticação por JWT quebrada pelo tipo do claim "sub"

**Histórico:** este módulo encontrou que `app/routers/auth.py::login()` montava o
claim `"sub"` do token com `user["id"]` — um `int` nativo do MySQL — enquanto
`python-jose` valida no decode, por padrão (`verify_sub=True`), que `"sub"` seja uma
*string*, rejeitando qualquer outro tipo com
`JWTClaimsError("Subject must be a string.")`. `app/core/security.decode_token()`
capturava isso como um `JWTError` genérico e devolvia 401
`"Token inválido ou expirado."`. Na prática: o login funcionava e devolvia um
`access_token`, mas **esse mesmo token era rejeitado em qualquer rota protegida**,
inclusive `/api/auth/me` e o refresh via cookie — nenhum usuário conseguia de fato
*usar* a API além de logar.

O achado foi confirmado de forma independente pelo Módulo 3, e a correção (`login()`
grava `"sub"` como string; `get_current_user()` faz `int(payload["sub"])`) já estava
no código-alvo desta refatoração no momento em que isso foi reconciliado com o
coordenador — três fontes convergindo.

**Testes (afirmam o comportamento CORRETO/atual):**
- `test_auth.py::TestLogin::test_regressao_sub_do_jwt_emitido_pelo_login_deve_ser_string`
  — teste de regressão dedicado: lê (sem verificar assinatura) o payload de um token
  de login real e confere que `"sub"` é string. Não pressupõe como a correção foi
  feita — continua válido mesmo que a estratégia mude, desde que `"sub"` siga sendo
  string.
- `test_auth.py::TestMe::test_token_de_login_real_autentica_em_rotas_protegidas` —
  login real → token real → `/api/auth/me` e `/api/items` devolvem 200 (não mais 401).
- `test_auth.py::TestRefresh::test_refresh_com_cookie_de_login_real_emite_novo_access_token`
  — o cookie de refresh emitido pelo login também autentica corretamente.
- `test_unit_seguranca.py::TestTokens::test_decode_token_rejeita_sub_nao_string_regressao_bug_1`
  — no nível unitário, documenta que `decode_token()` sempre vai recusar um `"sub"`
  não-string (isso é comportamento correto/intencional do `jose`, não o bug em si; o
  bug era `login()` produzir esse tipo errado).

**Consequência para o resto da suíte:** como o bug estava corrigido no alvo, os
fixtures `client_gestor`/`client_tecnico`/`client_aprendiz` (`conftest.py`) fazem login
de verdade via `POST /api/auth/login` e mandam um token JWT real no header
`Authorization` — **sem nenhum contorno**. Isso significa que toda a matriz de RBAC
(`test_rbac.py`) e os testes de fluxo (itens, empréstimos, periféricos, histórico,
relatórios) exercitam de fato `get_current_user()`, `decode_token()` e a checagem de
expiração, não só a lógica de `gestor_only`/`gestor_or_tecnico` isoladamente.

## Mudanças esperadas neste ciclo

Os dois itens abaixo são bugs reais e verificados, com **correção já em andamento
pelo Módulo 2 neste mesmo ciclo**. Os testes continuam afirmando o comportamento
**atual** de propósito — não foram (e não devem ser) atualizados para o comportamento
novo. Estão marcados `@pytest.mark.mudanca_esperada` no código para serem
distinguíveis de bugs comuns: depois do merge com o Módulo 2, espera-se que **falhem**
— isso é o sinal de que a correção chegou, não uma regressão. Quando isso acontecer,
reescreva cada teste para afirmar o novo comportamento e mova o item para
"Bugs corrigidos com teste de regressão" acima.

### nº 3 — Mensagem técnica (não amigável) quando falta o arquivo de modelo de uma revenda válida

Em `generate_return_term_bytes()` (e, pela mesma estrutura de código, também em
`generate_loan_term_bytes()`), a checagem amigável
`"Modelo de termo [...] não encontrado para {revenda}"` só é alcançada quando a
revenda é **desconhecida** (`TERMO_DEVOLUCAO_MODELOS.get(revenda)` devolve `None`). Se
a revenda é válida mas o `.docx` correspondente está ausente/corrompido no disco, o
código pula direto para `Document(modelo_path)`, cuja exceção é capturada de forma
genérica e devolvida crua como `"Erro ao gerar documento: {e}"` — sem a checagem de
existência do arquivo rodar antes. A transação no banco não é afetada (o item
permanece no status anterior), só a mensagem ao usuário é confusa.

**O Módulo 2 vai corrigir isso:** a checagem amigável deixará de ser código morto para
revenda válida e passará a disparar de verdade quando o arquivo não existir.

Teste: `test_loan_workflow.py::TestIniciarDevolucaoTransicoesInvalidas::test_iniciar_devolucao_com_modelo_ausente_retorna_erro_tecnico_nao_amigavel`
(só é executado quando os `.docx` realmente não existem na máquina — ver
"Limitações conhecidas deste ambiente" acima; `generate_loan_term_bytes()` não tem um
teste dedicado equivalente, mas o mesmo padrão de código se aplica lá).

### nº 4 — (item 16) `confirm_return()` não desfaz o vínculo em `equipment_peripherals`

Ao confirmar uma devolução, o código atualiza cada periférico vinculado para
`status='Disponível'` (correto), mas nunca executa um `DELETE` na tabela de vínculo.
Resultado: depois da devolução confirmada, o periférico aparece corretamente como
"Disponível" em `/api/peripherals`, mas **continua** aparecendo como vinculado àquele
equipamento em `GET /api/items/{id}/peripherals` — o vínculo só era desfeito por um
"desvincular" manual (`DELETE /api/peripherals/links/{link_id}`).

**O Módulo 2 vai corrigir isso:** `confirm_return()` passará a também apagar as linhas
de `equipment_peripherals` do equipamento devolvido e registrar
`'Desvínculo Periférico'` no histórico para cada periférico desvinculado.

Teste: `test_peripherals.py::TestEfeitoDoFluxoDeEmprestimoSobrePerifericos::test_confirmar_devolucao_libera_periferico_mas_nao_remove_o_vinculo`.

## Lista de BUG CONHECIDO (fora de escopo desta rodada)

Bugs reais, verificados diretamente contra o código de `backend/app/**` (e, quando
possível, reproduzidos executando as funções reais da aplicação, não uma
reimplementação), sem correção planejada nesta rodada. Os testes continuam afirmando
o comportamento atual; nenhuma ação é esperada além de manter os testes como estão.

- **nº 2 — `format_date()` não reformata datas que já vêm com hora embutida em
  string.** `app/db/utils.py::format_date()` só reconhece objetos `datetime`/`date`
  ou strings `"YYYY-MM-DD"`; uma string `"YYYY-MM-DD HH:MM:SS"` falha no `strptime`
  interno e é devolvida **sem alteração**. Isso afeta diretamente
  `InventoryDBManager.issue()`, que monta a mensagem de erro "data de empréstimo
  anterior ao cadastro" chamando `format_date(str(item['date_registered']))` — o
  `str()` prévio já inclui a hora, então o usuário vê a data crua do MySQL (ex.:
  `"2026-06-01 09:00:00"`) em vez de `"01/06/2026"`.
  Testes: `test_unit_formatacao.py::TestFormatDate::test_string_iso_com_hora_nao_e_reformatada`,
  `test_loan_workflow.py::TestIniciarEmprestimoTransicoesInvalidas::test_data_de_emprestimo_anterior_ao_cadastro`.

- **nº 5 — Vincular o mesmo par equipamento/periférico duas vezes devolve um erro
  cru, não amigável.** `equipment_peripherals` tem `UNIQUE(equipment_id,
  peripheral_id)`, mas `link_peripheral_to_equipment()` não trata essa violação como
  `add_peripheral()` trata o `UNIQUE` do identificador (que devolve
  `"Já existe um periférico com este Identificador (Nº de Série)."`). Aqui, o
  usuário recebe a exceção do PyMySQL praticamente crua, só prefixada com
  `"Erro ao vincular: "`.
  Teste: `test_peripherals.py::TestVincularEDesvincular::test_vincular_o_mesmo_par_duas_vezes_retorna_erro_tecnico_nao_amigavel`.

- **nº 6 — `remove_user()`/`update_password()` reportam sucesso mesmo quando o id não
  existe.** Achado novo deste módulo (não estava na auditoria inicial do
  coordenador) — registrado como item novo de backlog. Nenhuma das duas funções
  verifica a existência do usuário antes do `DELETE`/`UPDATE`, nem olha
  `cursor.rowcount` depois — um comando que não afeta nenhuma linha não levanta
  exceção no MySQL, então ambas sempre devolvem `(True, "...com sucesso.")`. Um
  Gestor não tem como saber, pela resposta da API, se removeu um usuário de verdade
  ou se o id já não existia.
  Testes: `test_users.py::TestRemoverUsuario::test_remover_usuario_inexistente_tambem_retorna_sucesso`,
  `test_users.py::TestAtualizarSenha::test_atualizar_senha_de_usuario_inexistente_tambem_retorna_sucesso`.

Além da lista acima, há uma **inconsistência observada** que não foi classificada
como bug por não ter um "comportamento errado" claro, só uma falta de padronização:
`GET`/`PUT /api/items/{id}` devolvem `404 "Item não encontrado."` para um id
inexistente, enquanto `DELETE /api/items/{id}` devolve `400 "ID não encontrado."`
para a mesma situação — status code e mensagem diferentes para o mesmo tipo de erro,
dependendo do endpoint. Ver
`test_items.py::TestRemoverItem::test_remover_item_inexistente_retorna_400_nao_404`.

## Reclassificado: comportamento esperado, não bug

### nº 7 — `format_cpf()` não valida dígito verificador

Reclassificado a pedido do coordenador: `format_cpf()` é um **formatador**, não um
validador — seu contrato é "tem 11 dígitos? então pontua; senão, devolve como veio",
sem checar dígito verificador. Isso é esperado: a validação de CPF de verdade vive em
`isValidCpf` no frontend e em `app/schemas/validators.py` no backend, não nesta
função. O teste foi reescrito como documentação desse contrato (não caracteriza mais
um "bug"):
`test_unit_formatacao.py::TestFormatCpf::test_sem_11_digitos_nao_formata_pois_nao_e_papel_desta_funcao_validar`.
