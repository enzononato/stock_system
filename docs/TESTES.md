# Testes de caracterização (Módulo 7a)

Esta suíte **não testa o comportamento desejado** do sistema — ela testa o
comportamento **atual**, como rede de segurança para a refatoração em andamento do
context manager de conexão em `backend/app/db/inventory_manager_db.py` (e das demais
mudanças paralelas). Depois de qualquer merge, rode a suíte de novo: testes que
passavam e passaram a falhar apontam exatamente o que mudou. Alguns testes documentam
bugs reais do sistema hoje (marcados `# BUG CONHECIDO` no código) — eles **devem**
continuar falhando-o-comportamento-antigo até alguém decidir corrigir o bug de
propósito; se um desses passar a "passar como se o bug não existisse mais mais sem
ninguém ter mexido nele de propósito", investigue, pode ser efeito colateral de outra
mudança.

## Sumário

- [Como subir o banco de teste](#como-subir-o-banco-de-teste)
- [Variáveis de ambiente](#variáveis-de-ambiente)
- [Como rodar a suíte](#como-rodar-a-suíte)
- [Como interpretar uma falha depois de uma refatoração](#como-interpretar-uma-falha-depois-de-uma-refatoração)
- [Limitações conhecidas deste ambiente](#limitações-conhecidas-deste-ambiente)
- [Lista completa de BUG CONHECIDO](#lista-completa-de-bug-conhecido)

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
```

`backend/pytest.ini` registra os markers `unit` e `integration` (`--strict-markers`
está ligado: usar um marker não registrado é erro de coleta, de propósito).

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
3. **Um teste com `# BUG CONHECIDO` que passa a falhar de um jeito novo** (não
   simplesmente "parou de reproduzir o bug") pode indicar que o bug mudou de forma
   (ex.: a mensagem de erro crua mudou de texto porque a exceção agora vem de outro
   lugar) — vale conferir se é só isso ou se é sintoma de outra coisa.
4. **Um teste com `# BUG CONHECIDO` que passa a "passar como se o bug tivesse sumido"**
   sem que ninguém tenha mexido de propósito naquele bug é o sinal mais importante desta
   suíte: normalmente indica um efeito colateral não intencional da refatoração (bom ou
   ruim, mas não intencional) que vale entender antes de simplesmente apagar o teste.
5. **Nunca "conserte" um teste de caracterização só trocando o valor esperado para o
   que o código passou a devolver**, sem entender por quê. Se a mudança for
   intencional (ex.: o bug foi corrigido de propósito nesta rodada), atualize o teste
   e troque o comentário `# BUG CONHECIDO` por uma nota indicando que foi corrigido em
   tal mudança — não apague o histórico silenciosamente.
6. **Testes que dependem de `backend/modelos/*.docx`** (geração de termo de empréstimo
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

## Lista completa de BUG CONHECIDO

Cada item abaixo tem um teste dedicado (ou está citado dentro de um) marcado
`# BUG CONHECIDO` (ou `BUG CONHECIDO (CRÍTICO)`/`(item 16)`) no código-fonte da suíte.
São comportamentos **atuais e reais** do sistema, não hipóteses — todos foram
verificados diretamente contra o código de `backend/app/**` (e, quando possível,
reproduzidos executando as funções reais da aplicação, não uma reimplementação).

1. **[CRÍTICO] Autenticação por JWT está quebrada para qualquer usuário.**
   `app/routers/auth.py::login()` monta o claim `"sub"` do token com
   `user["id"]` — um `int` nativo do MySQL — mas `python-jose` valida no decode,
   por padrão (`verify_sub=True`), que `"sub"` seja uma *string*, e rejeita
   qualquer outro tipo com `JWTClaimsError("Subject must be a string.")`.
   `app/core/security.decode_token()` captura isso como um `JWTError` genérico e
   devolve 401 `"Token inválido ou expirado."`. Na prática: o login funciona e
   devolve um `access_token`, mas **esse mesmo token é rejeitado em qualquer rota
   protegida**, inclusive `/api/auth/me` e o refresh via cookie (o refresh token
   tem o mesmo problema). Hoje, nenhum usuário consegue de fato *usar* a API além
   de logar.
   Testes: `test_unit_seguranca.py::TestTokens::test_decode_token_falha_com_sub_inteiro_como_o_login_real_gera`,
   `test_auth.py::TestMe::test_token_de_login_real_nao_autentica_em_nenhuma_rota`,
   `test_auth.py::TestRefresh::test_refresh_com_cookie_de_login_real_tambem_falha`.
   *Nota metodológica:* como este bug inviabilizaria a caracterização de tudo que
   vem depois do login (RBAC, itens, empréstimos, periféricos, histórico,
   relatórios — o objetivo central deste módulo), os demais arquivos de teste usam
   `app.dependency_overrides` (fixtures `client_gestor`/`client_tecnico`/
   `client_aprendiz` em `conftest.py`) para simular um usuário já autenticado,
   contornando especificamente o passo quebrado (decodificar um JWT real) sem
   esconder o bug — que tem sua própria caracterização direta, sem contorno, nos
   testes citados acima.

2. **`format_date()` não reformata datas que já vêm com hora embutida em string.**
   `app/db/utils.py::format_date()` só reconhece objetos `datetime`/`date` ou
   strings `"YYYY-MM-DD"`; uma string `"YYYY-MM-DD HH:MM:SS"` falha no
   `strptime` interno e é devolvida **sem alteração**. Isso afeta diretamente
   `InventoryDBManager.issue()`, que monta a mensagem de erro "data de empréstimo
   anterior ao cadastro" chamando `format_date(str(item['date_registered']))` — o
   `str()` prévio já inclui a hora, então o usuário vê a data crua do MySQL (ex.:
   `"2026-06-01 09:00:00"`) em vez de `"01/06/2026"`.
   Testes: `test_unit_formatacao.py::TestFormatDate::test_string_iso_com_hora_nao_e_reformatada`,
   `test_loan_workflow.py::TestIniciarEmprestimoTransicoesInvalidas::test_data_de_emprestimo_anterior_ao_cadastro`.

3. **Mensagem de erro técnica (não amigável) quando falta o arquivo de modelo de
   uma revenda válida.** Em `generate_return_term_bytes()` (e, pela mesma
   estrutura de código, também em `generate_loan_term_bytes()`), a checagem
   amigável `"Modelo de termo [...] não encontrado para {revenda}"` só é
   alcançada quando a revenda é **desconhecida** (`TERMO_DEVOLUCAO_MODELOS.get(revenda)`
   devolve `None`). Se a revenda é válida mas o `.docx` correspondente está
   ausente/corrompido no disco, o código pula direto para `Document(modelo_path)`,
   cuja exceção é capturada de forma genérica e devolvida crua como
   `"Erro ao gerar documento: {e}"` — sem a checagem de existência do arquivo
   rodar antes. A transação no banco não é afetada (o item permanece no status
   anterior), só a mensagem ao usuário é confusa.
   Teste: `test_loan_workflow.py::TestIniciarDevolucaoTransicoesInvalidas::test_iniciar_devolucao_com_modelo_ausente_retorna_erro_tecnico_nao_amigavel`
   (só é executado quando os `.docx` realmente não existem na máquina — ver
   "Limitações conhecidas deste ambiente" acima; `generate_loan_term_bytes()` não
   tem um teste dedicado equivalente, mas o mesmo padrão de código se aplica lá).

4. **(item 16) `confirm_return()` não desfaz o vínculo em `equipment_peripherals`.**
   Ao confirmar uma devolução, o código atualiza cada periférico vinculado para
   `status='Disponível'` (correto), mas nunca executa um `DELETE` na tabela de
   vínculo. Resultado: depois da devolução confirmada, o periférico aparece
   corretamente como "Disponível" em `/api/peripherals`, mas **continua**
   aparecendo como vinculado àquele equipamento em
   `GET /api/items/{id}/peripherals` — o vínculo só é desfeito por um
   "desvincular" manual (`DELETE /api/peripherals/links/{link_id}`).
   Teste: `test_peripherals.py::TestEfeitoDoFluxoDeEmprestimoSobrePerifericos::test_confirmar_devolucao_libera_periferico_mas_nao_remove_o_vinculo`.

5. **Vincular o mesmo par equipamento/periférico duas vezes devolve um erro cru,
   não amigável.** `equipment_peripherals` tem `UNIQUE(equipment_id,
   peripheral_id)`, mas `link_peripheral_to_equipment()` não trata essa violação
   como `add_peripheral()` trata o `UNIQUE` do identificador (que devolve
   `"Já existe um periférico com este Identificador (Nº de Série)."`). Aqui, o
   usuário recebe a exceção do PyMySQL praticamente crua, só prefixada com
   `"Erro ao vincular: "`.
   Teste: `test_peripherals.py::TestVincularEDesvincular::test_vincular_o_mesmo_par_duas_vezes_retorna_erro_tecnico_nao_amigavel`.

6. **`remove_user()`/`update_password()` reportam sucesso mesmo quando o id não
   existe.** Nenhuma das duas funções verifica a existência do usuário antes do
   `DELETE`/`UPDATE`, nem olha `cursor.rowcount` depois — um comando que não afeta
   nenhuma linha não levanta exceção no MySQL, então ambas sempre devolvem
   `(True, "...com sucesso.")`. Um Gestor não tem como saber, pela resposta da
   API, se removeu um usuário de verdade ou se o id já não existia.
   Testes: `test_users.py::TestRemoverUsuario::test_remover_usuario_inexistente_tambem_retorna_sucesso`,
   `test_users.py::TestAtualizarSenha::test_atualizar_senha_de_usuario_inexistente_tambem_retorna_sucesso`.

7. **`format_cpf()` não valida o CPF, só formata se tiver exatamente 11 dígitos.**
   Qualquer string com uma quantidade de dígitos diferente de 11 volta sem
   qualquer alteração ou sinalização de erro — não há checagem de dígito
   verificador. Baixo risco, mas documentado porque é usado nos termos gerados.
   Teste: `test_unit_formatacao.py::TestFormatCpf::test_quantidade_de_digitos_diferente_de_11_retorna_original`.

Além da lista acima (que segue exatamente os comentários `# BUG CONHECIDO` no
código), há uma **inconsistência observada** que não foi classificada como bug por
não ter um "comportamento errado" claro, só uma falta de padronização: `GET`/`PUT
/api/items/{id}` devolvem `404 "Item não encontrado."` para um id inexistente,
enquanto `DELETE /api/items/{id}` devolve `400 "ID não encontrado."` para a mesma
situação — status code e mensagem diferentes para o mesmo tipo de erro, dependendo do
endpoint. Ver `test_items.py::TestRemoverItem::test_remover_item_inexistente_retorna_400_nao_404`.
