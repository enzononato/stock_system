# CI (Módulo 8)

Este documento descreve o workflow `.github/workflows/ci.yml` e como rodar cada
job localmente antes de abrir um PR.

## Visão geral

O workflow dispara em `push` para `main` e em `pull_request`, com dois jobs
independentes (rodam em paralelo, um não depende do outro):

| Job        | O que faz                                                                 |
|------------|-----------------------------------------------------------------------------|
| `frontend` | `npm ci` → `tsc --noEmit` → `vitest run` → `vite build`                     |
| `backend`  | Sobe um MySQL 8.0.40 efêmero como *service container* → `pip install` → `pytest` |

Não há um terceiro job de lint/typecheck do backend — ver
[Próximos passos](#próximos-passos-lint-typecheck-do-backend) abaixo.

## Job `frontend`

Roda inteiramente dentro de `frontend/` (`working-directory` fixado no job).

### Rodar localmente

```bash
cd frontend
npm ci

npx tsc --noEmit     # checagem de tipos, inclusive dos arquivos *.test.ts(x)
npm run test:run     # Vitest, modo não-interativo (o mesmo do CI)
npm run build        # tsc + vite build
```

Outros scripts úteis que o CI não usa, mas ajudam no dia a dia:

```bash
npm run test           # Vitest em modo watch
npm run test:coverage  # Vitest + relatório de cobertura (v8), pasta coverage/
```

### Infraestrutura de teste

- **Vitest** + **@testing-library/react**/**user-event** + **jsdom**, configurados em
  `frontend/vitest.config.ts` — que reaproveita `frontend/vite.config.ts` via
  `mergeConfig` (mesmo alias `@` → `./src` usado pela aplicação).
- **MSW** (`msw/node`) simula toda chamada HTTP; nenhum teste de frontend abre uma
  conexão de rede real. O servidor MSW compartilhado vive em `src/test/server.ts`
  e é ligado/resetado/desligado em `src/test/setup.ts` (`onUnhandledRequest:
  'error'` — uma rota sem handler é erro de teste, não uma chamada real escapando).
- Handlers "de fábrica" (`src/test/handlers.ts`) cobrem `GET /api/constants`,
  `/api/items` e `/api/history` com respostas vazias; qualquer teste que precise
  de um corpo/status específico sobrescreve com `server.use(...)`.
- `src/test/render.tsx` expõe `renderWithClient`/`renderHookWithClient`, que já
  encapsulam um `QueryClientProvider` de teste (sem retry, sem cache entre
  montagens) para componentes/hooks que usam `@tanstack/react-query`.

## Job `backend`

O MySQL de teste é um *service container* do próprio job — nunca um host
externo, e nunca o MySQL de produção (`72.61.53.20` / `stock_sys_db`).
As credenciais/nome do banco no workflow espelham `docker-compose.test.yml` (raiz
do repo), usado para o mesmo propósito localmente.

**Versão do MySQL fixada de propósito** (`mysql:8.0.40`, não `mysql:8`): a tag
flutuante `mysql:8` passou a resolver para a série 8.4, que removeu a opção
`default-authentication-plugin` — o servidor aborta na inicialização e o job
quebraria sozinho, sem nenhuma mudança no repositório.

### Rodar localmente

Com Docker instalado:

```bash
# Sobe o MySQL de teste efêmero (dados em tmpfs — tudo é perdido ao parar)
docker compose -f docker-compose.test.yml up -d
docker compose -f docker-compose.test.yml ps   # espera ficar "healthy"

cd backend
python -m venv .venv && source .venv/bin/activate   # ou o equivalente no Windows
pip install -r requirements-dev.txt

export TEST_DB_HOST=127.0.0.1
export TEST_DB_PORT=3307        # porta publicada por docker-compose.test.yml
export TEST_DB_USER=stock_test
export TEST_DB_PASSWORD=stock_test_pw
export TEST_DB_NAME=stock_sys_test

python -m pytest -m "not mudanca_esperada"

docker compose -f docker-compose.test.yml down  # descarta os dados de teste
```

Veja `docs/TESTES.md` para o guia completo da suíte (o guard de segurança contra
rodar por engano contra produção, os markers `unit`/`integration`/
`mudanca_esperada`, e a lista de bugs já documentados).

**Por que `-m "not mudanca_esperada"` e não a suíte inteira:** os testes marcados
`mudanca_esperada` afirmam de propósito um comportamento antigo com correção já
em andamento em paralelo por outro módulo — eles devem passar a **falhar**
assim que essa correção for mesclada, e isso não é uma regressão. Rodar todos
eles no CI de rotina só adicionaria ruído; a seção correspondente de
`docs/TESTES.md` explica quando e como reagir a essas falhas esperadas.

### Diferença do que roda localmente

A porta do MySQL no `docker-compose.test.yml` é `3307` (para não colidir com um
MySQL local em `3306`); no CI, o *service container* publica `3306:3306`
diretamente (não há conflito possível na VM efêmera do runner). Fora essa porta,
as credenciais e o nome do banco são idênticos nos dois ambientes.

## Próximos passos: lint/typecheck do backend

Este repositório não tem, hoje, nenhuma ferramenta de lint/typecheck configurada
para o backend (sem `ruff`/`flake8`/`black`/`mypy`/`pylint` em
`requirements*.txt`, `pyproject.toml`, `.flake8` ou similar). Por isso o workflow
não inclui um terceiro job para isso — adicionar uma configuração nova não
pedida por outro módulo seria inventar um padrão de código sem acordo prévio da
equipe.

Quando o time decidir adotar uma ferramenta, o padrão recomendado é:

```yaml
  lint-backend:
    name: Backend — lint/typecheck
    runs-on: ubuntu-latest
    defaults:
      run:
        working-directory: backend
    steps:
      - uses: actions/checkout@v4.2.2
      - uses: actions/setup-python@v5.3.0
        with:
          python-version: '3.11.9'
          cache: pip
          cache-dependency-path: requirements-dev.txt
      - run: pip install -r requirements-dev.txt
      - run: ruff check .        # ou flake8 / mypy app, conforme o que for adotado
```

## Cache e versões fixas

- **Node** (`20.18.1`) e **Python** (`3.11.9`) fixados em versão exata, não só a
  major — a mesma lição do `mysql:8.0.40` acima: uma tag flutuante pode mudar de
  comportamento entre execuções sem nenhuma mudança no repositório.
- **Actions** (`actions/checkout`, `actions/setup-node`, `actions/setup-python`)
  fixadas em uma versão exata (`vX.Y.Z`), não em `@main`/apenas `@v4`.
- **Cache de dependências**: `setup-node` com `cache: npm` (chave por
  `frontend/package-lock.json`) e `setup-python` com `cache: pip` (chave por
  `backend/requirements-dev.txt`) — reduz o tempo de instalação nas execuções
  subsequentes sem exigir um passo de cache manual (`actions/cache`).
