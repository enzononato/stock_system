"""
Fixtures compartilhadas da suíte de testes de caracterização (Módulo 7a).

CONTEXTO IMPORTANTE
--------------------
Este arquivo é o ponto mais sensível de toda a suíte: `backend/.env` (quando existe na
máquina do desenvolvedor) aponta para o MySQL de **produção**. Qualquer import descuidado
de `app.main`/`app.db.*` antes de reconfigurar as variáveis de ambiente pode disparar uma
conexão real de produção — porque `InventoryDBManager.__init__()` chama `_create_tables()`
imediatamente, e cada router (`items.py`, `loans.py`, `peripherals.py`, `history.py`,
`reports.py`, `documents.py`) instancia `InventoryDBManager()` **no nível de módulo**, ou
seja, um simples `from app.main import app` já é suficiente para abrir conexões reais.

Por isso, as regras deste arquivo são:

1. NUNCA importar `app.main`, `app.routers.*`, `app.db.inventory_manager_db` ou
   `app.db.user_manager_db` no nível de módulo (topo do arquivo). Todo import desses
   módulos acontece **dentro** do corpo de uma fixture, depois que `_ambiente_de_teste`
   já validou e sobrescreveu as variáveis DB_* do processo.
2. A fixture `_ambiente_de_teste` é a ÚNICA que lê `TEST_DB_*` e escreve em `os.environ`.
   Ela recusa (`pytest.skip`) rodar sem `TEST_DB_HOST`/`TEST_DB_NAME` explícitos, e aborta
   a sessão inteira (`pytest.exit`) se esses valores coincidirem com os de produção.
3. Módulos puramente utilitários (`app.core.config`, `app.core.security`) não fazem I/O ao
   serem importados (só leem `Settings`, que é só validação Pydantic) — por isso os testes
   `unit` podem importá-los livremente, sem depender do guard.

Se o MySQL de teste não estiver acessível (Docker não subiu, credenciais erradas, ou —
no caso deste worktree específico — o módulo `app/db/database_mysql.py` ainda não existir
porque outro agente está criando o novo context manager de conexão em paralelo), a suíte
inteira deve pular com uma mensagem clara, nunca "explodir" com um erro de import.
"""
import os
import socket
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from uuid import uuid4

import pymysql
import pymysql.cursors
import pytest

# ────────────────────────────────────────────────────────────────────────────
# Configuração do ambiente de teste (guard de segurança)
# ────────────────────────────────────────────────────────────────────────────

# Valores conhecidos de PRODUÇÃO (ver backend/app/core/config.py e backend/.env.example).
# Se TEST_DB_HOST/TEST_DB_NAME baterem com qualquer um destes, abortamos a sessão.
_PROD_HOST_MARKERS = {"72.61.53.20"}
_PROD_DB_NAME_MARKERS = {"stock_sys_db"}


def _markers_de_producao() -> tuple[set, set]:
    """Reúne valores considerados 'produção' para a checagem de segurança.

    Além dos defaults hardcoded em app/core/config.py, também tenta ler um eventual
    backend/.env real (não versionado, presente apenas na máquina do desenvolvedor) para
    pegar host/nome reais, caso tenham sido trocados dos defaults. Falha silenciosamente
    (best effort) — a ausência do arquivo não deve quebrar a suíte.
    """
    hosts = set(_PROD_HOST_MARKERS)
    names = set(_PROD_DB_NAME_MARKERS)
    try:
        env_path = Path(__file__).resolve().parent.parent / ".env"
        if env_path.exists():
            for linha in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
                linha = linha.strip()
                if linha.startswith("DB_HOST="):
                    hosts.add(linha.split("=", 1)[1].strip())
                elif linha.startswith("DB_NAME="):
                    names.add(linha.split("=", 1)[1].strip())
    except Exception:
        pass
    return hosts, names


@dataclass(frozen=True)
class _TestDBConfig:
    host: str
    port: int
    user: str
    password: str
    name: str


def pytest_configure(config):
    """Reconfigura o ambiente ANTES da coleta — este hook é o único ponto seguro.

    `app.core.config.settings` é um singleton criado no primeiro import de
    `app.core.config`. Arquivos de teste unitário importam `settings` no nível de
    módulo, e imports de módulo acontecem na COLETA, antes de qualquer fixture.
    Se as variáveis só fossem sobrescritas numa fixture (como era antes), o
    singleton já teria sido construído a partir do backend/.env real e a camada
    da aplicação apontaria para PRODUÇÃO durante a suíte inteira — enquanto a
    limpeza entre testes, que usa conexão crua, zeraria o banco de teste. Foi
    exatamente isso que aconteceu: os testes escreviam em produção.

    Sobrescrever aqui garante que qualquer import posterior de `app.core.config`
    já enxergue o banco de teste. A verificação final de que isso funcionou está
    em `_assegurar_app_apontando_para_teste`.
    """
    host = (os.environ.get("TEST_DB_HOST") or "").strip()
    name = (os.environ.get("TEST_DB_NAME") or "").strip()
    if not host or not name:
        return  # sem configuração de teste, os testes de integração pulam adiante

    prod_hosts, prod_names = _markers_de_producao()
    if host in prod_hosts or name in prod_names:
        return  # o guard de _ambiente_de_teste aborta a sessão com mensagem clara

    os.environ["DB_HOST"] = host
    os.environ["DB_PORT"] = os.environ.get("TEST_DB_PORT", "3307")
    os.environ["DB_USER"] = os.environ.get("TEST_DB_USER", "stock_test")
    os.environ["DB_PASSWORD"] = os.environ.get("TEST_DB_PASSWORD", "stock_test_pw")
    os.environ["DB_NAME"] = name
    os.environ["DB_CHARSET"] = "utf8mb4"
    os.environ["LOGIN_RATE_LIMIT"] = "10000/minute"


def _assegurar_app_apontando_para_teste(cfg: "_TestDBConfig") -> None:
    """Confere para onde a APLICAÇÃO realmente vai conectar, e não apenas o que
    pedimos via variáveis de ambiente.

    Validar `TEST_DB_*` não basta: o que decide o destino das escritas é o
    `settings` já materializado dentro do processo. Se por qualquer motivo ele
    tiver sido construído antes da nossa reconfiguração, esta checagem aborta a
    sessão inteira em vez de deixar a suíte gravar no banco errado.
    """
    from app.core.config import settings

    if settings.DB_HOST != cfg.host or settings.DB_NAME != cfg.name:
        pytest.exit(
            "ABORTADO POR SEGURANÇA: a aplicação está configurada para "
            f"{settings.DB_HOST}/{settings.DB_NAME}, mas os testes exigem "
            f"{cfg.host}/{cfg.name}. O singleton app.core.config.settings foi "
            "construído antes da reconfiguração do ambiente (algum import de "
            "app.* aconteceu cedo demais). Rodar assim gravaria no banco errado.",
            returncode=1,
        )


@pytest.fixture(scope="session")
def _ambiente_de_teste(tmp_path_factory):
    """Valida TEST_DB_* e, se seguro, reconfigura o ambiente do processo.

    - Sem TEST_DB_HOST/TEST_DB_NAME  -> pytest.skip (mensagem clara, nada quebra).
    - TEST_DB_HOST/TEST_DB_NAME == produção -> pytest.exit (aborta a sessão inteira).
    - Caso contrário -> sobrescreve DB_HOST/DB_USER/DB_PASSWORD/DB_NAME/DB_CHARSET/DB_PORT
      e STORAGE_BACKEND/LOCAL_STORAGE_PATH em os.environ, para que Settings() (lida por
      app.core.config no primeiro import) enxergue apenas o banco de teste.
    """
    host = (os.environ.get("TEST_DB_HOST") or "").strip()
    name = (os.environ.get("TEST_DB_NAME") or "").strip()

    if not host or not name:
        pytest.skip(
            "TEST_DB_HOST e/ou TEST_DB_NAME não definidos. Configure as variáveis de ambiente "
            "TEST_DB_HOST, TEST_DB_PORT (padrão 3307), TEST_DB_USER, TEST_DB_PASSWORD e "
            "TEST_DB_NAME apontando para o MySQL de teste (suba com "
            "`docker compose -f docker-compose.test.yml up -d`) antes de rodar os testes de "
            "integração. Veja docs/TESTES.md."
        )

    prod_hosts, prod_names = _markers_de_producao()
    if host in prod_hosts or name in prod_names:
        pytest.exit(
            "ABORTADO POR SEGURANÇA: TEST_DB_HOST/TEST_DB_NAME coincidem com valores "
            f"conhecidos de PRODUÇÃO (host={host!r}, name={name!r}). Os testes se recusam a "
            "rodar contra esse banco para não arriscar escrever em produção. Configure um "
            "banco de teste dedicado (docker-compose.test.yml) e nunca reaproveite as "
            "credenciais de produção em TEST_DB_*.",
            returncode=1,
        )

    try:
        port = int(os.environ.get("TEST_DB_PORT", "3307") or 3307)
    except ValueError:
        pytest.skip(f"TEST_DB_PORT inválido: {os.environ.get('TEST_DB_PORT')!r} não é um inteiro.")

    user = os.environ.get("TEST_DB_USER", "stock_test")
    password = os.environ.get("TEST_DB_PASSWORD", "stock_test_pw")

    # A partir daqui é seguro sobrescrever o que Settings()/get_connection() vão enxergar.
    os.environ["DB_HOST"] = host
    os.environ["DB_PORT"] = str(port)  # best-effort: só usado se o context manager novo ler DB_PORT
    os.environ["DB_USER"] = user
    os.environ["DB_PASSWORD"] = password
    os.environ["DB_NAME"] = name
    os.environ["DB_CHARSET"] = "utf8mb4"
    # As fixtures autenticam fazendo login de verdade, e vários testes logam
    # múltiplas vezes. Com o limite de produção (5/min por IP, e o TestClient
    # usa sempre o mesmo "testclient"), a suíte passaria a receber 429 a partir
    # da sexta chamada — falha de infraestrutura de teste, não de comportamento.
    # O limite em si é testado explicitamente em test_auth.py, que sobrescreve
    # este valor no escopo dele.
    os.environ["LOGIN_RATE_LIMIT"] = "10000/minute"
    os.environ["STORAGE_BACKEND"] = "local"
    os.environ["LOCAL_STORAGE_PATH"] = str(tmp_path_factory.mktemp("storage"))

    return _TestDBConfig(host=host, port=port, user=user, password=password, name=name)


def _conectar_direto(cfg: _TestDBConfig, *, database: str = None, timeout: int = 5):
    """Conexão pymysql crua, independente da camada de app — usada apenas pela infra de
    testes (setup de schema, truncamento) para nunca depender do código que está sendo
    refatorado em paralelo por outro módulo."""
    return pymysql.connect(
        host=cfg.host,
        port=cfg.port,
        user=cfg.user,
        password=cfg.password,
        database=database,
        charset="utf8mb4",
        connect_timeout=timeout,
        cursorclass=pymysql.cursors.DictCursor,
    )


# ────────────────────────────────────────────────────────────────────────────
# Schema do banco de teste
# ────────────────────────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def _test_schema(_ambiente_de_teste):
    """Cria o schema do zero no banco de teste (uma vez por sessão).

    - Cria o banco (se não existir) e a tabela `usuarios` via SQL direto, pois ela não é
      criada por InventoryDBManager._create_tables().
    - Reaproveita InventoryDBManager._create_tables() (chamada automaticamente pelo
      __init__) para criar items/history/peripherals/equipment_peripherals, exatamente
      como o app real faz — se esse método mudar na refatoração, o teste acompanha.
    """
    cfg = _ambiente_de_teste

    try:
        conn = _conectar_direto(cfg)
    except Exception as e:
        pytest.skip(
            f"Não foi possível conectar ao MySQL de teste em {cfg.host}:{cfg.port} ({e}). "
            "Suba o banco com `docker compose -f docker-compose.test.yml up -d` (ver "
            "docs/TESTES.md) e confirme as variáveis TEST_DB_*."
        )
    try:
        with conn.cursor() as cur:
            cur.execute(f"CREATE DATABASE IF NOT EXISTS `{cfg.name}` CHARACTER SET utf8mb4")
        conn.commit()
    finally:
        conn.close()

    try:
        conn = _conectar_direto(cfg, database=cfg.name)
        try:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS usuarios (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        username VARCHAR(100) NOT NULL UNIQUE,
                        password VARCHAR(255) NOT NULL,
                        role VARCHAR(50) NOT NULL
                    )
                    """
                )
            conn.commit()
        finally:
            conn.close()
    except Exception as e:
        pytest.skip(f"Não foi possível criar a tabela 'usuarios' no banco de teste: {e}")

    # Última linha de defesa, antes de qualquer escrita pela camada da aplicação:
    # confirma que o `settings` materializado no processo aponta para o banco de
    # teste. Se apontar para outro lugar, aborta a sessão inteira.
    _assegurar_app_apontando_para_teste(cfg)

    try:
        from app.db.inventory_manager_db import InventoryDBManager

        InventoryDBManager()  # __init__ chama _create_tables() (items/history/peripherals/eq_peripherals)
    except Exception as e:
        pytest.skip(
            "Não foi possível importar/instanciar "
            f"app.db.inventory_manager_db.InventoryDBManager: {e}. Verifique se "
            "app/db/database_mysql.py existe (pode ainda não ter sido mesclado de outro "
            "módulo da refatoração) e se o MySQL de teste está acessível."
        )

    return cfg


@pytest.fixture
def limpar_banco(_test_schema):
    """Trunca todas as tabelas antes de cada teste, para isolamento total entre testes.

    Usa FOREIGN_KEY_CHECKS=0 ao redor do TRUNCATE em vez de uma ordem manual de tabelas —
    mais robusto a mudanças de schema e igualmente seguro dentro de uma única conexão.
    """
    cfg = _test_schema
    conn = _conectar_direto(cfg, database=cfg.name)
    try:
        with conn.cursor() as cur:
            cur.execute("SET FOREIGN_KEY_CHECKS=0")
            for tabela in ("equipment_peripherals", "history", "items", "peripherals", "usuarios"):
                cur.execute(f"TRUNCATE TABLE `{tabela}`")
            cur.execute("SET FOREIGN_KEY_CHECKS=1")
        conn.commit()
    finally:
        conn.close()
    yield


# ────────────────────────────────────────────────────────────────────────────
# App FastAPI / cliente HTTP
# ────────────────────────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def _fastapi_app(_test_schema):
    try:
        from app.main import app as fastapi_app
    except Exception as e:
        pytest.skip(f"Não foi possível importar app.main: {e}")
    return fastapi_app


@pytest.fixture
def client(_fastapi_app, limpar_banco):
    """Cliente HTTP (TestClient/httpx) sem autenticação, contra um banco limpo."""
    from fastapi.testclient import TestClient

    with TestClient(_fastapi_app) as c:
        yield c


# ────────────────────────────────────────────────────────────────────────────
# Camada de dados (para fixtures de fábrica) — instâncias "cruas" reaproveitando
# os managers reais, não INSERTs soltos.
# ────────────────────────────────────────────────────────────────────────────

@pytest.fixture
def inv_manager(_test_schema, limpar_banco):
    from app.db.inventory_manager_db import InventoryDBManager

    return InventoryDBManager()


@pytest.fixture
def user_manager(_test_schema, limpar_banco):
    from app.db.user_manager_db import UserDBManager

    return UserDBManager()


# ────────────────────────────────────────────────────────────────────────────
# Fábricas de itens / periféricos
# ────────────────────────────────────────────────────────────────────────────

@pytest.fixture
def criar_item(inv_manager):
    """Fábrica de itens via InventoryDBManager.add_item (método real da camada de dados)."""

    def _criar(**overrides):
        dados = {
            "tipo": "Notebook",
            "brand": "Dell",
            "model": "Latitude 5420",
            "identificador": f"SN-{uuid4().hex[:8].upper()}",
            "revenda": "Revalle Juazeiro",
            "setor": "TI",
            "date_registered": datetime.now(),
        }
        dados.update(overrides)
        ok, resultado = inv_manager.add_item(dados, "fixture_teste")
        assert ok, f"Falha ao criar item de teste via add_item: {resultado}"
        return inv_manager.find(resultado)

    return _criar


@pytest.fixture
def item_disponivel(criar_item):
    """Item recém-cadastrado, status 'Disponível' (o default da coluna)."""
    item = criar_item()
    assert item["status"] == "Disponível"
    return item


@pytest.fixture
def item_pendente(inv_manager, item_disponivel):
    """Item com empréstimo iniciado (issue), status 'Pendente'."""
    ok, msg = inv_manager.issue(
        item_disponivel["id"],
        "Fulano de Tal",
        "11122233344",
        "101 - Puxada",
        "Analista",
        "TI",
        "Revalle Juazeiro",
        datetime.now().strftime("%d/%m/%Y"),
        "fixture_teste",
    )
    assert ok, msg
    item = inv_manager.find(item_disponivel["id"])
    assert item["status"] == "Pendente"
    return item


@pytest.fixture
def item_indisponivel(inv_manager, item_pendente):
    """Item com empréstimo confirmado (confirm_loan), status 'Indisponível'."""
    ok, msg = inv_manager.confirm_loan(item_pendente["id"], "fixture_teste", "termos_assinados/fixture.pdf")
    assert ok, msg
    item = inv_manager.find(item_pendente["id"])
    assert item["status"] == "Indisponível"
    return item


@pytest.fixture
def criar_periferico(inv_manager):
    """Fábrica de periféricos via InventoryDBManager.add_peripheral (método real)."""

    def _criar(**overrides):
        dados = {
            "tipo": "Mouse",
            "brand": "Logitech",
            "model": "M90",
            "identificador": f"PERIF-{uuid4().hex[:8].upper()}",
        }
        dados.update(overrides)
        ok, msg = inv_manager.add_peripheral(dados, "fixture_teste")
        assert ok, f"Falha ao criar periférico de teste via add_peripheral: {msg}"
        encontrados = [
            p for p in inv_manager.list_peripherals(include_inactive=True)
            if p["identificador"] == dados["identificador"]
        ]
        assert encontrados, "Periférico criado não encontrado por identificador."
        return encontrados[0]

    return _criar


@pytest.fixture
def periferico_disponivel(criar_periferico):
    perif = criar_periferico()
    assert perif["status"] == "Disponível"
    return perif


@pytest.fixture
def forcar_devolucao_iniciada(_test_schema):
    """Utilitário de teste (NÃO é o código de produção): replica só a parte de banco de
    dados de `InventoryDBManager.generate_return_term_bytes()` (linhas finais, depois da
    geração do .docx): status -> 'Pendente Devolução' + INSERT do history 'Devolução'.

    Existe porque este ambiente de teste não tem os arquivos .docx em backend/modelos/
    (não fazem parte da posse deste módulo — ver docs/TESTES.md) e `generate_return_term_bytes`
    retorna erro *antes* de tocar o banco quando o modelo não pode ser aberto. Isso bloquearia
    a caracterização de tudo que vem depois (confirm_return, estorno de Devolução, relatório
    mensal). A geração do .docx em si É caracterizada separadamente (com skip claro quando o
    modelo não está disponível) em test_loan_workflow.py.
    """
    cfg = _test_schema

    def _forcar(item_id: int, logged_user: str = "fixture_teste"):
        conn = _conectar_direto(cfg, database=cfg.name)
        try:
            with conn.cursor() as cur:
                cur.execute("UPDATE items SET status='Pendente Devolução' WHERE id=%s", (item_id,))
                cur.execute(
                    "INSERT INTO history (item_id, operador, data_operacao, operation) "
                    "VALUES (%s,%s,%s,'Devolução')",
                    (item_id, logged_user, datetime.now()),
                )
            conn.commit()
        finally:
            conn.close()

    return _forcar


@pytest.fixture
def termo_devolucao_disponivel():
    """True se o .docx modelo de devolução para 'Revalle Juazeiro' existir neste ambiente."""
    from app.core.config import TERMO_DEVOLUCAO_MODELOS

    caminho = TERMO_DEVOLUCAO_MODELOS.get("Revalle Juazeiro")
    return bool(caminho and os.path.exists(caminho))


@pytest.fixture
def termo_emprestimo_disponivel():
    """True se o .docx modelo de termo de responsabilidade para 'Revalle Juazeiro' existir."""
    from app.core.config import TERMO_MODELOS

    caminho = TERMO_MODELOS.get("Revalle Juazeiro")
    return bool(caminho and os.path.exists(caminho))


# ────────────────────────────────────────────────────────────────────────────
# Usuários / tokens / clientes autenticados por role
# ────────────────────────────────────────────────────────────────────────────

@pytest.fixture
def criar_usuario(user_manager):
    def _criar(username: str, password: str, role: str):
        ok, msg = user_manager.add_user(username, password, role)
        assert ok, f"Falha ao criar usuário de teste '{username}': {msg}"
        return user_manager.get_user_by_username(username)

    return _criar


@pytest.fixture
def senha_padrao_teste():
    return "SenhaForte#123"


@pytest.fixture
def usuario_gestor(criar_usuario, senha_padrao_teste):
    return criar_usuario("gestor.teste", senha_padrao_teste, "Gestor")


@pytest.fixture
def usuario_tecnico(criar_usuario, senha_padrao_teste):
    return criar_usuario("tecnico.teste", senha_padrao_teste, "Técnico")


@pytest.fixture
def usuario_aprendiz(criar_usuario, senha_padrao_teste):
    return criar_usuario("aprendiz.teste", senha_padrao_teste, "Jovem Aprendiz")


# ────────────────────────────────────────────────────────────────────────────
# Histórico (ver docs/TESTES.md, seção "Bugs corrigidos com teste de regressão"):
# este módulo encontrou, caracterizou e reportou um bug crítico de autenticação —
# app/routers/auth.py montava o claim "sub" do JWT com o id INTEIRO do usuário, e
# python-jose exige por padrão (verify_sub=True) que "sub" seja uma string, então
# TODO token de acesso emitido pelo login era rejeitado por decode_token() com 401
# "Token inválido ou expirado.". O achado foi confirmado de forma independente pelo
# Módulo 3, e já estava corrigido no código-alvo desta refatoração (login grava
# "sub" como string; get_current_user faz int(payload["sub"])) no momento em que
# isso foi reconciliado com o coordenador. Por isso os fixtures abaixo usam tokens
# REAIS obtidos via POST /api/auth/login (sem nenhum contorno/override) — o que
# agora exercita de verdade get_current_user, decode_token() e a checagem de
# expiração em todo teste de RBAC/workflow, exatamente como deveria. A demonstração
# do bug já corrigido e o teste de regressão que impede sua reintrodução ficam em
# test_auth.py.
# ────────────────────────────────────────────────────────────────────────────

def _login_e_obter_access_token(client, username: str, password: str) -> str:
    resp = client.post("/api/auth/login", json={"username": username, "password": password})
    assert resp.status_code == 200, f"Login de teste falhou para '{username}': {resp.text}"
    return resp.json()["access_token"]


@pytest.fixture
def token_gestor(client, usuario_gestor, senha_padrao_teste):
    return _login_e_obter_access_token(client, usuario_gestor["username"], senha_padrao_teste)


@pytest.fixture
def token_tecnico(client, usuario_tecnico, senha_padrao_teste):
    return _login_e_obter_access_token(client, usuario_tecnico["username"], senha_padrao_teste)


@pytest.fixture
def token_aprendiz(client, usuario_aprendiz, senha_padrao_teste):
    return _login_e_obter_access_token(client, usuario_aprendiz["username"], senha_padrao_teste)


@pytest.fixture
def client_gestor(client, token_gestor):
    """Cliente HTTP autenticado como Gestor com um token JWT real (via login de verdade)."""
    client.headers.update({"Authorization": f"Bearer {token_gestor}"})
    return client


@pytest.fixture
def client_tecnico(client, token_tecnico):
    """Cliente HTTP autenticado como Técnico com um token JWT real (via login de verdade)."""
    client.headers.update({"Authorization": f"Bearer {token_tecnico}"})
    return client


@pytest.fixture
def client_aprendiz(client, token_aprendiz):
    """Cliente HTTP autenticado como Jovem Aprendiz com um token JWT real (via login de verdade)."""
    client.headers.update({"Authorization": f"Bearer {token_aprendiz}"})
    return client
