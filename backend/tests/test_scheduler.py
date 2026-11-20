"""Testes do agendador de tarefas (app/services/scheduler.py) e dos jobs (app/jobs.py).

Horários guardados em `scheduler_execucoes.ultima_execucao` são UTC (DATETIME sem fuso).
`America/Sao_Paulo` é UTC-3 o ano todo (sem horário de verão desde 2019), então
05:00 em São Paulo = 08:00 UTC.
"""
import asyncio
from datetime import datetime, timedelta, timezone

import pymysql
import pytest

PREFIXO = "t7_"


def _utc(*args) -> datetime:
    return datetime(*args, tzinfo=timezone.utc)


@pytest.fixture
def agendador(_test_schema):
    """Registro de jobs vazio e linhas de teste de `scheduler_execucoes` limpas
    (essa tabela não é truncada pelo `limpar_banco`)."""
    from app.db.database_mysql import get_cursor
    from app.services import scheduler

    def limpar():
        scheduler.limpar_jobs()
        with get_cursor() as cur:
            cur.execute("DELETE FROM scheduler_execucoes WHERE nome LIKE %s", (PREFIXO + "%",))

    limpar()
    yield scheduler
    limpar()


def _estado(nome):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT * FROM scheduler_execucoes WHERE nome=%s", (nome,))
        return cur.fetchone()


def _conexao_crua(cfg):
    return pymysql.connect(
        host=cfg.host, port=cfg.port, user=cfg.user, password=cfg.password,
        database=cfg.name, charset="utf8mb4",
    )


# ── registro ───────────────────────────────────────────────────────────────

@pytest.mark.unit
def test_registrar_job_valida_argumentos():
    from app.services import scheduler

    scheduler.limpar_jobs()
    try:
        with pytest.raises(ValueError):
            scheduler.registrar_job("x", lambda: None)
        with pytest.raises(ValueError):
            scheduler.registrar_job("x", lambda: None, a_cada_segundos=5, diariamente_as="05:00")
        with pytest.raises(ValueError):
            scheduler.registrar_job("x", lambda: None, a_cada_segundos=0)
        for ruim in ("5h", "25:00", "05:60", "05:00:00", ""):
            with pytest.raises(ValueError):
                scheduler.registrar_job("x", lambda: None, diariamente_as=ruim)
        assert scheduler.jobs_registrados() == {}
        scheduler.registrar_job("x", lambda: None, diariamente_as="05:00")
        assert set(scheduler.jobs_registrados()) == {"x"}
    finally:
        scheduler.limpar_jobs()


# ── job_vencido (pura) ─────────────────────────────────────────────────────

@pytest.mark.unit
def test_job_vencido_intervalo():
    from app.services.scheduler import Job, job_vencido

    job = Job("x", lambda: None, a_cada_segundos=60)
    agora = _utc(2026, 1, 10, 12, 0, 0)

    assert job_vencido(job, None, agora) is True
    assert job_vencido(job, agora - timedelta(seconds=59), agora) is False
    assert job_vencido(job, agora - timedelta(seconds=60), agora) is True
    # `ultima` vem do MySQL sem fuso (UTC): vale o mesmo que com fuso.
    assert job_vencido(job, datetime(2026, 1, 10, 11, 59, 30), agora) is False
    assert job_vencido(job, datetime(2026, 1, 10, 11, 58, 0), agora) is True


@pytest.mark.unit
def test_job_vencido_diario():
    from app.services.scheduler import Job, job_vencido

    job = Job("x", lambda: None, diariamente_as="05:00")  # 05:00 em SP = 08:00 UTC

    # Antes do horário de hoje: não vence (mesmo nunca tendo rodado).
    assert job_vencido(job, None, _utc(2026, 1, 10, 7, 59)) is False
    # Depois do horário e nunca executado: vence.
    assert job_vencido(job, None, _utc(2026, 1, 10, 8, 0)) is True
    assert job_vencido(job, None, _utc(2026, 1, 10, 20, 0)) is True
    # Já executado hoje (depois do horário): não vence.
    assert job_vencido(job, datetime(2026, 1, 10, 8, 1), _utc(2026, 1, 10, 20, 0)) is False
    # Executado ontem: vence depois do horário de hoje, não antes.
    ontem = datetime(2026, 1, 9, 8, 5)
    assert job_vencido(job, ontem, _utc(2026, 1, 10, 7, 59)) is False
    assert job_vencido(job, ontem, _utc(2026, 1, 10, 8, 0)) is True
    # Fuso: 02:30 UTC de 10/01 ainda é 23:30 de 09/01 em SP, e o "hoje" é 09/01.
    assert job_vencido(job, None, _utc(2026, 1, 10, 2, 30)) is True  # nunca rodou, 05:00 de 09/01 já passou
    assert job_vencido(job, datetime(2026, 1, 9, 8, 5), _utc(2026, 1, 10, 2, 30)) is False  # já rodou "hoje"
    # `agora` sem fuso também é tratado como UTC.
    assert job_vencido(job, None, datetime(2026, 1, 10, 8, 0)) is True


# ── execução ───────────────────────────────────────────────────────────────

@pytest.mark.integration
def test_executar_pendentes_roda_uma_vez_e_grava(agendador):
    chamadas = []
    agendador.registrar_job(PREFIXO + "a", lambda: chamadas.append(1), a_cada_segundos=3600)
    agora = _utc(2026, 1, 10, 12, 0)

    assert agendador.executar_pendentes(agora) == [PREFIXO + "a"]
    assert agendador.executar_pendentes(agora + timedelta(seconds=30)) == []
    assert chamadas == [1]

    linha = _estado(PREFIXO + "a")
    assert linha["ultimo_status"] == "ok"
    assert linha["ultimo_erro"] is None
    assert linha["ultima_execucao"] == datetime(2026, 1, 10, 12, 0)  # UTC, sem fuso

    # Passado o intervalo, roda de novo.
    assert agendador.executar_pendentes(agora + timedelta(hours=1)) == [PREFIXO + "a"]
    assert chamadas == [1, 1]


@pytest.mark.integration
def test_erro_no_job_registra_e_continua(agendador):
    chamadas = []

    def quebra():
        raise RuntimeError("falha com segredo de conexao")

    agendador.registrar_job(PREFIXO + "a_quebra", quebra, a_cada_segundos=60)
    agendador.registrar_job(PREFIXO + "b_ok", lambda: chamadas.append(1), a_cada_segundos=60)

    executados = agendador.executar_pendentes(_utc(2026, 1, 10, 12, 0))

    assert set(executados) == {PREFIXO + "a_quebra", PREFIXO + "b_ok"}
    assert chamadas == [1]
    com_erro = _estado(PREFIXO + "a_quebra")
    assert com_erro["ultimo_status"] == "erro"
    assert "RuntimeError" in com_erro["ultimo_erro"]
    assert "segredo" not in com_erro["ultimo_erro"]  # só o tipo: o texto do driver pode vazar host/usuário
    assert _estado(PREFIXO + "b_ok")["ultimo_status"] == "ok"


@pytest.mark.integration
def test_job_falhou_grava_a_mensagem_sem_traceback(agendador, caplog):
    from app.services.scheduler import JobFalhou

    def quebra():
        raise JobFalhou("Base de colaboradores indisponível no momento.")

    agendador.registrar_job(PREFIXO + "indisp", quebra, a_cada_segundos=60)
    with caplog.at_level("WARNING"):
        agendador.executar_pendentes(_utc(2026, 1, 10, 12, 0))

    linha = _estado(PREFIXO + "indisp")
    assert linha["ultimo_status"] == "erro"
    assert linha["ultimo_erro"] == "Base de colaboradores indisponível no momento."
    assert all(r.exc_info is None for r in caplog.records)


@pytest.mark.integration
def test_lock_impede_execucao_simultanea(agendador, _test_schema):
    from app.core import config

    chamadas = []
    agendador.registrar_job(PREFIXO + "x", lambda: chamadas.append(1), a_cada_segundos=60)
    nome_lock = f"patrimonio_job_{PREFIXO}x:{config.settings.DB_NAME}"[:64]
    assert agendador.nome_do_lock(PREFIXO + "x") == nome_lock

    conn = _conexao_crua(_test_schema)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT GET_LOCK(%s, 0)", (nome_lock,))
            assert cur.fetchone()[0] == 1

        assert agendador.executar_pendentes(_utc(2026, 1, 10, 12, 0)) == []
        assert chamadas == []
        assert _estado(PREFIXO + "x") is None

        with conn.cursor() as cur:
            cur.execute("SELECT RELEASE_LOCK(%s)", (nome_lock,))
    finally:
        conn.close()

    assert agendador.executar_pendentes(_utc(2026, 1, 10, 12, 0)) == [PREFIXO + "x"]
    assert chamadas == [1]


@pytest.mark.unit
def test_nome_do_lock_limitado_a_64_caracteres():
    from app.services import scheduler

    assert len(scheduler.nome_do_lock("n" * 100)) == 64


@pytest.mark.integration
def test_lock_e_liberado_depois_da_execucao(agendador, _test_schema):
    """Depois de rodar (mesmo com erro) o lock volta a ficar livre."""
    agendador.registrar_job(PREFIXO + "l", lambda: 1 / 0, a_cada_segundos=60)
    agendador.executar_pendentes(_utc(2026, 1, 10, 12, 0))

    conn = _conexao_crua(_test_schema)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT GET_LOCK(%s, 0)", (agendador.nome_do_lock(PREFIXO + "l"),))
            assert cur.fetchone()[0] == 1
    finally:
        conn.close()


@pytest.mark.integration
def test_ultima_execucao_relida_dentro_do_lock(agendador, monkeypatch):
    """Outro worker pode ter executado entre a leitura e o lock: não roda duas vezes."""
    from app.db.database_mysql import get_cursor

    chamadas = []
    agendador.registrar_job(PREFIXO + "r", lambda: chamadas.append(1), a_cada_segundos=3600)
    agora = _utc(2026, 1, 10, 12, 0)

    original = agendador._ler_estados

    def le_e_depois_outro_worker_grava():
        estados = original()
        with get_cursor() as cur:
            cur.execute(
                "INSERT INTO scheduler_execucoes (nome, ultima_execucao, ultimo_status) "
                "VALUES (%s, %s, 'ok')",
                (PREFIXO + "r", datetime(2026, 1, 10, 11, 59, 0)),
            )
        return estados

    monkeypatch.setattr(agendador, "_ler_estados", le_e_depois_outro_worker_grava)

    assert agendador.executar_pendentes(agora) == []
    assert chamadas == []


@pytest.mark.integration
def test_diario_so_roda_depois_do_horario_e_uma_vez_por_dia(agendador):
    chamadas = []
    agendador.registrar_job(PREFIXO + "d", lambda: chamadas.append(1), diariamente_as="05:00")

    assert agendador.executar_pendentes(_utc(2026, 1, 10, 7, 0)) == []  # 04:00 em SP
    assert agendador.executar_pendentes(_utc(2026, 1, 10, 8, 1)) == [PREFIXO + "d"]
    assert agendador.executar_pendentes(_utc(2026, 1, 10, 20, 0)) == []
    assert agendador.executar_pendentes(_utc(2026, 1, 11, 8, 1)) == [PREFIXO + "d"]
    assert chamadas == [1, 1]


# ── laço assíncrono ────────────────────────────────────────────────────────

@pytest.mark.unit
def test_iniciar_nao_faz_nada_com_agendador_desligado(monkeypatch):
    from app.core import config
    from app.services import scheduler

    monkeypatch.setattr(config.settings, "SCHEDULER_ENABLED", False)

    async def principal():
        scheduler.iniciar()
        assert scheduler._tarefa is None
        await scheduler.parar()  # sem tarefa: não falha

    asyncio.run(principal())


@pytest.mark.unit
def test_laco_executa_tiques_sobrevive_a_erro_e_para_limpo(monkeypatch):
    from app.core import config
    from app.services import scheduler

    monkeypatch.setattr(config.settings, "SCHEDULER_ENABLED", True)
    monkeypatch.setattr(scheduler, "TIQUE_SEGUNDOS", 0.01)
    chamadas = []

    def tique():
        chamadas.append(1)
        if len(chamadas) == 1:
            raise RuntimeError("falha no tique")
        return []

    monkeypatch.setattr(scheduler, "executar_pendentes", tique)

    async def principal():
        scheduler.iniciar()
        tarefa = scheduler._tarefa
        assert tarefa is not None
        scheduler.iniciar()  # segunda chamada não duplica a tarefa
        assert scheduler._tarefa is tarefa
        await asyncio.sleep(0.15)
        assert not tarefa.done()  # o erro do primeiro tique não matou o laço
        await scheduler.parar()
        assert tarefa.done()
        assert scheduler._tarefa is None

    asyncio.run(principal())
    assert len(chamadas) >= 2


# ── jobs da aplicação ──────────────────────────────────────────────────────

@pytest.mark.unit
def test_registrar_jobs_so_com_base_de_rh_configurada(monkeypatch):
    from app import jobs
    from app.core import config
    from app.services import scheduler

    scheduler.limpar_jobs()
    try:
        monkeypatch.setattr(config.settings, "COLABORADORES_DATABASE_URL", "")
        jobs.registrar_jobs()
        assert set(scheduler.jobs_registrados()) == {"limpar_coletas"}

        monkeypatch.setattr(config.settings, "COLABORADORES_DATABASE_URL", "postgresql://x/y")
        monkeypatch.setattr(config.settings, "COLABORADORES_SYNC_HORA", "06:30")
        jobs.registrar_jobs()
        job = scheduler.jobs_registrados()["sincronizar_colaboradores"]
        assert job.diariamente_as == "06:30"
        assert job.a_cada_segundos is None

        jobs.registrar_jobs()  # idempotente
        assert set(scheduler.jobs_registrados()) == {"limpar_coletas", "sincronizar_colaboradores"}
    finally:
        scheduler.limpar_jobs()


@pytest.mark.unit
def test_job_sincronizar_colaboradores_trata_sincronizacao_em_andamento_como_normal(monkeypatch):
    from app import jobs
    from app.db.colaborador_db import SincronizacaoEmAndamento

    def em_andamento(db, ctx):
        raise SincronizacaoEmAndamento()

    monkeypatch.setattr(jobs, "executar_sincronizacao", em_andamento)
    monkeypatch.setattr(jobs, "ColaboradorDBManager", lambda: object())
    jobs.sincronizar_colaboradores()  # não levanta


@pytest.mark.unit
def test_job_sincronizar_colaboradores_registra_base_indisponivel(monkeypatch):
    from app import jobs
    from app.services.colaboradores_source import ColaboradoresIndisponivel
    from app.services.scheduler import JobFalhou

    def indisponivel(db, ctx):
        raise ColaboradoresIndisponivel()

    monkeypatch.setattr(jobs, "executar_sincronizacao", indisponivel)
    monkeypatch.setattr(jobs, "ColaboradorDBManager", lambda: object())
    with pytest.raises(JobFalhou) as e:
        jobs.sincronizar_colaboradores()
    assert "indisponível" in str(e.value)


@pytest.mark.unit
def test_job_sincronizar_colaboradores_usa_contexto_do_sistema(monkeypatch):
    from app import jobs

    recebidos = []
    monkeypatch.setattr(jobs, "executar_sincronizacao", lambda db, ctx: recebidos.append(ctx))
    monkeypatch.setattr(jobs, "ColaboradorDBManager", lambda: object())
    jobs.sincronizar_colaboradores()
    assert recebidos[0].username == "sistema"


# ── robustez (rodada de revisão) ───────────────────────────────────────────

class _CursorEspiao:
    """Cursor da conexão do lock: registra os comandos e quebra no INSERT do estado."""

    def __init__(self, real, comandos):
        self._real = real
        self._comandos = comandos

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self._real.close()
        return False

    def execute(self, sql, args=None):
        self._comandos.append(sql)
        if "INSERT" in sql.upper():
            raise pymysql.err.OperationalError(2013, "Lost connection to MySQL server")
        return self._real.execute(sql, args)

    def fetchone(self):
        return self._real.fetchone()


class _ConexaoEspia:
    def __init__(self, real, comandos):
        self._real = real
        self._comandos = comandos

    def cursor(self, *a, **kw):
        return _CursorEspiao(self._real.cursor(*a, **kw), self._comandos)

    def __getattr__(self, nome):
        return getattr(self._real, nome)


@pytest.mark.integration
def test_estado_e_gravado_por_outra_conexao_que_nao_a_do_lock(agendador, monkeypatch):
    """A conexão do lock só segura/solta o lock e relê o estado; se ela estivesse morta na
    hora de gravar (INSERT quebra nela), o estado ainda assim é gravado."""
    comandos = []
    original = agendador.get_connection
    monkeypatch.setattr(agendador, "get_connection", lambda: _ConexaoEspia(original(), comandos))
    chamadas = []
    agendador.registrar_job(PREFIXO + "o", lambda: chamadas.append(1), diariamente_as="05:00")

    assert agendador.executar_pendentes(_utc(2026, 1, 10, 8, 1)) == [PREFIXO + "o"]

    assert chamadas == [1]
    assert _estado(PREFIXO + "o")["ultimo_status"] == "ok"
    assert all("INSERT" not in c.upper() for c in comandos)
    assert any("GET_LOCK" in c for c in comandos) and any("RELEASE_LOCK" in c for c in comandos)
    # Estado gravado: não roda de novo no tique seguinte.
    assert agendador.executar_pendentes(_utc(2026, 1, 10, 8, 2)) == []
    assert chamadas == [1]


@pytest.mark.integration
def test_falha_ao_gravar_estado_avisa_sem_traceback_e_libera_o_lock(agendador, _test_schema, monkeypatch, caplog):
    chamadas = []
    job = agendador.Job(PREFIXO + "w", lambda: chamadas.append(1), a_cada_segundos=60)

    def sem_banco(*a, **kw):
        raise pymysql.err.OperationalError(2003, "Can't connect to MySQL server")

    monkeypatch.setattr(agendador, "get_cursor", sem_banco)
    with caplog.at_level("WARNING"):
        assert agendador._executar_com_trava(job, _utc(2026, 1, 10, 12, 0)) is True
    assert chamadas == [1]
    avisos = [r for r in caplog.records if r.levelname == "WARNING" and "gravar o estado" in r.getMessage()]
    assert len(avisos) == 1 and avisos[0].exc_info is None

    conn = _conexao_crua(_test_schema)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT GET_LOCK(%s, 0)", (agendador.nome_do_lock(job.nome),))
            assert cur.fetchone()[0] == 1  # o lock foi solto apesar da falha
    finally:
        conn.close()


@pytest.mark.integration
def test_ultima_execucao_sem_microssegundos_nao_pula_o_proximo_tique(agendador):
    """DATETIME arredonda 12:00:00.7 para 12:00:01; sem truncar, o tique de 12:01:00.2
    veria só 59,2 s e pularia o job de 60 s."""
    chamadas = []
    agendador.registrar_job(PREFIXO + "m", lambda: chamadas.append(1), a_cada_segundos=60)

    assert agendador.executar_pendentes(_utc(2026, 1, 10, 12, 0, 0, 700000)) == [PREFIXO + "m"]
    assert _estado(PREFIXO + "m")["ultima_execucao"] == datetime(2026, 1, 10, 12, 0, 0)
    assert agendador.executar_pendentes(_utc(2026, 1, 10, 12, 1, 0, 200000)) == [PREFIXO + "m"]
    assert chamadas == [1, 1]


@pytest.mark.unit
def test_banco_fora_do_ar_vira_aviso_de_uma_linha_e_outros_erros_levam_traceback(monkeypatch, caplog):
    from app.core import config
    from app.services import scheduler

    monkeypatch.setattr(config.settings, "SCHEDULER_ENABLED", True)
    monkeypatch.setattr(scheduler, "TIQUE_SEGUNDOS", 0.01)
    erros = [
        pymysql.err.OperationalError(2003, "Can't connect"),
        pymysql.err.InterfaceError(0, ""),
        RuntimeError("bug"),
    ]

    def tique():
        if erros:
            raise erros.pop(0)
        return []

    monkeypatch.setattr(scheduler, "executar_pendentes", tique)

    async def principal():
        scheduler.iniciar()
        await asyncio.sleep(0.2)
        await scheduler.parar()

    with caplog.at_level("WARNING"):
        asyncio.run(principal())

    registros = [r for r in caplog.records if "tique do agendador" in r.getMessage()]
    assert len(registros) == 3
    assert [r.levelname for r in registros] == ["WARNING", "WARNING", "ERROR"]
    assert registros[0].exc_info is None and registros[1].exc_info is None
    assert registros[2].exc_info is not None
