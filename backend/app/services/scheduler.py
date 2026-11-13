"""Agendador de tarefas periódicas, seguro para vários workers.

Cada worker da API roda o mesmo laço (uma `asyncio.Task` com tique de 60 s). Para que uma
tarefa não rode em dobro, antes de executar cada job o worker pega um `GET_LOCK` do MySQL
(nome por job e por database), relê a última execução já com o lock e só então roda. O lock
fica numa conexão própria (fora do pool) que é fechada ao final: se o processo morrer no meio,
o MySQL libera o lock sozinho.

Estado em `scheduler_execucoes`. `ultima_execucao` é gravada em UTC (DATETIME sem fuso) e
lida de volta como UTC. Jobs diários usam o horário de `America/Sao_Paulo`.

Falhas nunca derrubam o laço: erro de job fica registrado na tabela e no log; erro de
tique (banco fora do ar, por exemplo) é logado e o próximo tique tenta de novo.
"""
import asyncio
import logging
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Callable, Optional
from zoneinfo import ZoneInfo

from app.core import config as _config
from app.db.database_mysql import get_connection, get_cursor

logger = logging.getLogger(__name__)

FUSO = ZoneInfo("America/Sao_Paulo")
TIQUE_SEGUNDOS: float = 60

_PADRAO_HORA = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")
_TAMANHO_MAX_ERRO = 500


class JobFalhou(Exception):
    """Falha esperada de um job: a mensagem (já amigável) vai para `ultimo_erro` e para o
    log como aviso, sem traceback."""


@dataclass(frozen=True)
class Job:
    nome: str
    func: Callable[[], None]
    a_cada_segundos: Optional[int] = None
    diariamente_as: Optional[str] = None  # "HH:MM", fuso America/Sao_Paulo

    @property
    def horario(self) -> tuple[int, int]:
        h, m = self.diariamente_as.split(":")
        return int(h), int(m)


_jobs: dict[str, Job] = {}


def registrar_job(
    nome: str,
    func: Callable[[], None],
    *,
    a_cada_segundos: Optional[int] = None,
    diariamente_as: Optional[str] = None,
) -> None:
    """Registra (ou substitui) um job. Exatamente um entre `a_cada_segundos` e `diariamente_as`."""
    if (a_cada_segundos is None) == (diariamente_as is None):
        raise ValueError("Informe exatamente um entre a_cada_segundos e diariamente_as.")
    if a_cada_segundos is not None and a_cada_segundos <= 0:
        raise ValueError("a_cada_segundos deve ser maior que zero.")
    if diariamente_as is not None and not _PADRAO_HORA.match(diariamente_as):
        raise ValueError(f"diariamente_as deve estar no formato HH:MM (recebido {diariamente_as!r}).")
    _jobs[nome] = Job(nome, func, a_cada_segundos, diariamente_as)


def jobs_registrados() -> dict[str, Job]:
    return dict(_jobs)


def limpar_jobs() -> None:
    """Esvazia o registro (uso em testes)."""
    _jobs.clear()


def nome_do_lock(nome_do_job: str) -> str:
    """Nome do `GET_LOCK` do job. Vale para o servidor MySQL inteiro, então leva o nome do
    database (vários ambientes/suítes dividem o servidor). Limite do MySQL: 64 caracteres."""
    return f"patrimonio_job_{nome_do_job}:{_config.settings.DB_NAME}"[:64]


def _em_utc(dt: datetime) -> datetime:
    """Datetime com fuso UTC. Sem fuso (como vem do MySQL) já é UTC."""
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def job_vencido(job: Job, ultima: Optional[datetime], agora: datetime) -> bool:
    """Se o job deve rodar agora. Função pura; `ultima` e `agora` sem fuso valem como UTC.

    - Intervalo: nunca rodou, ou já passou o intervalo desde a última.
    - Diário: hoje (no fuso de São Paulo) já passou do horário e a última execução é anterior
      ao horário de hoje.
    """
    agora = _em_utc(agora)
    ultima = _em_utc(ultima) if ultima is not None else None
    if job.a_cada_segundos is not None:
        return ultima is None or agora - ultima >= timedelta(seconds=job.a_cada_segundos)
    local = agora.astimezone(FUSO)
    hora, minuto = job.horario
    horario_de_hoje = local.replace(hour=hora, minute=minuto, second=0, microsecond=0)
    if local < horario_de_hoje:
        return False
    return ultima is None or ultima < horario_de_hoje


def _ler_estados() -> dict[str, Optional[datetime]]:
    """`ultima_execucao` (UTC sem fuso) de cada job gravado."""
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT nome, ultima_execucao FROM scheduler_execucoes")
        return {r["nome"]: r["ultima_execucao"] for r in cur.fetchall()}


def _rodar(job: Job) -> tuple[str, Optional[str]]:
    """Executa o job e devolve (status, erro). Nunca levanta."""
    try:
        job.func()
    except JobFalhou as e:
        logger.warning("Job %s falhou: %s", job.nome, e)
        return "erro", str(e)[:_TAMANHO_MAX_ERRO]
    except Exception as e:
        logger.exception("Job %s falhou com erro inesperado.", job.nome)
        # Só o tipo vai para a tabela: o texto de um erro de driver pode trazer host/usuário.
        return "erro", f"Erro inesperado ({type(e).__name__})."[:_TAMANHO_MAX_ERRO]
    return "ok", None


def _executar_com_trava(job: Job, agora: datetime) -> bool:
    """Roda o job se conseguir o lock e ele ainda estiver vencido. True se rodou."""
    conn = get_connection()  # conexão própria: fechá-la libera o lock mesmo se algo falhar
    try:
        conn.autocommit(True)
        lock = nome_do_lock(job.nome)
        with conn.cursor() as cur:
            cur.execute("SELECT GET_LOCK(%s, 0) AS ok", (lock,))
            if cur.fetchone()["ok"] != 1:
                return False  # outro worker está executando este job
            try:
                # Relê já com o lock: outro worker pode ter terminado entre a leitura e o lock.
                cur.execute(
                    "SELECT ultima_execucao FROM scheduler_execucoes WHERE nome=%s", (job.nome,)
                )
                linha = cur.fetchone()
                if not job_vencido(job, linha["ultima_execucao"] if linha else None, agora):
                    return False
                status, erro = _rodar(job)
                cur.execute(
                    "INSERT INTO scheduler_execucoes (nome, ultima_execucao, ultimo_status, ultimo_erro) "
                    "VALUES (%s, %s, %s, %s) "
                    "ON DUPLICATE KEY UPDATE ultima_execucao=VALUES(ultima_execucao), "
                    "ultimo_status=VALUES(ultimo_status), ultimo_erro=VALUES(ultimo_erro)",
                    (job.nome, _em_utc(agora).replace(tzinfo=None), status, erro),
                )
                return True
            finally:
                try:
                    cur.execute("SELECT RELEASE_LOCK(%s)", (lock,))
                except Exception:
                    logger.exception("Falha ao liberar o lock do job %s.", job.nome)
    finally:
        conn.close()


def executar_pendentes(agora: Optional[datetime] = None) -> list[str]:
    """Roda os jobs vencidos e devolve os nomes dos que executaram neste worker.

    A falha de um job (ou da infraestrutura para ele) não interrompe os demais.
    """
    agora = _em_utc(agora) if agora is not None else datetime.now(timezone.utc)
    jobs = list(_jobs.values())
    if not jobs:
        return []
    estados = _ler_estados()
    executados: list[str] = []
    for job in jobs:
        if not job_vencido(job, estados.get(job.nome), agora):
            continue
        try:
            if _executar_com_trava(job, agora):
                executados.append(job.nome)
        except Exception:
            logger.exception("Falha ao executar o job %s.", job.nome)
    return executados


_tarefa: Optional["asyncio.Task[None]"] = None


async def _laco() -> None:
    while True:
        try:
            await asyncio.to_thread(executar_pendentes)
        except asyncio.CancelledError:
            raise
        except Exception:
            logger.exception("Falha no tique do agendador.")
        await asyncio.sleep(TIQUE_SEGUNDOS)


def iniciar() -> None:
    """Inicia o laço do agendador no event loop atual. Não faz nada com
    `SCHEDULER_ENABLED=false` ou se já estiver iniciado."""
    global _tarefa
    if not _config.settings.SCHEDULER_ENABLED:
        logger.info("Agendador desligado (SCHEDULER_ENABLED=false).")
        return
    if _tarefa is not None and not _tarefa.done():
        return
    _tarefa = asyncio.get_running_loop().create_task(_laco(), name="agendador")
    logger.info("Agendador iniciado com %d job(s).", len(_jobs))


async def parar() -> None:
    """Cancela o laço e espera ele encerrar."""
    global _tarefa
    tarefa, _tarefa = _tarefa, None
    if tarefa is None:
        return
    tarefa.cancel()
    try:
        await tarefa
    except asyncio.CancelledError:
        pass
    except Exception:
        logger.exception("Erro ao encerrar o agendador.")
