"""Deadlock do MySQL (1213) nas escritas do módulo de acessos vira 409, não 500."""
import pymysql
import pytest

pytestmark = pytest.mark.unit


def test_deadlock_vira_conflito_com_mensagem_em_portugues():
    from app.db.acesso_db import MSG_DEADLOCK, Conflito, mapear_deadlock

    @mapear_deadlock
    def escrita():
        raise pymysql.err.OperationalError(1213, "Deadlock found when trying to get lock")

    with pytest.raises(Conflito) as exc:
        escrita()
    assert str(exc.value) == "Outra alteração estava em andamento. Tente novamente."
    assert exc.value.status_code == 409
    assert MSG_DEADLOCK == str(exc.value)


def test_outros_erros_do_banco_passam_sem_mudar():
    from app.db.acesso_db import mapear_deadlock

    @mapear_deadlock
    def escrita():
        raise pymysql.err.OperationalError(2003, "Can't connect")

    with pytest.raises(pymysql.err.OperationalError):
        escrita()
