"""Fonte de colaboradores (Postgres de RH): descoberta de colunas, normalização e leitura somente leitura."""
import dataclasses

import psycopg
import pytest

from app.services import colaboradores_source as src
from tests.fixtures_colaboradores import (
    CPF_AFASTADO,
    CPF_ATIVO,
    CPF_DESLIGADO,
    CPF_MASCARADO,
)


def test_descobrir_colunas_por_sinonimo():
    mapa = src.descobrir_colunas(["CPF", "Nome Completo", "E-mail", "Situação"])
    assert mapa == {"cpf": "CPF", "nome": "Nome Completo", "email": "E-mail", "situacao": "Situação"}


def test_override_tem_prioridade():
    mapa = src.descobrir_colunas(
        ["CPF", "Nome Completo", "Cargo", "Função"],
        override="cargo=Função, nome = Nome Completo",
    )
    assert mapa["cargo"] == "Função"
    assert mapa["nome"] == "Nome Completo"
    assert mapa["cpf"] == "CPF"


def test_sem_cpf_levanta():
    with pytest.raises(src.ColaboradoresSourceError, match="coluna não encontrada para cpf"):
        src.descobrir_colunas(["Nome Completo", "Cargo"])


def test_sem_nome_levanta():
    with pytest.raises(src.ColaboradoresSourceError, match="coluna não encontrada para nome"):
        src.descobrir_colunas(["CPF", "Cargo"])


def test_normalizar_nome_coluna():
    assert src.normalizar_nome_coluna("Matrícula") == "matricula"
    assert src.normalizar_nome_coluna("E-mail") == "e_mail"
    assert src.normalizar_nome_coluna("Nome Completo") == "nome_completo"


@pytest.mark.parametrize(
    "valor,demissao,esperado",
    [
        ("Ativo", None, "ativo"),
        (None, None, "ativo"),
        ("", None, "ativo"),
        ("Ativo", "2024-03-01", "desligado"),
        ("Ativo", "", "ativo"),
        ("Desligado", None, "desligado"),
        ("Demitido", None, "desligado"),
        ("INATIVO", None, "desligado"),
        ("Rescisão", None, "desligado"),
        (False, None, "desligado"),
        ("false", None, "desligado"),
        ("0", None, "desligado"),
        ("N", None, "desligado"),
        ("Não", None, "desligado"),
        (True, None, "ativo"),
        ("Afastado", None, "afastado"),
        ("Licença maternidade", None, "afastado"),
        ("Afastado", "2024-03-01", "desligado"),
    ],
)
def test_normalizar_situacao(valor, demissao, esperado):
    assert src.normalizar_situacao(valor, demissao) == esperado


def test_ler_todos_da_base_ficticia(base_rh):
    todos = src.ler_todos()
    assert len(todos) == 4
    por_cpf = {c.cpf: c for c in todos}
    assert set(por_cpf) == {CPF_ATIVO, CPF_DESLIGADO, CPF_AFASTADO, CPF_MASCARADO}

    assert por_cpf[CPF_ATIVO].nome == "Maria da Silva"
    assert por_cpf[CPF_ATIVO].situacao == "ativo"
    assert por_cpf[CPF_ATIVO].matricula == "1001"
    assert por_cpf[CPF_ATIVO].email == "maria@revalle.com.br"
    assert por_cpf[CPF_ATIVO].telefone == "(74) 99999-0001"
    assert por_cpf[CPF_ATIVO].gestor_nome == "Carlos Souza"
    assert por_cpf[CPF_DESLIGADO].situacao == "desligado"
    assert por_cpf[CPF_AFASTADO].situacao == "afastado"
    assert por_cpf[CPF_AFASTADO].gestor_nome is None
    assert por_cpf[CPF_MASCARADO].situacao == "ativo"  # CPF mascarado na origem normalizado

    campos = {f.name for f in dataclasses.fields(src.ColaboradorRH)}
    assert not any("nasc" in c for c in campos)
    valores = " ".join(str(v) for c in todos for v in dataclasses.asdict(c).values())
    assert "1990" not in valores and "1985" not in valores


def test_ler_todos_ignora_cpf_invalido_e_nome_vazio(base_rh):
    base_rh(
        'INSERT INTO "colaboradores-revalle" ("CPF", "Nome Completo") VALUES (%s, %s), (%s, %s)',
        ("123456789012", "Cpf Longo", "39053344705", "  "),
    )
    assert len(src.ler_todos()) == 4


def test_buscar_por_cpf(base_rh):
    achado = src.buscar_por_cpf("935.411.347-80")
    assert achado is not None and achado.cpf == CPF_MASCARADO and achado.nome == "Pedro Alves"
    assert src.buscar_por_cpf(CPF_ATIVO).nome == "Maria da Silva"
    assert src.buscar_por_cpf("00000000000") is None
    assert src.buscar_por_cpf("abc") is None


def test_sessao_somente_leitura(base_rh):
    with src.conectar() as conn:
        with pytest.raises(psycopg.errors.ReadOnlySqlTransaction):
            conn.execute('INSERT INTO "colaboradores-revalle" ("CPF", "Nome Completo") VALUES (%s, %s)', ("1", "X"))


def test_url_vazia_indisponivel(monkeypatch):
    from app.core import config

    settings = config.settings

    monkeypatch.setattr(settings, "COLABORADORES_DATABASE_URL", "")
    with pytest.raises(src.ColaboradoresIndisponivel, match="Base de colaboradores indisponível no momento."):
        src.ler_todos()
    with pytest.raises(src.ColaboradoresIndisponivel):
        src.buscar_por_cpf(CPF_ATIVO)


def test_conexao_recusada_indisponivel_sem_vazar_credenciais(monkeypatch):
    from app.core import config

    settings = config.settings

    monkeypatch.setattr(settings, "COLABORADORES_DATABASE_URL", "postgresql://usr:segredo123@127.0.0.1:1/x")
    with pytest.raises(src.ColaboradoresIndisponivel) as exc:
        src.ler_todos()
    assert "segredo123" not in str(exc.value)


def test_tabela_inexistente_levanta_erro_de_fonte(base_rh, monkeypatch):
    from app.core import config

    settings = config.settings

    monkeypatch.setattr(settings, "COLABORADORES_TABLE", "nao-existe")
    with pytest.raises(src.ColaboradoresSourceError, match="não encontrada"):
        src.ler_todos()
