"""Schemas, validações e migração 0500 das linhas telefônicas."""
from decimal import Decimal

import pytest

from app.schemas.linhas import (
    LinhaUpdate,
    formatar_numero_linha,
    normalizar_iccid,
    normalizar_numero_linha,
    validar_custo,
    validar_operadora,
)

MSG_TAMANHO = "informe DDD + 8 ou 9 dígitos."


@pytest.mark.unit
@pytest.mark.parametrize(
    "entrada, esperado",
    [
        ("(74) 99999-0001", "74999990001"),
        ("+55 74 99999-0001", "74999990001"),
        ("5574999990001", "74999990001"),
        ("74 99999 0001", "74999990001"),
        ("(74) 3611-0000", "7436110000"),
    ],
)
def test_normalizar_numero_linha(entrada, esperado):
    assert normalizar_numero_linha(entrada) == esperado


@pytest.mark.unit
@pytest.mark.parametrize(
    "entrada, mensagem",
    [
        ("7499999", MSG_TAMANHO),
        ("", MSG_TAMANHO),
        (None, MSG_TAMANHO),
        ("20999990001", "DDD inexistente."),
        ("23999990001", "DDD inexistente."),
        ("74899990001", "celular com 9 dígitos deve começar com 9."),
        ("7499999999", "Celular deve ter 11 dígitos (DDD + 9 + número)."),
        ("(74) 6999-9999", "Celular deve ter 11 dígitos (DDD + 9 + número)."),
    ],
)
def test_normalizar_numero_linha_erros(entrada, mensagem):
    with pytest.raises(ValueError) as exc:
        normalizar_numero_linha(entrada)
    assert str(exc.value) == mensagem


@pytest.mark.unit
def test_formatar_numero_linha():
    assert formatar_numero_linha("74999990001") == "(74) 99999-0001"
    assert formatar_numero_linha("7436110000") == "(74) 3611-0000"
    assert formatar_numero_linha("123") == "123"
    assert formatar_numero_linha(None) == ""


@pytest.mark.unit
def test_ddds_validos_tem_67_codigos():
    from app.schemas.linhas import DDDS_VALIDOS

    assert len(DDDS_VALIDOS) == 67
    assert {"11", "74", "99", "21"} <= DDDS_VALIDOS
    assert not {"10", "20", "23", "29", "36", "50", "52", "56", "70", "72", "76", "78", "90"} & DDDS_VALIDOS


@pytest.mark.unit
def test_validar_custo():
    assert validar_custo("49.9") == Decimal("49.90")
    assert validar_custo(0) == Decimal("0.00")
    assert validar_custo(49.9) == Decimal("49.90")
    for invalido, mensagem in [
        ("-1", "não pode ser negativo."),
        ("49,90", "valor inválido."),
        ("1.234", "use no máximo 2 casas decimais."),
        ("100000000", "valor acima do permitido."),
        (None, "informe o valor."),
        ("", "informe o valor."),
        ("NaN", "valor inválido."),
    ]:
        with pytest.raises(ValueError) as exc:
            validar_custo(invalido)
        assert str(exc.value) == mensagem


@pytest.mark.unit
def test_validar_operadora():
    assert validar_operadora("vivo") == "Vivo"
    assert validar_operadora("tim") == "TIM"
    with pytest.raises(ValueError) as exc:
        validar_operadora("Oi")
    assert str(exc.value) == "escolha Vivo, Claro, TIM, Algar ou Outra."


@pytest.mark.unit
def test_normalizar_iccid():
    assert normalizar_iccid("8955 0000 0000 0000 0001") == "89550000000000000001"
    assert normalizar_iccid("") is None
    assert normalizar_iccid(None) is None
    with pytest.raises(ValueError) as exc:
        normalizar_iccid("123")
    assert str(exc.value) == "informe de 18 a 22 dígitos."


@pytest.mark.unit
def test_linha_update_null_em_campo_obrigatorio():
    with pytest.raises(ValueError, match="não pode ficar vazio."):
        LinhaUpdate.model_validate({"numero": None})
    for campo in ("operadora", "custo_mensal", "unidade_id"):
        with pytest.raises(ValueError, match="não pode ficar vazio."):
            LinhaUpdate.model_validate({campo: None})
    assert LinhaUpdate.model_validate({"item_id": None}).model_dump(exclude_unset=True) == {"item_id": None}
    assert LinhaUpdate().model_dump(exclude_unset=True) == {}


@pytest.mark.integration
def test_migracao_0500_cria_tabelas(_test_schema):
    from app.db.database_mysql import get_cursor

    def colunas(tabela):
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT column_name AS nome, column_type AS tipo FROM information_schema.columns "
                "WHERE table_schema = DATABASE() AND table_name = %s",
                (tabela,),
            )
            return {r["nome"]: r["tipo"] for r in cur.fetchall()}

    linhas = colunas("linhas")
    assert {
        "numero", "operadora", "plano", "custo_mensal", "status", "unidade_id", "item_id",
        "colaborador_id", "iccid", "observacoes", "criada_em", "atualizada_em",
    } <= set(linhas)
    assert linhas["custo_mensal"].lower() == "decimal(10,2)"
    eventos = colunas("linha_eventos")
    assert {
        "linha_id", "tipo", "campo", "de_valor", "para_valor", "item_id", "item_id_anterior",
        "ocorrido_em", "por_usuario", "observacao",
    } <= set(eventos)


class _CursorQueQuebra:
    """get_cursor falso: levanta o erro do MySQL ao entrar, como um deadlock na primeira instrução."""

    def __init__(self, codigo, tentativas):
        self.codigo, self.tentativas = codigo, tentativas

    def __call__(self, *args, **kwargs):
        import pymysql

        self.tentativas.append(1)
        raise pymysql.err.OperationalError(self.codigo, "bloqueio simulado")


@pytest.mark.unit
@pytest.mark.parametrize("codigo", [1213, 1205])
def test_bloqueio_do_banco_tenta_de_novo_uma_vez_e_vira_409(monkeypatch, codigo):
    from app.db import linha_db
    from app.services.audit import AuditContext

    tentativas: list = []
    monkeypatch.setattr(linha_db, "get_cursor", _CursorQueQuebra(codigo, tentativas))
    ctx = AuditContext(usuario_id=None, username="teste", papel="Gestor", ip=None)
    dados = {"numero": "74999990001", "operadora": "Vivo", "custo_mensal": Decimal("1.00"), "unidade_id": 1}

    with pytest.raises(linha_db.LinhaConflito) as exc:
        linha_db.LinhaDBManager().criar(dados, ctx)
    assert str(exc.value) == "Outra alteração estava em andamento. Tente novamente."
    assert len(tentativas) == 2

    tentativas.clear()
    with pytest.raises(linha_db.LinhaConflito):
        linha_db.LinhaDBManager().atualizar(1, {"plano": "X"}, ctx)
    assert len(tentativas) == 2

    tentativas.clear()
    with pytest.raises(linha_db.LinhaConflito):
        linha_db.LinhaDBManager().alterar_status(1, "suspensa", ctx)
    assert len(tentativas) == 2


@pytest.mark.unit
def test_bloqueio_na_primeira_tentativa_e_sucesso_na_segunda():
    import pymysql
    from app.db.linha_db import _tolerante_a_bloqueio

    chamadas = []

    @_tolerante_a_bloqueio
    def gravar():
        chamadas.append(1)
        if len(chamadas) == 1:
            raise pymysql.err.OperationalError(1213, "deadlock")
        return "ok"

    assert gravar() == "ok"
    assert len(chamadas) == 2


@pytest.mark.unit
def test_outro_erro_do_mysql_nao_e_repetido():
    import pymysql
    from app.db.linha_db import _tolerante_a_bloqueio

    chamadas = []

    @_tolerante_a_bloqueio
    def gravar():
        chamadas.append(1)
        raise pymysql.err.OperationalError(1045, "acesso negado")

    with pytest.raises(pymysql.err.OperationalError):
        gravar()
    assert len(chamadas) == 1
