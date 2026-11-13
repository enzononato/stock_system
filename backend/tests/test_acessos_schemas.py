"""Schemas do módulo de acessos: nenhum campo de senha em lugar nenhum."""
import pytest
from pydantic import ValidationError

from app.schemas import acessos
from app.schemas.acessos import (
    MSG_SEM_SENHA,
    PALAVRAS_SENHA,
    AcessoCreate,
    AcessoUpdate,
    EventoCancelar,
    EventoCreate,
    PacoteUpdate,
    SistemaCreate,
    TarefaResolver,
    chave_cargo,
)

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "classe, dados",
    [
        (SistemaCreate, {"nome": "X", "senha": "1"}),
        (AcessoCreate, {"colaborador_id": 1, "sistema_id": 1, "login": "a", "password": "x"}),
        (AcessoUpdate, {"Senha_AD": "x"}),
        (PacoteUpdate, {"cargo": "Vendedor", "sistema_ids": [], "senha_vpn": "x"}),
        (EventoCreate, {"colaborador_id": 1, "tipo": "admissao", "pwd": "x"}),
        (TarefaResolver, {"status": "feito", "login": "a", "PASSWORD": "x"}),
        (EventoCancelar, {"motivo_cancelamento": "x", "passwd": "x"}),
    ],
)
def test_campos_de_senha_recusados(classe, dados):
    with pytest.raises(ValidationError) as exc:
        classe(**dados)
    assert MSG_SEM_SENHA in str(exc.value)


def test_campo_desconhecido_recusado():
    with pytest.raises(ValidationError):
        AcessoCreate(colaborador_id=1, sistema_id=1, login="a", apelido="x")


def test_respostas_sem_campo_de_senha():
    nomes = ("SistemaResponse", "PacoteResponse", "AcessoResponse", "TarefaResponse",
             "EquipamentoResumo", "EventoResumo", "EventoDetalhe")
    for nome in nomes:
        for campo in getattr(acessos, nome).model_fields:
            assert not any(p in campo.lower() for p in PALAVRAS_SENHA), (nome, campo)


def test_login_aparado_e_obrigatorio():
    assert AcessoCreate(colaborador_id=1, sistema_id=1, login="  maria.silva ").login == "maria.silva"
    with pytest.raises(ValidationError) as exc:
        AcessoCreate(colaborador_id=1, sistema_id=1, login="   ")
    assert "Informe o login." in str(exc.value)
    with pytest.raises(ValidationError) as exc:
        AcessoCreate(colaborador_id=1, sistema_id=1, login="a" * 151)
    assert "Login muito longo" in str(exc.value)


def test_pacote_normaliza_cargo_e_remove_repetidos():
    p = PacoteUpdate(cargo="  Vendedor   Externo ", sistema_ids=[3, 1, 3])
    assert p.cargo == "Vendedor Externo"
    assert p.sistema_ids == [3, 1]
    with pytest.raises(ValidationError) as exc:
        PacoteUpdate(cargo="  ", sistema_ids=[])
    assert "Informe o cargo." in str(exc.value)


def test_chave_cargo():
    assert chave_cargo("VENDEDOR  ") == chave_cargo("vendedor") == chave_cargo("Vendédor")
    assert chave_cargo(None) == ""


def test_resolver_status_validos():
    assert TarefaResolver(status="manter").status == "manter"
    with pytest.raises(ValidationError):
        TarefaResolver(status="outro")


def test_cancelar_exige_motivo():
    with pytest.raises(ValidationError) as exc:
        EventoCancelar(motivo_cancelamento="  ")
    assert "Informe o motivo do cancelamento." in str(exc.value)


def test_manter_limita_a_observacao_para_caber_no_prefixo():
    assert TarefaResolver(status="manter", observacao="x" * 484).status == "manter"
    with pytest.raises(ValidationError) as exc:
        TarefaResolver(status="manter", observacao="x" * 485)
    assert "máximo 484 caracteres" in str(exc.value)
    assert TarefaResolver(status="nao_se_aplica", observacao="x" * 500).status == "nao_se_aplica"
