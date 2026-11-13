"""Regras puras da conferência do inventário (sem banco)."""
import pytest

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("texto, esperado", [
    ("https://patrimonio.revalle.com.br/e/123", ("id", 123)),
    ("httpsÇ;;patrimonio.revalle.com.br;e;123", ("id", 123)),
    ("HTTPS://X/E/45/", ("id", 45)),
    ("#77", ("id", 77)),
    (" nº 8 ", ("id", 8)),
    ("77", ("id", 77)),
    ("SN-E-1234", ("identificador", "SN-E-1234")),
    ("abc  123", ("identificador", "ABC 123")),
])
def test_resolver_codigo(texto, esperado):
    from app.services.inventario import Alvo, resolver_codigo

    assert resolver_codigo(texto) == Alvo(*esperado)


def test_codigo_vazio_levanta():
    from app.services.inventario import CodigoInvalido, resolver_codigo

    with pytest.raises(CodigoInvalido, match="Informe o código lido."):
        resolver_codigo("   ")


def test_chave_da_leitura():
    from app.services.inventario import Alvo, chave_da_leitura

    assert chave_da_leitura(Alvo("id", 5), 5) == "item:5"
    assert chave_da_leitura(Alvo("id", 99999), None) == "id:99999"
    assert chave_da_leitura(Alvo("identificador", "XYZ"), None) == "codigo:XYZ"
    assert len(chave_da_leitura(Alvo("identificador", "A" * 300), None)) == 120


def test_normalizar_texto():
    from app.services.inventario import normalizar_texto

    assert normalizar_texto("  Revalle   JUAZEIRO ") == "revalle juazeiro"
    assert normalizar_texto("Recepção") == "recepcao"
    assert normalizar_texto(None) == ""


UNIDADE = "Revalle Juazeiro"


def _e(**kw):
    from app.services.inventario import Esperado

    base = dict(item_id=1, ativo=True, revenda="Revalle Juazeiro", local_id=None,
                colaborador_id=None, usuario_nome=None)
    base.update(kw)
    return Esperado(**base)


def _en(**kw):
    from app.services.inventario import Encontrado

    base = dict(local_id=None, usuario_conferido=False, colaborador_id=None, colaborador_nome=None)
    base.update(kw)
    return Encontrado(**base)


# (id, kwargs do Esperado ou None, kwargs do Encontrado, saída)
CASOS = [
    ("sem_item", None, {}, ("nao_cadastrado", [])),
    ("inativo", {"ativo": False}, {}, ("nao_cadastrado", ["baixado"])),
    ("unidade_com_espaco_e_caixa", {"revenda": "REVALLE JUAZEIRO "}, {}, ("ok", [])),
    ("outra_unidade", {"revenda": "Petrolina"}, {}, ("outra_unidade", ["outra_unidade"])),
    ("sem_unidade", {"revenda": None}, {}, ("outra_unidade", ["sem_unidade"])),
    ("local_diferente", {"local_id": 1}, {"local_id": 2}, ("local_errado", ["local"])),
    ("item_sem_local", {"local_id": None}, {"local_id": 2}, ("ok", [])),
    ("nada_encontrado_de_local", {"local_id": 1}, {"local_id": None}, ("ok", [])),
    ("outra_unidade_com_locais_diferentes", {"revenda": "Petrolina", "local_id": 1},
     {"local_id": 2}, ("outra_unidade", ["outra_unidade"])),
    ("usuario_nao_conferido", {"colaborador_id": 5}, {"colaborador_id": 6}, ("ok", [])),
    ("colaborador_diferente", {"colaborador_id": 5},
     {"usuario_conferido": True, "colaborador_id": 6}, ("usuario_diferente", ["usuario"])),
    ("legado_mesmo_nome", {"colaborador_id": None, "usuario_nome": "Fulano de Tal"},
     {"usuario_conferido": True, "colaborador_id": 9, "colaborador_nome": "FULANO DE TAL"},
     ("ok", [])),
    ("legado_outro_nome", {"colaborador_id": None, "usuario_nome": "Fulano de Tal"},
     {"usuario_conferido": True, "colaborador_id": 9, "colaborador_nome": "Beltrano"},
     ("usuario_diferente", ["usuario"])),
    ("legado_ninguem", {"colaborador_id": None, "usuario_nome": "Fulano de Tal"},
     {"usuario_conferido": True}, ("usuario_diferente", ["usuario"])),
    ("ninguem_esperado_ninguem_achado", {}, {"usuario_conferido": True}, ("ok", [])),
    ("ninguem_esperado_alguem_achado", {},
     {"usuario_conferido": True, "colaborador_id": 7, "colaborador_nome": "X"},
     ("usuario_diferente", ["usuario"])),
    ("usuario_e_local", {"colaborador_id": 5, "local_id": 1},
     {"usuario_conferido": True, "colaborador_id": 6, "local_id": 2},
     ("usuario_diferente", ["usuario", "local"])),
]


@pytest.mark.parametrize("esperado_kw, encontrado_kw, saida", [c[1:] for c in CASOS], ids=[c[0] for c in CASOS])
def test_classificar(esperado_kw, encontrado_kw, saida):
    from app.services.inventario import classificar

    esperado = None if esperado_kw is None else _e(**esperado_kw)
    assert classificar(esperado, UNIDADE, _en(**encontrado_kw)) == saida


def test_montar_detalhe():
    from app.services.inventario import montar_detalhe

    vazio = dict(revenda=None, local_esperado_nome=None, local_encontrado_nome=None,
                 usuario_esperado=None, usuario_encontrado=None)
    assert montar_detalhe(
        ["usuario", "local"],
        **{**vazio, "usuario_esperado": "Fulano", "local_esperado_nome": "Sala TI",
           "local_encontrado_nome": "Recepção"},
    ) == (
        "Usuário cadastrado: Fulano · encontrado com: ninguém. "
        "Local cadastrado: Sala TI · encontrado em: Recepção."
    )
    assert montar_detalhe(["outra_unidade"], **{**vazio, "revenda": "Juazeiro"}) == "Cadastrado na unidade JUAZEIRO."
    assert montar_detalhe(["baixado"], **vazio) == "Equipamento baixado no sistema."
    assert montar_detalhe(["sem_unidade"], **vazio) == "Equipamento sem unidade cadastrada."
    assert montar_detalhe([], **vazio) is None


def test_rotulos_e_ordem():
    from app.services.inventario import ORDEM_CATEGORIAS, ROTULO_AINDA_NAO_LIDO, ROTULOS_CATEGORIA

    assert ORDEM_CATEGORIAS[0] == "nao_encontrado" and ORDEM_CATEGORIAS[-1] == "ok"
    assert set(ORDEM_CATEGORIAS) == set(ROTULOS_CATEGORIA)
    assert ROTULOS_CATEGORIA["ok"] == "Conferido"
    assert ROTULO_AINDA_NAO_LIDO == "Ainda não lido"
