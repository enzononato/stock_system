"""Camada de dados das coletas de máquina (`coletas_maquina`)."""
from datetime import datetime, timedelta, timezone

import pytest

from app.db.coleta_db import ColetaDBManager
from tests.fixtures_auth import ida_e_volta_migracao


def _utc(**delta) -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None) + timedelta(**delta)


@pytest.fixture
def coletas():
    return ColetaDBManager()


@pytest.fixture
def tecnico_id(usuario_tecnico):
    return usuario_tecnico["id"]


def test_migracao_1000_ida_e_volta(executar_sql):
    ida_e_volta_migracao(1000)
    # a tabela existe depois (a consulta falharia se não existisse)
    executar_sql("SELECT 1 FROM coletas_maquina LIMIT 1")


def test_registrar_dados_uma_vez_so(coletas, tecnico_id):
    h = "a" * 64
    coleta_id = coletas.criar(h, tecnico_id, _utc(minutes=15))

    assert coletas.registrar_dados(h, {"brand": "Dell"}, "10.0.0.5") is True
    assert coletas.registrar_dados(h, {"brand": "HP"}, "10.0.0.6") is False

    coleta = coletas.por_id(coleta_id)
    assert coleta["dados"] == {"brand": "Dell"}
    assert coleta["recebido_ip"] == "10.0.0.5"
    assert coleta["recebido_em"] is not None
    assert coletas.por_hash(h)["id"] == coleta_id


def test_registrar_dados_vencido(coletas, tecnico_id):
    h = "b" * 64
    coleta_id = coletas.criar(h, tecnico_id, _utc(minutes=-1))

    assert coletas.registrar_dados(h, {"brand": "Dell"}, "10.0.0.5") is False
    coleta = coletas.por_id(coleta_id)
    assert coleta["dados"] is None
    assert coleta["recebido_em"] is None


def test_registrar_dados_codigo_inexistente(coletas):
    assert coletas.registrar_dados("c" * 64, {}, "10.0.0.5") is False
    assert coletas.por_hash("c" * 64) is None
    assert coletas.por_id(999999) is None


def test_correspondencias_por_serial_mac_anydesk(coletas, criar_item):
    item = criar_item(
        identificador="ABC123 ",
        endereco_fisico="aa-bb-cc-dd-ee-ff",
        anydesk="1 234 567 890",
    )
    outro = criar_item(identificador="OUTRO-1")

    achados = coletas.correspondencias("abc123", [], None)
    assert [(a["id"], a["motivos"]) for a in achados] == [(item["id"], ["serial"])]

    achados = coletas.correspondencias(None, ["AABBCCDDEEFF"], None)
    assert [(a["id"], a["motivos"]) for a in achados] == [(item["id"], ["mac"])]

    achados = coletas.correspondencias(None, [], "1234567890")
    assert [(a["id"], a["motivos"]) for a in achados] == [(item["id"], ["anydesk"])]

    # um item que casa por dois critérios aparece uma vez, com os dois motivos
    achados = coletas.correspondencias("ABC123", ["11:22:33:44:55:66", "AA:BB:CC:DD:EE:FF"], None)
    assert len(achados) == 1
    assert achados[0]["id"] == item["id"]
    assert achados[0]["motivos"] == ["serial", "mac"]

    achados = coletas.correspondencias("ABC123", ["AABBCCDDEEFF"], "123 456 7890")
    assert [(a["id"], a["motivos"]) for a in achados] == [(item["id"], ["serial", "mac", "anydesk"])]

    primeiro = achados[0]
    assert primeiro["ativo"] is True
    assert primeiro["tipo"] == "Notebook"
    assert primeiro["brand"] == "Dell"
    assert primeiro["model"] == "Latitude 5420"
    assert primeiro["status"] == "Disponível"
    assert primeiro["revenda"] == "Revalle Juazeiro"
    assert outro["id"] not in [a["id"] for a in achados]


def test_correspondencias_mac_tambem_no_campo_mac(coletas, criar_item):
    item = criar_item(mac="aa:bb:cc:00:11:22")
    achados = coletas.correspondencias(None, ["AA-BB-CC-00-11-22"], None)
    assert [(a["id"], a["motivos"]) for a in achados] == [(item["id"], ["mac"])]


def test_correspondencias_ignora_vazios(coletas, criar_item):
    criar_item(identificador="", endereco_fisico="", mac="", anydesk="")
    assert coletas.correspondencias(None, [], None) == []
    assert coletas.correspondencias("", [""], "") == []
    assert coletas.correspondencias("  ", ["::", "--"], "abc") == []


def test_correspondencias_item_baixado(coletas, criar_item, executar_sql):
    baixado = criar_item(identificador="SN-BAIXADO")
    ativo = criar_item(identificador="SN-BAIXADO")
    executar_sql("UPDATE items SET is_active = 0 WHERE id = %s", (baixado["id"],))

    achados = coletas.correspondencias("SN-BAIXADO", [], None)
    # ativos primeiro; o baixado vem marcado como inativo
    assert [(a["id"], a["ativo"]) for a in achados] == [(ativo["id"], True), (baixado["id"], False)]


def test_apagar_antigas(coletas, tecnico_id, executar_sql):
    velha = coletas.criar("d" * 64, tecnico_id, _utc(minutes=15))
    nova = coletas.criar("e" * 64, tecnico_id, _utc(minutes=15))
    executar_sql(
        "UPDATE coletas_maquina SET criado_em = %s WHERE id = %s", (_utc(days=-8), velha)
    )

    assert coletas.apagar_antigas() == 1
    assert coletas.por_id(velha) is None
    assert coletas.por_id(nova) is not None
