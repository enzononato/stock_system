"""Eventos de acesso: abertura, tarefas geradas e consulta (/api/acessos/eventos)."""
import pytest

from tests.acessos_helpers import (
    CPF_JOAO,
    auditoria,
    contar,
    criar_acesso,
    criar_colaborador,
    criar_sistema,
    definir_pacote,
    vincular_item,
)

pytestmark = pytest.mark.integration


def _tarefas(evento, *campos):
    return [tuple(t[c] for c in campos) for t in evento["tarefas"]]


@pytest.mark.unit
def test_diferenca_pacotes():
    from app.db.acesso_evento_db import diferenca_pacotes

    def s(i, nome):
        return {"id": i, "nome": nome, "descricao": None, "ativo": True}

    anterior = [s(1, "AD"), s(2, "E-mail"), s(5, "Promax")]
    novo = [s(1, "AD"), s(2, "E-mail"), s(3, "VPN"), s(4, "WMS")]
    entram, saem = diferenca_pacotes(anterior, novo)
    assert [x["nome"] for x in entram] == ["VPN", "WMS"]
    assert [x["nome"] for x in saem] == ["Promax"]


@pytest.mark.unit
def test_pendencias_do_evento():
    from app.db.acesso_evento_db import pendencias_do_evento

    assert pendencias_do_evento("admissao", 2, 0) == ["2 tarefas pendentes"]
    assert pendencias_do_evento("desligamento", 1, 1) == [
        "1 tarefa pendente", "1 equipamento ainda com o colaborador",
    ]
    assert pendencias_do_evento("mudanca_funcao", 0, 3) == []


def _montar_vendedor():
    ad, email = criar_sistema("AD"), criar_sistema("E-mail")
    promax = criar_sistema("Promax", ativo=False)
    definir_pacote("Vendedor", ad, email, promax)
    return ad, email, promax


def test_admissao_gera_tarefas_do_pacote(client_tecnico):
    _montar_vendedor()
    c = criar_colaborador()
    r = client_tecnico.post("/api/acessos/eventos", json={"colaborador_id": c, "tipo": "admissao"})
    assert r.status_code == 201
    ev = r.json()
    assert ev["cargo_novo"] == "Vendedor"
    assert ev["status"] == "aberto"
    assert _tarefas(ev, "sistema_nome", "acao", "status") == [
        ("AD", "criar", "pendente"), ("E-mail", "criar", "pendente"),
    ]
    assert ev["pendencias"] == ["2 tarefas pendentes"]
    assert ev["pode_concluir"] is False
    assert ev["equipamentos"] == []
    assert ev["colaborador_nome"] == "Maria Silva"
    assert ev["total_tarefas"] == 2 and ev["tarefas_pendentes"] == 2
    assert ev["aberto_por"] == "tecnico.teste"
    linhas = auditoria("acesso_evento", "criar")
    assert len(linhas) == 1
    assert linhas[0]["depois"]["tarefas"] == 2


def test_admissao_acha_pacote_com_cargo_em_outra_caixa_e_espacos(client_gestor):
    ad, email = criar_sistema("AD"), criar_sistema("E-mail")
    r = client_gestor.put(
        "/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": [ad, email]}
    )
    assert r.status_code == 200
    c = criar_colaborador(cargo="VENDEDOR  ")
    r = client_gestor.post("/api/acessos/eventos", json={"colaborador_id": c, "tipo": "admissao"})
    assert r.status_code == 201
    assert _tarefas(r.json(), "sistema_nome", "acao") == [("AD", "criar"), ("E-mail", "criar")]


def test_admissao_com_acesso_ja_ativo(client_tecnico):
    ad, email, _ = _montar_vendedor()
    c = criar_colaborador()
    criar_acesso(c, email, "maria@revalle")
    r = client_tecnico.post("/api/acessos/eventos", json={"colaborador_id": c, "tipo": "admissao"})
    ev = r.json()
    por_sistema = {t["sistema_nome"]: t for t in ev["tarefas"]}
    assert por_sistema["E-mail"]["status"] == "nao_se_aplica"
    assert por_sistema["E-mail"]["observacao"] == "Acesso já ativo: maria@revalle"
    assert por_sistema["E-mail"]["login"] == "maria@revalle"
    assert por_sistema["E-mail"]["acesso_id"] is not None
    assert por_sistema["AD"]["status"] == "pendente"
    assert ev["pendencias"] == ["1 tarefa pendente"]


def test_admissao_sem_pacote_ou_sem_cargo_422(client_tecnico):
    _montar_vendedor()
    c = criar_colaborador(cargo="Motorista")
    r = client_tecnico.post("/api/acessos/eventos", json={"colaborador_id": c, "tipo": "admissao"})
    assert r.status_code == 422
    assert r.json()["detail"] == 'O cargo "Motorista" não tem pacote de acessos definido.'

    sem_cargo = criar_colaborador("João Souza", cpf=CPF_JOAO, cargo=None)
    r = client_tecnico.post(
        "/api/acessos/eventos", json={"colaborador_id": sem_cargo, "tipo": "admissao"}
    )
    assert r.status_code == 422
    assert r.json()["detail"] == "Informe o cargo: o colaborador não tem cargo cadastrado."
    assert contar("acesso_eventos") == 0


def test_segundo_evento_aberto_409(client_tecnico):
    _montar_vendedor()
    c = criar_colaborador()
    corpo = {"colaborador_id": c, "tipo": "admissao"}
    assert client_tecnico.post("/api/acessos/eventos", json=corpo).status_code == 201
    r = client_tecnico.post("/api/acessos/eventos", json=corpo)
    assert r.status_code == 409
    assert r.json()["detail"] == "Já existe um evento de acesso aberto para este colaborador."
    assert contar("acesso_eventos") == 1


def test_colaborador_inexistente_404(client_tecnico):
    r = client_tecnico.post("/api/acessos/eventos", json={"colaborador_id": 999, "tipo": "admissao"})
    assert r.status_code == 404
    assert r.json()["detail"] == "Colaborador não encontrado."


def test_desligamento_gera_desativacoes_e_lista_equipamentos(client_tecnico, criar_item):
    ad, email, vpn = criar_sistema("AD"), criar_sistema("E-mail"), criar_sistema("VPN")
    c = criar_colaborador()
    criar_acesso(c, ad, "maria.silva")
    criar_acesso(c, vpn, "msilva")
    criar_acesso(c, email, "maria@revalle", status="desativado")
    a, b = criar_item(), criar_item()
    vincular_item(a["id"], colaborador_id=c)
    vincular_item(b["id"], cpf="111.444.777-35", status="Pendente Devolução")

    r = client_tecnico.post(
        "/api/acessos/eventos", json={"colaborador_id": c, "tipo": "desligamento"}
    )
    assert r.status_code == 201
    ev = r.json()
    assert _tarefas(ev, "sistema_nome", "acao", "login") == [
        ("AD", "desativar", "maria.silva"), ("VPN", "desativar", "msilva"),
    ]
    assert all(t["acesso_id"] is not None for t in ev["tarefas"])
    assert [e["id"] for e in ev["equipamentos"]] == [a["id"], b["id"]]
    assert ev["pendencias"] == ["2 tarefas pendentes", "2 equipamentos ainda com o colaborador"]
    assert ev["cargo_anterior"] == "Vendedor"
    assert ev["cargo_novo"] is None
    assert ev["pode_concluir"] is False


def test_desligamento_sem_nada_conclui_na_hora(client_tecnico):
    c = criar_colaborador()
    r = client_tecnico.post(
        "/api/acessos/eventos", json={"colaborador_id": c, "tipo": "desligamento"}
    )
    assert r.status_code == 201
    ev = r.json()
    assert ev["status"] == "concluido"
    assert ev["concluido_em"] is not None
    assert ev["concluido_por"] == "tecnico.teste"
    assert ev["pendencias"] == []
    assert len(auditoria("acesso_evento", "concluir")) == 1


def test_mudanca_de_funcao_gera_a_diferenca(client_tecnico):
    ad, email, promax = criar_sistema("AD"), criar_sistema("E-mail"), criar_sistema("Promax")
    vpn, wms = criar_sistema("VPN"), criar_sistema("WMS")
    definir_pacote("Vendedor", ad, email, promax)
    definir_pacote("Supervisor", ad, email, vpn, wms)
    c = criar_colaborador(cargo="Supervisor")
    criar_acesso(c, ad, "maria.silva")
    criar_acesso(c, promax, "MSILVA")

    r = client_tecnico.post(
        "/api/acessos/eventos",
        json={"colaborador_id": c, "tipo": "mudanca_funcao", "cargo_anterior": "Vendedor"},
    )
    assert r.status_code == 201
    ev = r.json()
    assert ev["cargo_novo"] == "Supervisor"
    assert ev["cargo_anterior"] == "Vendedor"
    assert _tarefas(ev, "sistema_nome", "acao") == [
        ("VPN", "criar"), ("WMS", "criar"), ("Promax", "desativar"),
    ]
    assert ev["tarefas"][2]["login"] == "MSILVA"
    assert ev["equipamentos"] == []


def test_mudanca_valida_cargos(client_tecnico):
    c = criar_colaborador(cargo="Supervisor")
    r = client_tecnico.post(
        "/api/acessos/eventos", json={"colaborador_id": c, "tipo": "mudanca_funcao"}
    )
    assert r.status_code == 422
    assert r.json()["detail"] == "Informe o cargo anterior."
    r = client_tecnico.post(
        "/api/acessos/eventos",
        json={"colaborador_id": c, "tipo": "mudanca_funcao", "cargo_anterior": "supervisor "},
    )
    assert r.status_code == 422
    assert r.json()["detail"] == "O cargo novo é igual ao anterior."
    assert contar("acesso_eventos") == 0


def test_listar_eventos_filtra_e_pagina(client_tecnico):
    ad = criar_sistema("AD")
    definir_pacote("Vendedor", ad)
    maria = criar_colaborador()
    joao = criar_colaborador("João Souza", cpf=CPF_JOAO)
    post = lambda c, tipo: client_tecnico.post(  # noqa: E731
        "/api/acessos/eventos", json={"colaborador_id": c, "tipo": tipo}
    )
    assert post(maria, "desligamento").json()["status"] == "concluido"
    assert post(maria, "admissao").status_code == 201
    assert post(joao, "admissao").status_code == 201

    r = client_tecnico.get("/api/acessos/eventos?status=aberto&limit=1")
    assert r.status_code == 200
    corpo = r.json()
    assert corpo["total"] == 2
    assert len(corpo["items"]) == 1
    assert corpo["items"][0]["total_tarefas"] == 1
    assert corpo["items"][0]["tarefas_pendentes"] == 1
    assert corpo["items"][0]["colaborador_nome"] == "João Souza"  # aberto_em DESC, id DESC

    r = client_tecnico.get("/api/acessos/eventos?busca=111.444")
    assert {e["colaborador_nome"] for e in r.json()["items"]} == {"Maria Silva"}
    assert r.json()["total"] == 2

    r = client_tecnico.get("/api/acessos/eventos?busca=joão")
    assert r.json()["total"] == 1

    r = client_tecnico.get("/api/acessos/eventos?tipo=desligamento")
    assert r.json()["total"] == 1

    r = client_tecnico.get(f"/api/acessos/eventos?colaborador_id={joao}&offset=5")
    assert r.json()["items"] == [] and r.json()["total"] == 1


def test_eventos_404_e_aprendiz_403(client_tecnico, client_aprendiz):
    r = client_tecnico.get("/api/acessos/eventos/999")
    assert r.status_code == 404
    assert r.json()["detail"] == "Evento não encontrado."
    r = client_aprendiz.get("/api/acessos/eventos")
    assert r.status_code == 403


def test_dois_eventos_abertos_ao_mesmo_tempo_so_um_vale(_fastapi_app, client_tecnico):
    from concurrent.futures import ThreadPoolExecutor
    from threading import Barrier

    from fastapi.testclient import TestClient

    _montar_vendedor()
    c = criar_colaborador()
    cabecalho = {"Authorization": client_tecnico.headers["Authorization"]}
    largada = Barrier(2, timeout=10)

    def chamar(_):
        outro = TestClient(_fastapi_app)
        largada.wait()
        return outro.post(
            "/api/acessos/eventos", json={"colaborador_id": c, "tipo": "admissao"}, headers=cabecalho
        ).status_code

    with ThreadPoolExecutor(2) as pool:
        codigos = sorted(pool.map(chamar, range(2)))
    assert codigos == [201, 409]
    assert contar("acesso_eventos") == 1


def test_busca_so_casa_cpf_quando_o_termo_e_numerico(client_tecnico):
    _montar_vendedor()
    c = criar_colaborador()  # CPF 11144477735, nome "Maria Silva"
    assert client_tecnico.post(
        "/api/acessos/eventos", json={"colaborador_id": c, "tipo": "admissao"}
    ).status_code == 201
    total = lambda busca: client_tecnico.get(  # noqa: E731
        "/api/acessos/eventos", params={"busca": busca}
    ).json()["total"]
    assert total("444") == 1  # CPF parcial com 3 dígitos
    assert total("111.444.777-35") == 1  # CPF com máscara
    assert total("7") == 0  # um dígito só não vira busca por CPF
    assert total("44") == 0
    assert total("silva 7") == 0  # texto misturado com dígito: só nome
    assert total("silva") == 1
