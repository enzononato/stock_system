"""Alterações de linhas com eventos automáticos: PUT, mudança de status e histórico (F5)."""
import pytest

from tests.fixtures_linhas import *  # noqa: F401,F403

pytestmark = pytest.mark.integration


@pytest.fixture
def aparelhos(criar_aparelho):
    a = criar_aparelho(brand="Samsung", model="A15", identificador="CEL-A")
    b = criar_aparelho(brand="Motorola", model="G54", identificador="CEL-B")
    return a, b


def _eventos(client, linha_id):
    resp = client.get(f"/api/linhas/{linha_id}/eventos")
    assert resp.status_code == 200, resp.text
    return resp.json()


def test_trocar_aparelho_gera_evento(client_gestor, usuario_gestor, nova_linha, aparelhos):
    a, b = aparelhos
    linha = nova_linha(item_id=a["id"])

    resp = client_gestor.put(f"/api/linhas/{linha['id']}", json={"item_id": b["id"], "motivo": "Tela quebrada"})

    assert resp.status_code == 200, resp.text
    assert resp.json()["item_id"] == b["id"]
    evento = _eventos(client_gestor, linha["id"])[0]
    assert evento["tipo"] == "troca_aparelho"
    assert evento["de_valor"] == "Celular Samsung A15 · CEL-A"
    assert evento["para_valor"] == "Celular Motorola G54 · CEL-B"
    assert evento["item_id"] == b["id"]
    assert evento["item_id_anterior"] == a["id"]
    assert evento["observacao"] == "Tela quebrada"
    assert evento["por_usuario"] == usuario_gestor["username"]
    assert evento["linha_numero"] == linha["numero"]


def test_trocar_usuario_gera_evento(client_gestor, nova_linha, criar_colaborador):
    maria = criar_colaborador("Maria Souza")
    joao = criar_colaborador("João Lima")
    linha = nova_linha(colaborador_id=maria)

    assert client_gestor.put(f"/api/linhas/{linha['id']}", json={"colaborador_id": joao}).status_code == 200
    evento = _eventos(client_gestor, linha["id"])[0]
    assert evento["tipo"] == "troca_usuario"
    assert (evento["de_valor"], evento["para_valor"]) == ("Maria Souza", "João Lima")

    resp = client_gestor.put(f"/api/linhas/{linha['id']}", json={"colaborador_id": None})
    assert resp.status_code == 200
    assert resp.json()["colaborador_id"] is None
    evento = _eventos(client_gestor, linha["id"])[0]
    assert evento["tipo"] == "troca_usuario"
    assert evento["de_valor"] == "João Lima"
    assert evento["para_valor"] is None


def test_mudar_operadora_gera_portabilidade(client_gestor, nova_linha):
    linha = nova_linha(operadora="Vivo")

    resp = client_gestor.put(f"/api/linhas/{linha['id']}", json={"operadora": "claro"})

    assert resp.status_code == 200
    assert resp.json()["operadora"] == "Claro"
    evento = _eventos(client_gestor, linha["id"])[0]
    assert evento["tipo"] == "portabilidade"
    assert (evento["de_valor"], evento["para_valor"]) == ("Vivo", "Claro")


def test_edicao_gera_um_evento_por_campo(client_gestor, nova_linha, criar_unidade, unidade):
    unidade_b = criar_unidade(nome="Unidade B")
    linha = nova_linha()

    resp = client_gestor.put(
        f"/api/linhas/{linha['id']}",
        json={"plano": "Controle 30GB", "custo_mensal": "59.90", "unidade_id": unidade_b["id"]},
    )

    assert resp.status_code == 200, resp.text
    edicoes = {e["campo"]: e for e in _eventos(client_gestor, linha["id"]) if e["tipo"] == "edicao"}
    assert set(edicoes) == {"plano", "custo_mensal", "unidade"}
    assert (edicoes["custo_mensal"]["de_valor"], edicoes["custo_mensal"]["para_valor"]) == ("49.90", "59.90")
    assert (edicoes["unidade"]["de_valor"], edicoes["unidade"]["para_valor"]) == (unidade["nome"], "Unidade B")
    assert (edicoes["plano"]["de_valor"], edicoes["plano"]["para_valor"]) == ("Controle 20GB", "Controle 30GB")


def test_edicao_do_numero_e_limpeza_de_campos(client_gestor, nova_linha, eventos_no_banco):
    linha = nova_linha(numero="74999990001")

    resp = client_gestor.put(
        f"/api/linhas/{linha['id']}", json={"numero": "(74) 99999-0002", "plano": None}
    )

    assert resp.status_code == 200, resp.text
    assert resp.json()["numero"] == "74999990002"
    assert resp.json()["plano"] is None
    edicoes = {e["campo"]: e for e in eventos_no_banco(linha["id"]) if e["tipo"] == "edicao"}
    assert (edicoes["numero"]["de_valor"], edicoes["numero"]["para_valor"]) == ("74999990001", "74999990002")
    assert (edicoes["plano"]["de_valor"], edicoes["plano"]["para_valor"]) == ("Controle 20GB", None)


def test_put_para_numero_de_outra_linha_409(client_gestor, nova_linha):
    nova_linha(numero="74999990001")
    outra = nova_linha(numero="74999990002")

    resp = client_gestor.put(f"/api/linhas/{outra['id']}", json={"numero": "(74) 99999-0001"})

    assert resp.status_code == 409
    assert resp.json()["detail"] == "Já existe uma linha com este número."


def test_put_sem_mudanca_nao_gera_evento_nem_auditoria(client_gestor, nova_linha, eventos_no_banco, auditoria):
    linha = nova_linha()
    antes = len(eventos_no_banco(linha["id"]))

    resp = client_gestor.put(
        f"/api/linhas/{linha['id']}",
        json={"numero": linha["numero"], "operadora": "Vivo", "custo_mensal": "49.9", "plano": "Controle 20GB"},
    )

    assert resp.status_code == 200
    assert len(eventos_no_banco(linha["id"])) == antes
    assert auditoria("linha", "editar") == []
    resp = client_gestor.put(f"/api/linhas/{linha['id']}", json={})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Nenhum campo para atualizar."


def test_auditoria_editar_so_campos_alterados(client_gestor, nova_linha, aparelhos, auditoria):
    a, b = aparelhos
    linha = nova_linha(item_id=a["id"])

    client_gestor.put(f"/api/linhas/{linha['id']}", json={"item_id": b["id"]})

    registro = auditoria("linha", "editar")[0]
    assert registro["entidade_id"] == str(linha["id"])
    assert registro["antes"] == {"item_id": a["id"]}
    assert registro["depois"] == {"item_id": b["id"]}


def test_falha_ao_gravar_evento_desfaz_a_alteracao(
    nova_linha, linha_db, ctx_teste, auditoria, monkeypatch
):
    linha = nova_linha(operadora="Vivo")

    def quebrar(*args, **kwargs):
        raise RuntimeError("falha ao gravar o evento")

    monkeypatch.setattr("app.db.linha_db._registrar_evento", quebrar)
    with pytest.raises(RuntimeError):
        linha_db.atualizar(linha["id"], {"operadora": "Claro"}, ctx_teste)

    assert linha_db.obter(linha["id"])["operadora"] == "Vivo"
    assert auditoria("linha", "editar") == []


def test_trocar_para_aparelho_ocupado_409(client_gestor, nova_linha, aparelhos, eventos_no_banco):
    a, b = aparelhos
    primeira = nova_linha(item_id=a["id"])
    nova_linha(numero="74999990009", item_id=b["id"])

    resp = client_gestor.put(f"/api/linhas/{primeira['id']}", json={"item_id": b["id"]})

    assert resp.status_code == 409
    assert resp.json()["detail"] == "O aparelho já está vinculado à linha (74) 99999-0009."
    assert client_gestor.get(f"/api/linhas/{primeira['id']}").json()["item_id"] == a["id"]
    assert [e["tipo"] for e in eventos_no_banco(primeira["id"])] == ["ativacao"]


def test_suspender_reativar_cancelar(client_gestor, nova_linha, auditoria):
    linha = nova_linha()
    url = f"/api/linhas/{linha['id']}/status"

    resp = client_gestor.post(url, json={"status": "suspensa", "observacao": "Férias"})
    assert resp.status_code == 200, resp.text
    assert resp.json()["status"] == "suspensa"
    evento = _eventos(client_gestor, linha["id"])[0]
    assert evento["tipo"] == "suspensao"
    assert (evento["de_valor"], evento["para_valor"]) == ("ativa", "suspensa")
    assert evento["observacao"] == "Férias"

    assert client_gestor.post(url, json={"status": "ativa"}).json()["status"] == "ativa"
    assert _eventos(client_gestor, linha["id"])[0]["tipo"] == "reativacao"

    assert client_gestor.post(url, json={"status": "cancelada"}).json()["status"] == "cancelada"
    assert _eventos(client_gestor, linha["id"])[0]["tipo"] == "cancelamento"

    assert len(auditoria("linha", "suspender")) == 1
    assert len(auditoria("linha", "reativar")) == 1
    cancelar = auditoria("linha", "cancelar")
    assert len(cancelar) == 1
    assert cancelar[0]["antes"]["status"] == "ativa"
    assert cancelar[0]["depois"]["status"] == "cancelada"


def test_transicoes_invalidas_409(client_gestor, nova_linha):
    linha = nova_linha()
    url = f"/api/linhas/{linha['id']}/status"

    client_gestor.post(url, json={"status": "suspensa"})
    resp = client_gestor.post(url, json={"status": "suspensa"})
    assert resp.status_code == 409
    assert resp.json()["detail"] == "Só linhas ativas podem ser suspensas."

    client_gestor.post(url, json={"status": "ativa"})
    resp = client_gestor.post(url, json={"status": "ativa"})
    assert resp.status_code == 409
    assert resp.json()["detail"] == "A linha já está ativa."

    client_gestor.post(url, json={"status": "cancelada"})
    resp = client_gestor.post(url, json={"status": "cancelada"})
    assert resp.status_code == 409
    assert resp.json()["detail"] == "A linha já está cancelada."
    resp = client_gestor.post(url, json={"status": "suspensa"})
    assert resp.status_code == 409

    resp = client_gestor.post(url, json={"status": "bloqueada"})
    assert resp.status_code == 422
    assert resp.json()["detail"].endswith("escolha ativa, suspensa ou cancelada.")


def test_cancelar_libera_o_aparelho_e_reativar_funciona(client_gestor, nova_linha, aparelhos, eventos_no_banco):
    a, _ = aparelhos
    linha = nova_linha(item_id=a["id"])
    url = f"/api/linhas/{linha['id']}/status"

    resp = client_gestor.post(url, json={"status": "cancelada"})
    assert resp.json()["item_id"] is None
    cancelamento = _eventos(client_gestor, linha["id"])[0]
    assert cancelamento["tipo"] == "cancelamento"
    assert cancelamento["item_id_anterior"] == a["id"]

    assert client_gestor.post("/api/linhas", json={
        "numero": "74999990077", "operadora": "Vivo", "custo_mensal": "10.00",
        "unidade_id": linha["unidade_id"], "item_id": a["id"],
    }).status_code == 201

    resp = client_gestor.put(f"/api/linhas/{linha['id']}", json={"plano": "Novo"})
    assert resp.status_code == 409
    assert resp.json()["detail"] == "Linha cancelada: reative-a antes de alterar."

    resp = client_gestor.post(url, json={"status": "ativa"})
    assert resp.status_code == 200
    assert resp.json()["status"] == "ativa"
    assert resp.json()["item_id"] is None


def test_eventos_do_aparelho_inclui_linha_que_saiu(client_gestor, nova_linha, aparelhos):
    a, b = aparelhos
    l1 = nova_linha(item_id=a["id"])
    client_gestor.put(f"/api/linhas/{l1['id']}", json={"item_id": b["id"]})

    de_a = client_gestor.get("/api/linhas/eventos", params={"item_id": a["id"]})
    assert de_a.status_code == 200
    assert [e["tipo"] for e in de_a.json()] == ["troca_aparelho", "ativacao"]
    assert all(e["linha_numero"] == l1["numero"] for e in de_a.json())
    de_b = client_gestor.get("/api/linhas/eventos", params={"item_id": b["id"]}).json()
    assert [e["tipo"] for e in de_b] == ["troca_aparelho"]


def test_permissoes_de_alteracao(client_aprendiz, nova_linha):
    linha = nova_linha()

    resp = client_aprendiz.put(f"/api/linhas/{linha['id']}", json={"plano": "X"})
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
    resp = client_aprendiz.post(f"/api/linhas/{linha['id']}/status", json={"status": "suspensa"})
    assert resp.status_code == 403
    assert client_aprendiz.get(f"/api/linhas/{linha['id']}/eventos").status_code == 200
    resp = client_aprendiz.get("/api/linhas/999999/eventos")
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Linha não encontrada."
