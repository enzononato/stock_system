"""F4 / Task 4: transferência de equipamento entre unidades."""
import json

import pytest

pytestmark = pytest.mark.integration

JUAZEIRO = "Revalle Juazeiro"
PETROLINA = "Revalle Petrolina"


@pytest.fixture
def unidades(criar_unidade):
    return criar_unidade(nome=JUAZEIRO), criar_unidade(nome=PETROLINA)


def _local(client, unidade, nome="Sala TI"):
    resp = client.post("/api/locais", json={"unidade_id": unidade["id"], "nome": nome})
    assert resp.status_code in (200, 201), resp.text
    return resp.json()["id"]


def _consultar(sql, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


def _transferir(client, item_id, unidade_destino, **extra):
    corpo = {"para_unidade_id": unidade_destino["id"], "motivo": "Mudança de setor"}
    corpo.update(extra)
    return client.post(f"/api/items/{item_id}/transferir", json=corpo)


def test_transferir_atualiza_item_e_registra(client_gestor, unidades, item_disponivel):
    juazeiro, petrolina = unidades
    local_id = _local(client_gestor, petrolina)
    resp = _transferir(client_gestor, item_disponivel["id"], petrolina, para_local_id=local_id)
    assert resp.status_code == 200, resp.text
    assert resp.json()["detail"] == "Equipamento transferido de Revalle Juazeiro para Revalle Petrolina."

    item = client_gestor.get(f"/api/items/{item_disponivel['id']}").json()
    assert item["revenda"] == PETROLINA
    assert item["local_id"] == local_id

    linhas = _consultar("SELECT * FROM transferencias WHERE item_id=%s", (item_disponivel["id"],))
    assert len(linhas) == 1
    assert linhas[0]["id"] == resp.json()["id"]
    assert linhas[0]["de_unidade"] == JUAZEIRO
    assert linhas[0]["para_unidade"] == PETROLINA
    assert linhas[0]["para_local_id"] == local_id
    assert linhas[0]["motivo"] == "Mudança de setor"
    assert linhas[0]["transferido_por"] == "gestor.teste"

    hist = _consultar(
        "SELECT * FROM history WHERE item_id=%s AND operation='Edição' ORDER BY id DESC",
        (item_disponivel["id"],),
    )
    assert hist[0]["details"] == "Transferência Revalle Juazeiro → Revalle Petrolina"
    assert hist[0]["revenda"] == PETROLINA

    auditoria = _consultar("SELECT * FROM audit_log WHERE entidade='item' AND acao='transferir'")
    assert len(auditoria) == 1
    assert json.loads(auditoria[0]["antes"]) == {"revenda": JUAZEIRO, "local_id": None}
    assert json.loads(auditoria[0]["depois"]) == {
        "revenda": PETROLINA, "local_id": local_id, "motivo": "Mudança de setor",
    }


def test_transferir_sem_local_limpa_o_local_antigo(client_gestor, unidades, criar_item):
    juazeiro, petrolina = unidades
    item = criar_item(local_id=_local(client_gestor, juazeiro))
    assert _transferir(client_gestor, item["id"], petrolina).status_code == 200
    assert client_gestor.get(f"/api/items/{item['id']}").json()["local_id"] is None


def test_transferir_para_a_mesma_unidade_400(client_gestor, unidades, item_disponivel):
    juazeiro, _ = unidades
    resp = _transferir(client_gestor, item_disponivel["id"], juazeiro)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "O equipamento já está nesta unidade."


def test_transferir_local_de_outra_unidade_400(client_gestor, unidades, item_disponivel):
    juazeiro, petrolina = unidades
    local_juazeiro = _local(client_gestor, juazeiro)
    resp = _transferir(client_gestor, item_disponivel["id"], petrolina, para_local_id=local_juazeiro)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "O local informado não pertence à unidade de destino."
    assert client_gestor.get(f"/api/items/{item_disponivel['id']}").json()["revenda"] == JUAZEIRO


def test_transferir_item_pendente_400(client_gestor, unidades, item_pendente):
    _, petrolina = unidades
    resp = _transferir(client_gestor, item_pendente["id"], petrolina)
    assert resp.status_code == 400
    assert resp.json()["detail"] == (
        "Conclua ou estorne o empréstimo ou a devolução pendente antes de transferir."
    )


def test_transferir_unidade_inativa_400(client_gestor, unidade_db, unidades, item_disponivel):
    _, petrolina = unidades
    ok, msg = unidade_db.deactivate_unidade(petrolina["id"])
    assert ok, msg
    resp = _transferir(client_gestor, item_disponivel["id"], petrolina)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Unidade de destino não encontrada ou inativa."


def test_transferir_item_inexistente_404(client_gestor, unidades):
    _, petrolina = unidades
    resp = _transferir(client_gestor, 999999, petrolina)
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Item não encontrado."


def test_transferir_sem_motivo_422(client_gestor, unidades, item_disponivel):
    _, petrolina = unidades
    resp = _transferir(client_gestor, item_disponivel["id"], petrolina, motivo="   ")
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Motivo: Informe o motivo da transferência."


def test_transferir_motivo_longo_422(client_gestor, unidades, item_disponivel):
    _, petrolina = unidades
    resp = _transferir(client_gestor, item_disponivel["id"], petrolina, motivo="x" * 256)
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Motivo: O motivo deve ter no máximo 255 caracteres."


def test_transferir_exige_gestor_ou_tecnico(client_aprendiz, client_tecnico, unidades, item_disponivel):
    _, petrolina = unidades
    resp = _transferir(client_aprendiz, item_disponivel["id"], petrolina)
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
    assert _transferir(client_tecnico, item_disponivel["id"], petrolina).status_code == 200
