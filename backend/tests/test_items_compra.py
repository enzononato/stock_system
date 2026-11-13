"""F4 / Task 3: dados de compra, local e anexo da NF no equipamento."""
import json
from datetime import date, timedelta

import pytest

from app.services.anexos import MAX_ANEXO_BYTES

pytestmark = pytest.mark.integration

PDF = {"arquivo": ("nf.pdf", b"%PDF-1.4 teste", "application/pdf")}
_contador = {"n": 0}


def _corpo(revenda="Revalle Juazeiro", **extra):
    _contador["n"] += 1
    corpo = {
        "tipo": "Notebook", "brand": "Dell", "model": "Latitude", "revenda": revenda,
        "date_registered": "01/02/2026", "identificador": f"SN-COMPRA-{_contador['n']}",
    }
    corpo.update(extra)
    return corpo


def _criar(client, **extra):
    resp = client.post("/api/items", json=_corpo(**extra))
    assert resp.status_code == 200, resp.text
    return resp.json()["id"]


def _local(client, unidade, nome="Sala TI"):
    resp = client.post("/api/locais", json={"unidade_id": unidade["id"], "nome": nome})
    assert resp.status_code in (200, 201), resp.text
    return resp.json()["id"]


def _auditoria(entidade, acao):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT * FROM audit_log WHERE entidade=%s AND acao=%s ORDER BY id", (entidade, acao)
        )
        linhas = list(cur.fetchall())
    for linha in linhas:
        linha["antes"] = json.loads(linha["antes"]) if linha["antes"] else None
        linha["depois"] = json.loads(linha["depois"]) if linha["depois"] else None
    return linhas


def test_criar_item_com_dados_de_compra(client_gestor):
    item_id = _criar(
        client_gestor, valor_compra=4599.9, data_compra="10/01/2026", garantia_ate="10/01/2027"
    )
    item = client_gestor.get(f"/api/items/{item_id}").json()
    assert item["valor_compra"] == "4599.90"
    assert item["data_compra"] == "2026-01-10"
    assert item["garantia_ate"] == "2027-01-10"


def test_valor_negativo_422(client_gestor):
    resp = client_gestor.post("/api/items", json=_corpo(valor_compra=-1))
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Valor de compra: Informe um valor de compra maior ou igual a zero."


def test_valor_com_tres_casas_422(client_gestor):
    resp = client_gestor.post("/api/items", json=_corpo(valor_compra=4599.999))
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Valor de compra: Use no máximo duas casas decimais no valor de compra."


def test_data_compra_futura_422(client_gestor):
    amanha = (date.today() + timedelta(days=1)).strftime("%d/%m/%Y")
    resp = client_gestor.post("/api/items", json=_corpo(data_compra=amanha))
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Data da compra: A data da compra não pode ser no futuro."


def test_garantia_antes_da_compra_400(client_gestor):
    resp = client_gestor.post(
        "/api/items", json=_corpo(data_compra="10/01/2026", garantia_ate="09/01/2026")
    )
    assert resp.status_code == 400
    assert resp.json()["detail"] == "A garantia não pode terminar antes da data da compra."


def test_criar_com_local_da_unidade(client_gestor, criar_unidade):
    u = criar_unidade(nome="Revalle Juazeiro")
    local_id = _local(client_gestor, u)
    item_id = _criar(client_gestor, local_id=local_id)
    assert client_gestor.get(f"/api/items/{item_id}").json()["local_id"] == local_id


def test_local_de_outra_unidade_400(client_gestor, criar_unidade):
    criar_unidade(nome="Revalle Juazeiro")
    outra = criar_unidade(nome="Revalle Petrolina")
    local_id = _local(client_gestor, outra)
    resp = client_gestor.post("/api/items", json=_corpo(local_id=local_id))
    assert resp.status_code == 400
    assert resp.json()["detail"] == "O local informado não pertence à unidade do equipamento."


def test_editar_limpa_valor_com_null(client_gestor):
    item_id = _criar(client_gestor, valor_compra=100, data_compra="10/01/2026")
    assert client_gestor.put(f"/api/items/{item_id}", json={"valor_compra": None}).status_code == 200
    item = client_gestor.get(f"/api/items/{item_id}").json()
    assert item["valor_compra"] is None
    assert item["data_compra"] == "2026-01-10"
    assert client_gestor.put(f"/api/items/{item_id}", json={"brand": "X"}).status_code == 200
    assert client_gestor.get(f"/api/items/{item_id}").json()["data_compra"] == "2026-01-10"


def test_editar_mudando_revenda_recusado(client_gestor):
    item_id = _criar(client_gestor)
    resp = client_gestor.put(f"/api/items/{item_id}", json={"revenda": "Outra Unidade"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Para mudar a unidade do equipamento, use a transferência."
    assert client_gestor.put(
        f"/api/items/{item_id}", json={"revenda": "Revalle Juazeiro", "brand": "Y"}
    ).status_code == 200


def test_editar_mantendo_local_inativo_aceito(client_gestor, criar_unidade):
    u = criar_unidade(nome="Revalle Juazeiro")
    local_id = _local(client_gestor, u)
    item_id = _criar(client_gestor, local_id=local_id)
    assert client_gestor.put(f"/api/locais/{local_id}", json={"ativo": False}).status_code == 200
    resp = client_gestor.put(f"/api/items/{item_id}", json={"local_id": local_id, "brand": "Z"})
    assert resp.status_code == 200, resp.text


def test_editar_garantia_confere_com_data_gravada(client_gestor):
    item_id = _criar(client_gestor, data_compra="10/01/2026")
    resp = client_gestor.put(f"/api/items/{item_id}", json={"garantia_ate": "01/01/2026"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "A garantia não pode terminar antes da data da compra."


def test_anexar_e_baixar_nf_pdf(client_gestor):
    item_id = _criar(client_gestor)
    resp = client_gestor.post(f"/api/items/{item_id}/nota-fiscal", files=PDF)
    assert resp.status_code == 200, resp.text
    assert resp.json()["detail"] == "Nota fiscal anexada."
    item = client_gestor.get(f"/api/items/{item_id}").json()
    assert item["nf_anexo_key"].startswith(f"notas_fiscais/nf_{item_id}_")

    baixado = client_gestor.get(f"/api/items/{item_id}/nota-fiscal")
    assert baixado.status_code == 200
    assert baixado.content == b"%PDF-1.4 teste"
    assert baixado.headers["content-type"] == "application/pdf"

    linhas = [x for x in _auditoria("item", "editar") if x["depois"] and "nf_anexo_key" in x["depois"]]
    assert len(linhas) == 1
    assert linhas[0]["depois"]["nf_anexo_key"] == item["nf_anexo_key"]


def test_nf_acima_de_10mb_413(client_gestor):
    item_id = _criar(client_gestor)
    grande = {"arquivo": ("nf.pdf", b"%PDF" + b"0" * MAX_ANEXO_BYTES, "application/pdf")}
    resp = client_gestor.post(f"/api/items/{item_id}/nota-fiscal", files=grande)
    assert resp.status_code == 413
    assert resp.json()["detail"] == "Arquivo excede o tamanho máximo de 10MB."


def test_nf_tipo_invalido_400(client_gestor):
    item_id = _criar(client_gestor)
    resp = client_gestor.post(
        f"/api/items/{item_id}/nota-fiscal", files={"arquivo": ("nf.pdf", b"MZ\x90", "application/pdf")}
    )
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Envie um PDF ou uma imagem PNG/JPG."


def test_nf_sem_anexo_404(client_gestor):
    item_id = _criar(client_gestor)
    resp = client_gestor.get(f"/api/items/{item_id}/nota-fiscal")
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Nenhuma nota fiscal anexada a este equipamento."


def test_nf_item_inexistente_404(client_gestor):
    resp = client_gestor.post("/api/items/999999/nota-fiscal", files=PDF)
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Item não encontrado."


def test_nf_exige_gestor_ou_tecnico(client_aprendiz, client_tecnico, item_disponivel):
    rota = f"/api/items/{item_disponivel['id']}/nota-fiscal"
    resp = client_aprendiz.post(rota, files=PDF)
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
    assert client_aprendiz.get(rota).status_code == 403
    assert client_tecnico.post(rota, files=PDF).status_code == 200
