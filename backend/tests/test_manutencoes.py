"""F4 / Task 5: ordens de manutenção (custo, anexo e efeito no status do equipamento)."""
import json
from datetime import date, datetime, timedelta

import pytest

pytestmark = pytest.mark.integration

PDF = {"anexo": ("os.pdf", b"%PDF-1.4 ordem", "application/pdf")}


def _abrir(client, item_id, defeito="Tela quebrada", **extra):
    resp = client.post("/api/manutencoes", json={"item_id": item_id, "defeito": defeito, **extra})
    assert resp.status_code == 200, resp.text
    return resp.json()["id"]


def _retorno(client, mid, resultado="retornada", custo=None, **kw):
    data = {"resultado": resultado}
    if custo is not None:
        data["custo"] = custo
    return client.post(f"/api/manutencoes/{mid}/retorno", data=data, **kw)


def _item(client, item_id):
    return client.get(f"/api/items/{item_id}").json()


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


def test_abrir_coloca_item_em_manutencao(client_gestor, item_disponivel):
    resp = client_gestor.post(
        "/api/manutencoes", json={"item_id": item_disponivel["id"], "defeito": "Tela quebrada"}
    )
    assert resp.status_code == 200
    assert resp.json()["detail"] == "Manutenção aberta."
    assert "id" in resp.json()
    assert _item(client_gestor, item_disponivel["id"])["status"] == "Em Manutenção"


def test_item_em_manutencao_nao_fica_disponivel_nem_emprestavel(client_gestor, inv_manager, item_disponivel):
    _abrir(client_gestor, item_disponivel["id"])
    resp = client_gestor.get("/api/items", params={"status": "Disponível"})
    assert item_disponivel["id"] not in [i["id"] for i in resp.json()["items"]]
    ok, msg = inv_manager.issue(
        item_disponivel["id"], "Fulano de Tal", "11122233344", "101 - Puxada", "Analista", "TI",
        "Revalle Juazeiro", datetime.now().strftime("%d/%m/%Y"), "fixture_teste",
    )
    assert (ok, msg) == (False, "Este item não está disponível para empréstimo.")


def test_abrir_para_item_com_colaborador_recusa(client_gestor, item_indisponivel):
    resp = client_gestor.post(
        "/api/manutencoes", json={"item_id": item_indisponivel["id"], "defeito": "Não liga"}
    )
    assert resp.status_code == 400
    assert resp.json()["detail"] == (
        "Só é possível abrir manutenção para equipamento sem colaborador. Faça a devolução antes."
    )


def test_abrir_item_inexistente_404(client_gestor):
    resp = client_gestor.post("/api/manutencoes", json={"item_id": 999999, "defeito": "x"})
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Item não encontrado."


def test_abrir_sem_defeito_422(client_gestor, item_disponivel):
    resp = client_gestor.post("/api/manutencoes", json={"item_id": item_disponivel["id"], "defeito": "  "})
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Defeito: Descreva o defeito."


def test_abrir_segunda_manutencao_recusa(client_gestor, item_disponivel):
    _abrir(client_gestor, item_disponivel["id"])
    resp = client_gestor.post(
        "/api/manutencoes", json={"item_id": item_disponivel["id"], "defeito": "Outro"}
    )
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Este equipamento já tem uma manutenção em andamento."


def test_fluxo_envio_reparo_retorno_com_custo_e_anexo(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    previsao = (date.today() + timedelta(days=7)).strftime("%d/%m/%Y")
    resp = client_gestor.post(
        f"/api/manutencoes/{mid}/enviar", json={"assistencia": "TecInfo", "previsao_retorno": previsao}
    )
    assert resp.status_code == 200
    assert resp.json()["detail"] == "Envio à assistência registrado."
    resp = client_gestor.post(f"/api/manutencoes/{mid}/em-reparo", json={})
    assert resp.status_code == 200
    assert resp.json()["detail"] == "Manutenção marcada como em reparo."
    resp = _retorno(client_gestor, mid, custo="150.00", files=PDF)
    assert resp.status_code == 200, resp.text
    assert resp.json() == {
        "detail": "Retorno registrado. O equipamento voltou a Disponível.", "sugerir_baixa": False,
    }
    assert _item(client_gestor, item_disponivel["id"])["status"] == "Disponível"

    m = client_gestor.get(f"/api/manutencoes/{mid}").json()
    assert m["status"] == "retornada"
    assert m["custo"] == "150.00"
    assert m["assistencia"] == "TecInfo"
    assert m["anexo_key"].startswith("manutencoes/")
    assert m["enviada_em"] and m["em_reparo_em"] and m["retornada_em"]
    assert m["encerrada_por"] == "gestor.teste"
    assert m["previsao_retorno"] == (date.today() + timedelta(days=7)).isoformat()
    assert client_gestor.get(f"/api/manutencoes/{mid}/anexo").status_code == 200


def test_anexo_inexistente_404(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    resp = client_gestor.get(f"/api/manutencoes/{mid}/anexo")
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Nenhum anexo nesta manutenção."


def test_manutencao_inexistente_404(client_gestor):
    resp = client_gestor.get("/api/manutencoes/999999")
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Manutenção não encontrada."
    assert client_gestor.post("/api/manutencoes/999999/em-reparo", json={}).status_code == 404


def test_previsao_antes_da_abertura_400(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    ontem = (date.today() - timedelta(days=1)).strftime("%d/%m/%Y")
    resp = client_gestor.post(
        f"/api/manutencoes/{mid}/enviar", json={"assistencia": "TecInfo", "previsao_retorno": ontem}
    )
    assert resp.status_code == 400
    assert resp.json()["detail"] == "A previsão de retorno não pode ser anterior à abertura."
    assert client_gestor.get(f"/api/manutencoes/{mid}").json()["status"] == "aberta"


def test_sem_conserto_mantem_item_e_permite_nova_manutencao(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    resp = _retorno(client_gestor, mid, resultado="sem_conserto")
    assert resp.status_code == 200
    assert resp.json() == {
        "detail": "Manutenção encerrada sem conserto. Avalie a baixa do equipamento.",
        "sugerir_baixa": True,
    }
    assert _item(client_gestor, item_disponivel["id"])["status"] == "Em Manutenção"
    assert client_gestor.get(f"/api/manutencoes/{mid}").json()["status"] == "sem_conserto"
    resp = client_gestor.post(
        "/api/manutencoes", json={"item_id": item_disponivel["id"], "defeito": "Segunda assistência"}
    )
    assert resp.status_code == 200


def test_cancelar_volta_disponivel(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    resp = client_gestor.post(f"/api/manutencoes/{mid}/cancelar", json={"motivo": "Aberta por engano"})
    assert resp.status_code == 200
    assert resp.json()["detail"] == "Manutenção cancelada. O equipamento voltou a Disponível."
    assert _item(client_gestor, item_disponivel["id"])["status"] == "Disponível"
    m = client_gestor.get(f"/api/manutencoes/{mid}").json()
    assert m["status"] == "cancelada"
    assert m["cancelada_em"]
    assert "Cancelada: Aberta por engano" in m["observacoes"]


def test_cancelar_sem_motivo_422(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    resp = client_gestor.post(f"/api/manutencoes/{mid}/cancelar", json={"motivo": " "})
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Motivo: Informe o motivo do cancelamento."


def test_transicao_invalida(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    assert _retorno(client_gestor, mid).status_code == 200
    resp = _retorno(client_gestor, mid)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Ação não permitida para manutenção com status Retornada."
    resp = client_gestor.post(f"/api/manutencoes/{mid}/cancelar", json={"motivo": "x"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Ação não permitida para manutenção com status Retornada."


def test_retorno_com_anexo_invalido_nao_altera(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    resp = _retorno(client_gestor, mid, files={"anexo": ("x.pdf", b"MZ", "application/pdf")})
    assert resp.status_code == 400
    assert client_gestor.get(f"/api/manutencoes/{mid}").json()["status"] == "aberta"
    assert _item(client_gestor, item_disponivel["id"])["status"] == "Em Manutenção"


def test_custo_negativo_400(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    resp = _retorno(client_gestor, mid, custo="-1")
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Informe um custo maior ou igual a zero."


def test_custo_invalido_400(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    resp = _retorno(client_gestor, mid, custo="1,50")
    assert resp.status_code == 400
    assert resp.json()["detail"] == (
        "Custo inválido: use números com ponto para os centavos (ex.: 150.00)."
    )
    assert client_gestor.get(f"/api/manutencoes/{mid}").json()["status"] == "aberta"


def _ciclo(client, item_id, custo):
    mid = _abrir(client, item_id)
    assert _retorno(client, mid, custo=custo).status_code == 200
    return mid


def test_custo_acumulado_por_equipamento(client_gestor, criar_item):
    a, b = criar_item(), criar_item()
    _ciclo(client_gestor, a["id"], "150.00")
    _ciclo(client_gestor, a["id"], "89.90")
    cancelada = _abrir(client_gestor, a["id"])
    client_gestor.post(f"/api/manutencoes/{cancelada}/cancelar", json={"motivo": "Desistiu"})
    _ciclo(client_gestor, b["id"], "1000.00")

    corpo = client_gestor.get("/api/manutencoes", params={"item_id": a["id"]}).json()
    assert corpo["total"] == 3
    assert corpo["custo_total"] == "239.90"
    assert {m["custo_acumulado_item"] for m in corpo["items"]} == {"239.90"}

    assert client_gestor.get("/api/manutencoes").json()["custo_total"] == "1239.90"
    parcial = client_gestor.get("/api/manutencoes", params={"item_id": a["id"], "limit": 1}).json()
    assert len(parcial["items"]) == 1
    assert parcial["total"] == 3
    assert parcial["custo_total"] == "239.90"


def test_lista_vazia_tem_custo_zero(client_gestor):
    corpo = client_gestor.get("/api/manutencoes").json()
    assert corpo == {"items": [], "total": 0, "custo_total": "0.00"}


def test_lista_traz_dados_do_equipamento(client_gestor, item_disponivel):
    _abrir(client_gestor, item_disponivel["id"])
    m = client_gestor.get("/api/manutencoes").json()["items"][0]
    assert m["item_tipo"] == "Notebook"
    assert m["item_brand"] == "Dell"
    assert m["item_identificador"] == item_disponivel["identificador"]
    assert m["item_revenda"] == "Revalle Juazeiro"
    assert m["item_status"] == "Em Manutenção"


def test_item_baixado_mantem_historico_de_custo(client_gestor, inv_manager, criar_item):
    item = criar_item(tipo="Impressora")
    _ciclo(client_gestor, item["id"], "50.00")
    ok, msg = inv_manager.remove(item["id"], "teste", "Obsolescência")
    assert ok, msg
    corpo = client_gestor.get("/api/manutencoes").json()
    assert corpo["total"] == 1
    assert corpo["custo_total"] == "50.00"


def test_filtros(client_gestor, criar_item, executar_sql):
    em_andamento = criar_item(tipo="Notebook", revenda="Revalle Juazeiro")
    encerrada = criar_item(tipo="Impressora", revenda="Revalle Petrolina", brand="HP")
    m1 = _abrir(client_gestor, em_andamento["id"], defeito="Bateria")
    client_gestor.post(f"/api/manutencoes/{m1}/enviar", json={"assistencia": "TecInfo"})
    m2 = _ciclo(client_gestor, encerrada["id"], "10.00")
    executar_sql("UPDATE manutencoes SET aberta_em='2026-01-10 09:00:00' WHERE id=%s", (m2,))

    def ids(**params):
        corpo = client_gestor.get("/api/manutencoes", params=params).json()
        return [m["id"] for m in corpo["items"]]

    assert ids(status="em_andamento") == [m1]
    assert ids(status="retornada") == [m2]
    assert ids(revenda="Revalle Petrolina") == [m2]
    assert ids(tipo="Notebook") == [m1]
    assert ids(de="2026-01-01", ate="2026-01-10") == [m2]
    assert ids(de="2026-01-11", ate="2026-01-31") == []
    assert ids(busca="TecInfo") == [m1]
    assert ids(busca="Bateria") == [m1]
    assert ids(busca="HP") == [m2]
    assert ids() == [m1, m2]  # aberta_em decrescente
    assert client_gestor.get("/api/manutencoes", params={"status": "invalido"}).status_code == 422


def test_exige_gestor_ou_tecnico(client_aprendiz, client_tecnico, item_disponivel):
    resp = client_aprendiz.get("/api/manutencoes")
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
    resp = client_aprendiz.post("/api/manutencoes", json={"item_id": item_disponivel["id"], "defeito": "x"})
    assert resp.status_code == 403
    assert client_tecnico.get("/api/manutencoes").status_code == 200
    assert client_tecnico.post(
        "/api/manutencoes", json={"item_id": item_disponivel["id"], "defeito": "x"}
    ).status_code == 200


def test_auditoria_de_manutencao(client_gestor, item_disponivel):
    mid = _abrir(client_gestor, item_disponivel["id"])
    assert _retorno(client_gestor, mid, custo="10.00").status_code == 200
    criar = _auditoria("manutencao", "criar")
    assert len(criar) == 1
    assert criar[0]["entidade_id"] == str(mid)
    assert criar[0]["antes"] is None
    assert criar[0]["depois"]["defeito"] == "Tela quebrada"
    assert criar[0]["depois"]["item_id"] == item_disponivel["id"]
    retornar = _auditoria("manutencao", "retornar")
    assert len(retornar) == 1
    assert retornar[0]["entidade_id"] == str(mid)
    assert retornar[0]["antes"]["status"] == "aberta"
    assert retornar[0]["depois"]["status"] == "retornada"
    assert retornar[0]["depois"]["item_id"] == item_disponivel["id"]
    assert retornar[0]["depois"]["item_status"] == "Disponível"


def test_auditoria_das_demais_acoes(client_gestor, criar_item):
    a, b, c = criar_item(), criar_item(), criar_item()
    m1 = _abrir(client_gestor, a["id"])
    client_gestor.post(f"/api/manutencoes/{m1}/enviar", json={"assistencia": "TecInfo"})
    client_gestor.post(f"/api/manutencoes/{m1}/em-reparo", json={})
    m2 = _abrir(client_gestor, b["id"])
    _retorno(client_gestor, m2, resultado="sem_conserto")
    m3 = _abrir(client_gestor, c["id"])
    client_gestor.post(f"/api/manutencoes/{m3}/cancelar", json={"motivo": "Engano"})
    for acao, mid in (("enviar", m1), ("em_reparo", m1), ("sem_conserto", m2), ("cancelar", m3)):
        linhas = _auditoria("manutencao", acao)
        assert [l["entidade_id"] for l in linhas] == [str(mid)], acao


def test_cancelar_segunda_ordem_apos_sem_conserto_mantem_em_manutencao(client_gestor, item_disponivel):
    id_item = item_disponivel["id"]
    primeira = _abrir(client_gestor, id_item)
    assert _retorno(client_gestor, primeira, resultado="sem_conserto").status_code == 200
    segunda = _abrir(client_gestor, id_item, defeito="Segunda assistência")
    resp = client_gestor.post(f"/api/manutencoes/{segunda}/cancelar", json={"motivo": "Desistiu"})
    assert resp.status_code == 200
    # O defeito conhecido continua: não vira "Disponível".
    assert _item(client_gestor, id_item)["status"] == "Em Manutenção"
    assert client_gestor.get(f"/api/manutencoes/{segunda}").json()["status"] == "cancelada"


def test_cancelar_depois_de_reparo_bem_sucedido_volta_disponivel(client_gestor, item_disponivel):
    id_item = item_disponivel["id"]
    primeira = _abrir(client_gestor, id_item)
    assert _retorno(client_gestor, primeira, resultado="sem_conserto").status_code == 200
    segunda = _abrir(client_gestor, id_item)
    assert _retorno(client_gestor, segunda, resultado="retornada").status_code == 200
    assert _item(client_gestor, id_item)["status"] == "Disponível"
    terceira = _abrir(client_gestor, id_item)
    client_gestor.post(f"/api/manutencoes/{terceira}/cancelar", json={"motivo": "Engano"})
    assert _item(client_gestor, id_item)["status"] == "Disponível"


def test_retorno_nao_sobrescreve_status_alterado_fora_da_manutencao(client_gestor, item_disponivel, executar_sql):
    id_item = item_disponivel["id"]
    mid = _abrir(client_gestor, id_item)
    # Simula outra operação (ex.: estorno) que mudou o status do item no meio do caminho.
    executar_sql("UPDATE items SET status='Indisponível' WHERE id=%s", (id_item,))
    resp = _retorno(client_gestor, mid, custo="10.00")
    assert resp.status_code == 409
    assert "status do equipamento mudou" in resp.json()["detail"]
    assert _item(client_gestor, id_item)["status"] == "Indisponível"
    m = client_gestor.get(f"/api/manutencoes/{mid}").json()
    assert m["status"] == "aberta"
    assert m["custo"] is None


def test_cancelar_nao_sobrescreve_status_alterado_fora_da_manutencao(client_gestor, item_disponivel, executar_sql):
    id_item = item_disponivel["id"]
    mid = _abrir(client_gestor, id_item)
    executar_sql("UPDATE items SET status='Indisponível' WHERE id=%s", (id_item,))
    resp = client_gestor.post(f"/api/manutencoes/{mid}/cancelar", json={"motivo": "Engano"})
    assert resp.status_code == 409
    assert _item(client_gestor, id_item)["status"] == "Indisponível"
    assert client_gestor.get(f"/api/manutencoes/{mid}").json()["status"] == "aberta"
