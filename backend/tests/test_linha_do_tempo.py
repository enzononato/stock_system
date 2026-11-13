"""F4 / Task 7: linha do tempo do equipamento (registro de fontes + endpoint)."""
from datetime import datetime
from decimal import Decimal

import pytest

from app.services import linha_do_tempo
from app.services.linha_do_tempo import Evento
from tests.dados_baixa import APAGAMENTO_OK


# ── Unidade ─────────────────────────────────────────────────────────────────

@pytest.mark.unit
def test_registrar_fonte_substitui_pelo_nome(monkeypatch):
    monkeypatch.setattr(linha_do_tempo, "_FONTES", {})

    @linha_do_tempo.registrar_fonte
    def fonte_a(item_id):
        return []

    @linha_do_tempo.registrar_fonte
    def fonte_b(item_id):
        return []

    def fonte_a_nova(item_id):
        return []

    fonte_a_nova.__name__ = "fonte_a"
    linha_do_tempo.registrar_fonte(fonte_a_nova)
    assert linha_do_tempo.fontes_registradas() == ["fonte_a", "fonte_b"]
    assert linha_do_tempo._FONTES["fonte_a"] is fonte_a_nova


@pytest.mark.unit
def test_empate_ordena_por_fonte_e_referencia(monkeypatch):
    linha_do_tempo.carregar_fontes()
    monkeypatch.setattr(linha_do_tempo, "_FONTES", {})
    momento = datetime(2026, 1, 1, 10, 0)

    @linha_do_tempo.registrar_fonte
    def primeira(item_id):
        return [
            Evento(momento, "a", "x", "a2", referencia_id=2),
            Evento(momento, "a", "x", "a1", referencia_id=1),
        ]

    @linha_do_tempo.registrar_fonte
    def segunda(item_id):
        return [
            Evento(momento, "b", "x", "b1", referencia_id=1),
            Evento(datetime(2025, 12, 31), "b", "x", "antes"),
        ]

    resultado = linha_do_tempo.montar_linha_do_tempo(1)
    assert [e.titulo for e in resultado.eventos] == ["antes", "a1", "a2", "b1"]
    assert resultado.fontes_com_erro == []


@pytest.mark.unit
def test_moeda_br():
    from app.services.linha_do_tempo_fontes import moeda_br

    assert moeda_br(Decimal("1234.5")) == "R$ 1.234,50"
    assert moeda_br(Decimal("0")) == "R$ 0,00"
    assert moeda_br(Decimal("1000000.99")) == "R$ 1.000.000,99"


# ── Integração ──────────────────────────────────────────────────────────────

def _eventos(client, item_id):
    resp = client.get(f"/api/items/{item_id}/linha-do-tempo")
    assert resp.status_code == 200, resp.text
    return resp.json()


@pytest.mark.integration
def test_fonte_com_erro_nao_derruba(client_gestor, item_disponivel, monkeypatch):
    linha_do_tempo.carregar_fontes()
    monkeypatch.setattr(linha_do_tempo, "_FONTES", dict(linha_do_tempo._FONTES))

    @linha_do_tempo.registrar_fonte
    def fonte_quebrada(item_id):
        raise RuntimeError("falha de teste")

    corpo = _eventos(client_gestor, item_disponivel["id"])
    assert corpo["fontes_com_erro"] == ["fonte_quebrada"]
    assert "cadastro" in [e["tipo"] for e in corpo["eventos"]]


@pytest.mark.integration
def test_linha_do_tempo_completa_em_ordem(client_gestor, criar_unidade, executar_sql):
    criar_unidade(nome="Revalle Juazeiro")
    petrolina = criar_unidade(nome="Revalle Petrolina")
    resp = client_gestor.post(
        "/api/items",
        json={
            "tipo": "Notebook", "brand": "Dell", "model": "Latitude", "revenda": "Revalle Juazeiro",
            "date_registered": "01/01/2026", "identificador": "SN-LINHA-1",
        },
    )
    item_id = resp.json()["id"]
    assert client_gestor.put(f"/api/items/{item_id}", json={"brand": "Dell Inc"}).status_code == 200
    resp = client_gestor.post(
        f"/api/items/{item_id}/transferir",
        json={"para_unidade_id": petrolina["id"], "motivo": "Mudança de filial"},
    )
    assert resp.status_code == 200
    mid = client_gestor.post(
        "/api/manutencoes", json={"item_id": item_id, "defeito": "Tela quebrada"}
    ).json()["id"]
    resp = client_gestor.post(
        f"/api/manutencoes/{mid}/retorno", data={"resultado": "retornada", "custo": "150.00"}
    )
    assert resp.status_code == 200
    resp = client_gestor.request(
        "DELETE", f"/api/items/{item_id}", data={"reason": "Doação", **APAGAMENTO_OK}
    )
    assert resp.status_code == 200, resp.text

    # DATETIME tem precisão de segundos: regrava os horários na ordem das ações.
    def minuto(n):
        return f"2026-01-01 10:{n:02d}:00"

    executar_sql("UPDATE history SET data_operacao=%s WHERE item_id=%s AND operation='Cadastro'", (minuto(0), item_id))
    executar_sql(
        "UPDATE history SET data_operacao=%s WHERE item_id=%s AND operation='Edição' AND details IS NULL",
        (minuto(1), item_id),
    )
    executar_sql(
        "UPDATE history SET data_operacao=%s WHERE item_id=%s AND operation='Edição' AND details LIKE 'Transfer%%'",
        (minuto(2), item_id),
    )
    executar_sql("UPDATE transferencias SET transferido_em=%s WHERE item_id=%s", (minuto(2), item_id))
    executar_sql("UPDATE manutencoes SET aberta_em=%s, retornada_em=%s WHERE id=%s", (minuto(3), minuto(4), mid))
    executar_sql("UPDATE history SET data_operacao=%s WHERE item_id=%s AND operation='Exclusão'", (minuto(5), item_id))

    corpo = _eventos(client_gestor, item_id)
    eventos = corpo["eventos"]
    assert corpo["fontes_com_erro"] == []
    assert [e["tipo"] for e in eventos] == [
        "cadastro", "edicao", "transferencia", "manutencao_aberta", "manutencao_retornada", "baixa",
    ]
    por_tipo = {e["tipo"]: e for e in eventos}
    assert por_tipo["transferencia"]["titulo"] == "Transferência Revalle Juazeiro → Revalle Petrolina"
    assert por_tipo["transferencia"]["descricao"] == "Mudança de filial"
    assert por_tipo["transferencia"]["usuario"] == "gestor.teste"
    assert por_tipo["manutencao_aberta"]["descricao"] == "Tela quebrada"
    assert por_tipo["manutencao_retornada"]["descricao"] == "Custo: R$ 150,00"
    assert por_tipo["baixa"]["titulo"] == "Baixa do equipamento — Doação"
    assert por_tipo["baixa"]["descricao"].startswith("Dados apagados (Formatação segura) por Técnico TI em ")
    assert por_tipo["cadastro"]["titulo"] == "Cadastro do equipamento"
    assert corpo["item_id"] == item_id


@pytest.mark.integration
def test_eventos_de_manutencao_sem_conserto_e_cancelamento(client_gestor, criar_item):
    a, b = criar_item(), criar_item()
    m1 = client_gestor.post("/api/manutencoes", json={"item_id": a["id"], "defeito": "Placa"}).json()["id"]
    client_gestor.post(f"/api/manutencoes/{m1}/enviar", json={"assistencia": "TecInfo", "previsao_retorno": "31/12/2099"})
    client_gestor.post(f"/api/manutencoes/{m1}/em-reparo", json={})
    client_gestor.post(f"/api/manutencoes/{m1}/retorno", data={"resultado": "sem_conserto", "custo": "20.00"})
    m2 = client_gestor.post("/api/manutencoes", json={"item_id": b["id"], "defeito": "x"}).json()["id"]
    client_gestor.post(f"/api/manutencoes/{m2}/cancelar", json={"motivo": "Engano"})

    por_tipo = {e["tipo"]: e for e in _eventos(client_gestor, a["id"])["eventos"]}
    assert por_tipo["manutencao_enviada"]["titulo"] == "Enviado à assistência TecInfo"
    assert por_tipo["manutencao_enviada"]["descricao"] == "Previsão de retorno: 31/12/2099"
    assert por_tipo["manutencao_em_reparo"]["titulo"] == "Em reparo na assistência"
    assert por_tipo["manutencao_sem_conserto"]["titulo"] == "Manutenção encerrada sem conserto"
    assert por_tipo["manutencao_sem_conserto"]["descricao"] == "Baixa sugerida. Custo: R$ 20,00"
    cancelada = {e["tipo"]: e for e in _eventos(client_gestor, b["id"])["eventos"]}["manutencao_cancelada"]
    assert cancelada["titulo"] == "Manutenção cancelada"
    assert cancelada["usuario"] == "gestor.teste"


@pytest.mark.integration
def test_historico_mapeia_emprestimo(client_gestor, item_pendente):
    corpo = _eventos(client_gestor, item_pendente["id"])
    titulos = [e["titulo"] for e in corpo["eventos"]]
    assert "Cadastro do equipamento" in titulos
    assert "Empréstimo para Fulano de Tal" in titulos


@pytest.mark.integration
def test_historico_nao_inclui_entradas_estornadas(client_gestor, item_pendente, executar_sql):
    executar_sql(
        "UPDATE history SET is_reversed=1 WHERE item_id=%s AND operation='Empréstimo'", (item_pendente["id"],)
    )
    titulos = [e["titulo"] for e in _eventos(client_gestor, item_pendente["id"])["eventos"]]
    assert "Cadastro do equipamento" in titulos
    assert not any(t.startswith("Empréstimo") for t in titulos)


@pytest.mark.integration
def test_baixa_antiga_sem_registro_de_apagamento(client_gestor, inv_manager, criar_item):
    item = criar_item(tipo="Impressora")
    ok, msg = inv_manager.remove(item["id"], "teste", "Obsolescência")
    assert ok, msg
    baixa = [e for e in _eventos(client_gestor, item["id"])["eventos"] if e["tipo"] == "baixa"]
    assert len(baixa) == 1
    assert baixa[0]["descricao"] is None


@pytest.mark.integration
def test_item_baixado_tem_linha_do_tempo(client_gestor, inv_manager, criar_item):
    item = criar_item()
    ok, msg = inv_manager.remove(item["id"], "teste", "Perda")
    assert ok, msg
    assert client_gestor.get(f"/api/items/{item['id']}/linha-do-tempo").status_code == 200


@pytest.mark.integration
def test_item_inexistente_404(client_gestor):
    resp = client_gestor.get("/api/items/999999/linha-do-tempo")
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Item não encontrado."


@pytest.mark.integration
def test_exige_gestor_ou_tecnico(client_aprendiz, client_tecnico, item_disponivel):
    rota = f"/api/items/{item_disponivel['id']}/linha-do-tempo"
    resp = client_aprendiz.get(rota)
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
    assert client_tecnico.get(rota).status_code == 200
