"""F7 / Task 4: GET /api/painel/alertas (calculado na hora, sem tabela de alertas)."""
from datetime import date, datetime, timedelta

import pytest

from tests.fixtures_linhas import criar_aparelho, dados_linha, nova_linha  # noqa: F401
from tests.fixtures_termos import criar_colaborador  # noqa: F401

pytestmark = pytest.mark.integration


def _todos(client):
    resp = client.get("/api/painel/alertas")
    assert resp.status_code == 200, resp.text
    return resp.json()


def _alertas(client, tipo):
    return [a for a in _todos(client)["alertas"] if a["tipo"] == tipo]


def _garantia(executar_sql, item, dias):
    executar_sql(
        "UPDATE items SET garantia_ate=%s WHERE id=%s", (date.today() + timedelta(days=dias), item["id"])
    )


def _termo(item_id, *, status="manual", tipo="entrega", ha=timedelta(days=4), **extra):
    from app.core.tempo import agora_local
    from app.db.termo_db import TermoDBManager

    return TermoDBManager().criar(
        item_id=item_id, tipo=tipo, status=status, solicitado_por="gestor.teste",
        solicitado_em=agora_local() - ha, **extra,
    )


def _recuar_devolucao(executar_sql, item_id, dias, extra="", hours=0):
    executar_sql(
        "UPDATE history SET data_operacao=%s" + extra + " WHERE item_id=%s AND operation='Devolução'",
        (datetime.now() - timedelta(days=dias, hours=hours), item_id),
    )


# ── Garantia ────────────────────────────────────────────────────────────────

def test_garantia_vencendo_no_limite(client_gestor, criar_item, executar_sql):
    itens = {dias: criar_item() for dias in (0, 30, 31, -1)}
    for dias, item in itens.items():
        _garantia(executar_sql, item, dias)
    alertas = _alertas(client_gestor, "garantia_vencendo")
    assert [a["item_id"] for a in alertas] == [itens[0]["id"], itens[30]["id"]]
    assert [a["titulo"] for a in alertas] == ["Garantia vence hoje", "Garantia vence em 30 dia(s)"]
    assert alertas[0]["link"] == f"/items/{itens[0]['id']}"
    assert alertas[0]["severidade"] == "baixa"
    assert alertas[0]["descricao"].endswith("· REVALLE JUAZEIRO")
    assert alertas[0]["descricao"].startswith(f"#{itens[0]['id']} Notebook Dell Latitude 5420")


def test_garantia_de_item_baixado_nao_alerta(client_gestor, criar_item, executar_sql):
    item = criar_item()
    _garantia(executar_sql, item, 5)
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (item["id"],))
    assert _alertas(client_gestor, "garantia_vencendo") == []


# ── Termos ──────────────────────────────────────────────────────────────────

def test_termo_pendente_respeita_o_limiar_e_some_quando_assinado(client_gestor, criar_item):
    from app.db.termo_db import TermoDBManager

    a, b = criar_item(), criar_item()
    termo_a = _termo(a["id"], ha=timedelta(days=3, hours=1))
    _termo(b["id"], ha=timedelta(days=2, hours=23))

    alertas = _alertas(client_gestor, "termo_pendente")
    assert [x["item_id"] for x in alertas] == [a["id"]]
    assert alertas[0]["titulo"] == "Termo de entrega pendente há 3 dias"
    assert alertas[0]["link"] == f"/items/{a['id']}?aba=termos"
    assert alertas[0]["chave"] == f"termo_pendente:{termo_a}"
    assert alertas[0]["descricao"].endswith("· sem colaborador · Assinatura manual")

    TermoDBManager().atualizar(termo_a, status="assinado", assinado_em=datetime.now())
    assert _alertas(client_gestor, "termo_pendente") == []


def test_termo_agendado_conta_da_data_agendada(client_gestor, criar_item):
    from app.core.tempo import agora_local

    item = criar_item()
    _termo(
        item["id"], status="agendado", ha=timedelta(days=10),
        agendado_para=agora_local() + timedelta(days=1),
    )
    assert _alertas(client_gestor, "termo_pendente") == []


def test_termo_legado_sem_solicitacao(client_gestor, item_pendente, executar_sql):
    executar_sql(
        "UPDATE history SET data_operacao=%s WHERE item_id=%s AND operation='Empréstimo'",
        (datetime.now() - timedelta(days=4, hours=1), item_pendente["id"]),
    )
    alertas = _alertas(client_gestor, "termo_pendente")
    assert len(alertas) == 1
    assert alertas[0]["titulo"] == "Termo de entrega sem assinatura há 4 dias"
    assert alertas[0]["chave"] == f"termo_pendente:item{item_pendente['id']}"
    assert "Sem solicitação" in alertas[0]["descricao"]
    assert "Fulano de Tal" in alertas[0]["descricao"]


# ── Devolução ───────────────────────────────────────────────────────────────

def test_devolucao_pendente(client_gestor, criar_item, forcar_devolucao_iniciada, executar_sql):
    a, b, c = criar_item(), criar_item(), criar_item()
    forcar_devolucao_iniciada(a["id"])
    forcar_devolucao_iniciada(b["id"])
    _recuar_devolucao(executar_sql, a["id"], 8, hours=1)
    _recuar_devolucao(executar_sql, b["id"], 6)
    # Sem nenhuma linha `Devolução` no histórico (banco antigo).
    executar_sql("UPDATE items SET status='Pendente Devolução' WHERE id=%s", (c["id"],))

    alertas = _alertas(client_gestor, "devolucao_pendente")
    por_item = {x["item_id"]: x for x in alertas}
    assert set(por_item) == {a["id"], c["id"]}
    assert por_item[a["id"]]["titulo"] == "Devolução pendente há 8 dias"
    assert por_item[a["id"]]["severidade"] == "alta"
    assert por_item[c["id"]]["titulo"] == "Devolução pendente sem data registrada"
    assert por_item[c["id"]]["desde"] is None and por_item[c["id"]]["dias"] is None
    assert alertas[0]["item_id"] == c["id"]  # sem data primeiro


def test_devolucao_estornada_nao_conta(client_gestor, criar_item, forcar_devolucao_iniciada, executar_sql):
    a = criar_item()
    forcar_devolucao_iniciada(a["id"])
    _recuar_devolucao(executar_sql, a["id"], 20, ", is_reversed=1")
    alertas = _alertas(client_gestor, "devolucao_pendente")
    assert [x["titulo"] for x in alertas] == ["Devolução pendente sem data registrada"]


# ── Manutenção ──────────────────────────────────────────────────────────────

def test_manutencao_parada_usa_ultimo_movimento(client_gestor, criar_item, executar_sql):
    a, b = criar_item(), criar_item()
    ids = []
    for item in (a, b):
        resp = client_gestor.post("/api/manutencoes", json={"item_id": item["id"], "defeito": "Tela"})
        assert resp.status_code == 200, resp.text
        ids.append(resp.json()["id"])
    agora = datetime.now()
    executar_sql(
        "UPDATE manutencoes SET aberta_em=%s, enviada_em=%s, status='enviada' WHERE id=%s",
        (agora - timedelta(days=20), agora - timedelta(days=5), ids[0]),
    )
    executar_sql("UPDATE manutencoes SET aberta_em=%s WHERE id=%s", (agora - timedelta(days=16, hours=1), ids[1]))

    alertas = _alertas(client_gestor, "manutencao_parada")
    assert [x["item_id"] for x in alertas] == [b["id"]]
    assert alertas[0]["titulo"] == "Manutenção parada há 16 dias"
    assert alertas[0]["link"] == f"/items/{b['id']}?aba=manutencoes"
    assert alertas[0]["chave"] == f"manutencao_parada:{ids[1]}"
    assert "Aberta" in alertas[0]["descricao"] and "sem assistência" in alertas[0]["descricao"]

    resp = client_gestor.post(f"/api/manutencoes/{ids[1]}/cancelar", json={"motivo": "Teste"})
    assert resp.status_code == 200, resp.text
    assert _alertas(client_gestor, "manutencao_parada") == []


# ── Colaborador desligado ───────────────────────────────────────────────────

def test_colaborador_desligado_por_vinculo_e_por_cpf_legado(
    client_gestor, criar_colaborador, criar_item, executar_sql
):
    from app.db.acesso_db import AcessoDBManager

    colab = criar_colaborador(situacao="desligado", cpf="52998224725")
    a, b, c = criar_item(), criar_item(), criar_item()
    executar_sql(
        "UPDATE items SET status='Indisponível', colaborador_id=%s WHERE id=%s", (colab["id"], a["id"])
    )
    executar_sql(
        "UPDATE items SET status='Indisponível', colaborador_id=NULL, cpf='529.982.247-25' WHERE id=%s",
        (b["id"],),
    )
    executar_sql("UPDATE items SET colaborador_id=%s WHERE id=%s", (colab["id"], c["id"]))

    alertas = _alertas(client_gestor, "colaborador_desligado")
    assert {x["item_id"] for x in alertas} == {a["id"], b["id"]}
    assert {x["item_id"] for x in alertas} == {
        e["id"] for e in AcessoDBManager().equipamentos_com_colaborador(colab["id"])
    }
    for x in alertas:
        assert x["titulo"] == "Maria Souza (desligado) ainda está com equipamento"
        assert x["colaborador_id"] == colab["id"]
        assert x["severidade"] == "alta"
        assert x["link"] == f"/items/{x['item_id']}"
        assert x["descricao"].endswith("· Indisponível")


def test_colaborador_desligado_por_cpf_normaliza_os_dois_lados(
    client_gestor, criar_colaborador, criar_item, executar_sql
):
    criar_colaborador(situacao="desligado", cpf="52998224725")
    item = criar_item()
    executar_sql(
        "UPDATE items SET status='Indisponível', colaborador_id=NULL, cpf='529 982 247/25' WHERE id=%s",
        (item["id"],),
    )
    assert [x["item_id"] for x in _alertas(client_gestor, "colaborador_desligado")] == [item["id"]]


def test_colaborador_ativo_nao_alerta(client_gestor, criar_colaborador, criar_item, executar_sql):
    colab = criar_colaborador(situacao="ativo")
    a = criar_item()
    executar_sql(
        "UPDATE items SET status='Indisponível', colaborador_id=%s WHERE id=%s", (colab["id"], a["id"])
    )
    assert _alertas(client_gestor, "colaborador_desligado") == []


# ── Linha em aparelho baixado ───────────────────────────────────────────────

def test_linha_ativa_em_aparelho_baixado(client_gestor, criar_aparelho, nova_linha, executar_sql):
    a, b = criar_aparelho(), criar_aparelho()
    ativa = nova_linha(item_id=a["id"])
    suspensa = nova_linha(item_id=b["id"])
    resp = client_gestor.post(f"/api/linhas/{suspensa['id']}/status", json={"status": "suspensa"})
    assert resp.status_code == 200, resp.text
    executar_sql("UPDATE items SET is_active=0 WHERE id IN (%s, %s)", (a["id"], b["id"]))

    alertas = _alertas(client_gestor, "linha_aparelho_baixado")
    assert len(alertas) == 1
    assert alertas[0]["link"] == f"/linhas?busca={ativa['numero']}"
    assert alertas[0]["chave"] == f"linha_aparelho_baixado:{ativa['id']}"
    assert alertas[0]["item_id"] == a["id"]
    assert "ativa em aparelho baixado" in alertas[0]["titulo"]
    assert alertas[0]["descricao"].startswith(f"#{a['id']} Celular Samsung A15")


# ── Limiares, ordem e permissões ────────────────────────────────────────────

def test_limiares_lidos_na_hora(client_gestor, criar_item, executar_sql, monkeypatch):
    from app.core import config

    item = criar_item()
    _garantia(executar_sql, item, 10)
    monkeypatch.setattr(config.settings, "ALERTA_GARANTIA_DIAS", 5)
    assert _alertas(client_gestor, "garantia_vencendo") == []
    monkeypatch.setattr(config.settings, "ALERTA_GARANTIA_DIAS", 20)
    corpo = _todos(client_gestor)
    assert [a["item_id"] for a in corpo["alertas"] if a["tipo"] == "garantia_vencendo"] == [item["id"]]
    assert corpo["limites"]["garantia_dias"] == 20
    assert set(corpo["limites"]) == {
        "garantia_dias", "termo_pendente_dias", "devolucao_pendente_dias", "manutencao_dias",
    }


def test_ordem_e_contagens(
    client_gestor, criar_item, criar_colaborador, criar_aparelho, nova_linha,
    forcar_devolucao_iniciada, executar_sql,
):
    garantia = criar_item()
    _garantia(executar_sql, garantia, 3)
    _termo(criar_item()["id"], ha=timedelta(days=5))
    devolucao = criar_item()
    forcar_devolucao_iniciada(devolucao["id"])
    _recuar_devolucao(executar_sql, devolucao["id"], 9)
    manutencao = criar_item()
    resp = client_gestor.post("/api/manutencoes", json={"item_id": manutencao["id"], "defeito": "Tela"})
    executar_sql(
        "UPDATE manutencoes SET aberta_em=%s WHERE id=%s",
        (datetime.now() - timedelta(days=30), resp.json()["id"]),
    )
    colab = criar_colaborador(situacao="desligado")
    com_colab = criar_item()
    executar_sql(
        "UPDATE items SET status='Indisponível', colaborador_id=%s WHERE id=%s",
        (colab["id"], com_colab["id"]),
    )
    aparelho = criar_aparelho()
    nova_linha(item_id=aparelho["id"])
    executar_sql("UPDATE items SET is_active=0 WHERE id=%s", (aparelho["id"],))

    corpo = _todos(client_gestor)
    assert set(corpo["por_tipo"]) == {
        "colaborador_desligado", "devolucao_pendente", "termo_pendente",
        "manutencao_parada", "linha_aparelho_baixado", "garantia_vencendo",
    }
    assert all(n == 1 for n in corpo["por_tipo"].values())
    assert corpo["total"] == sum(corpo["por_tipo"].values()) == len(corpo["alertas"]) == 6
    assert corpo["alertas"][0]["severidade"] == "alta"
    ordem = [a["severidade"] for a in corpo["alertas"]]
    assert ordem == sorted(ordem, key=["alta", "media", "baixa"].index)
    assert corpo["alertas"][-1]["tipo"] == "garantia_vencendo"
    assert corpo["gerado_em"]


def test_garantia_mais_proxima_vem_primeiro(client_gestor, criar_item, executar_sql):
    longe, perto = criar_item(), criar_item()
    _garantia(executar_sql, longe, 25)
    _garantia(executar_sql, perto, 2)
    assert [a["item_id"] for a in _alertas(client_gestor, "garantia_vencendo")] == [perto["id"], longe["id"]]


def test_alertas_vazios(client_gestor):
    corpo = _todos(client_gestor)
    assert corpo["total"] == 0 and corpo["alertas"] == []
    assert corpo["por_tipo"] == {
        "colaborador_desligado": 0, "devolucao_pendente": 0, "termo_pendente": 0,
        "manutencao_parada": 0, "linha_aparelho_baixado": 0, "garantia_vencendo": 0,
    }


def test_alertas_so_gestor_e_tecnico(client_aprendiz, client_tecnico):
    resp = client_aprendiz.get("/api/painel/alertas")
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito a Gestor ou Técnico."
    assert client_tecnico.get("/api/painel/alertas").status_code == 200
