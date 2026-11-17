"""F7 / Task 2: fontes de termos, linha telefônica e inventário na linha do tempo do equipamento."""
from datetime import datetime

import pytest

from tests.fixtures_linhas import criar_aparelho, dados_linha, nova_linha  # noqa: F401
from tests.fixtures_termos import criar_colaborador, unidade_juazeiro  # noqa: F401

pytestmark = pytest.mark.integration


def _montar(item_id):
    from app.services.linha_do_tempo import montar_linha_do_tempo

    return montar_linha_do_tempo(item_id)


def _tipos(linha, origem):
    return [e.tipo for e in linha.eventos if e.origem == origem]


def test_fontes_da_f7_registradas(_test_schema):
    from app.services.linha_do_tempo import carregar_fontes, fontes_registradas

    carregar_fontes()
    nomes = fontes_registradas()
    assert nomes[:4] == [
        "fonte_historico", "fonte_manutencoes", "fonte_transferencias", "fonte_baixa",
    ]
    assert nomes[4:] == ["fonte_termos", "fonte_linha_telefonica", "fonte_inventario"]


def test_termo_solicitado_enviado_e_assinado(criar_item, criar_colaborador):
    from app.db.termo_db import TermoDBManager

    item = criar_item()
    c = criar_colaborador()
    db = TermoDBManager()
    termo_id = db.criar(
        item_id=item["id"], tipo="entrega", status="manual", solicitado_por="gestor.teste",
        solicitado_em=datetime(2026, 1, 1, 10, 0), colaborador_id=c["id"],
        destinatario_email="maria.souza@gmail.com",
    )
    db.atualizar(
        termo_id, status="assinado", enviado_em=datetime(2026, 1, 1, 10, 5),
        assinado_em=datetime(2026, 1, 2, 9, 0),
    )

    eventos = [e for e in _montar(item["id"]).eventos if e.origem == "termo"]
    assert [e.tipo for e in eventos] == ["termo_solicitado", "termo_enviado", "termo_assinado"]
    assert eventos[0].titulo == "Termo de entrega solicitado"
    assert eventos[0].usuario == "gestor.teste"
    assert eventos[1].titulo == "Termo de entrega enviado para assinatura"
    assert eventos[1].descricao == "Destinatário: Maria Souza"
    assert eventos[2].titulo == "Termo de entrega assinado"
    assert all(e.referencia_id == termo_id for e in eventos)
    for e in eventos:
        assert "@" not in (e.descricao or "") and "@" not in e.titulo


def test_termo_cancelado_mostra_motivo(criar_item, criar_colaborador):
    from app.db.termo_db import TermoDBManager

    item = criar_item()
    db = TermoDBManager()
    termo_id = db.criar(
        item_id=item["id"], tipo="devolucao", status="manual", solicitado_por="gestor.teste",
        solicitado_em=datetime(2026, 1, 1, 10, 0),
    )
    db.atualizar(
        termo_id, status="cancelado", cancelado_em=datetime(2026, 1, 3, 8, 0),
        motivo="Substituída pela solicitação #9.",
    )
    eventos = [e for e in _montar(item["id"]).eventos if e.tipo == "termo_cancelado"]
    assert len(eventos) == 1
    assert eventos[0].titulo == "Solicitação de termo cancelada"
    assert eventos[0].descricao == "Substituída pela solicitação #9."


def test_termo_agendado_e_recusado(criar_item):
    from app.db.termo_db import TermoDBManager

    item = criar_item()
    db = TermoDBManager()
    db.criar(
        item_id=item["id"], tipo="uso", status="agendado", solicitado_por="gestor.teste",
        solicitado_em=datetime(2026, 1, 1, 10, 0), agendado_para=datetime(2026, 1, 5, 14, 30),
    )
    termo2 = db.criar(
        item_id=item["id"], tipo="entrega", status="manual", solicitado_por="gestor.teste",
        solicitado_em=datetime(2026, 1, 2, 10, 0),
    )
    db.atualizar(termo2, status="recusado", motivo="Recusado pelo signatário.")
    eventos = [e for e in _montar(item["id"]).eventos if e.origem == "termo"]
    assert eventos[0].descricao == "Agendado para 05/01/2026 14:30"
    recusado = [e for e in eventos if e.tipo == "termo_recusado"]
    assert len(recusado) == 1 and recusado[0].referencia_id == termo2
    assert recusado[0].titulo == "Termo de entrega recusado"
    assert recusado[0].descricao == "Recusado pelo signatário."


def test_linha_que_passou_pelo_aparelho(client_gestor, criar_aparelho, nova_linha):
    a = criar_aparelho(identificador="CEL-A")
    b = criar_aparelho(identificador="CEL-B")
    linha = nova_linha(item_id=a["id"])
    assert client_gestor.put(f"/api/linhas/{linha['id']}", json={"item_id": b["id"]}).status_code == 200
    assert client_gestor.put(f"/api/linhas/{linha['id']}", json={"plano": "Outro plano"}).status_code == 200

    from app.schemas.linhas import formatar_numero_linha

    num = formatar_numero_linha(linha["numero"])
    de_a = _montar(a["id"])
    titulos_a = {e.tipo: e.titulo for e in de_a.eventos if e.origem == "linha"}
    assert titulos_a["linha_ativacao"] == f"Linha {num}: Ativação"
    assert titulos_a["linha_troca_aparelho"] == f"Linha {num} saiu deste aparelho"
    assert "linha_edicao" not in titulos_a
    de_b = _montar(b["id"])
    titulos_b = {e.tipo: e.titulo for e in de_b.eventos if e.origem == "linha"}
    assert titulos_b == {"linha_troca_aparelho": f"Linha {num} passou para este aparelho"}
    troca = next(e for e in de_b.eventos if e.tipo == "linha_troca_aparelho")
    assert "→" in troca.descricao


def _abrir(client, unidade_id):
    resp = client.post("/api/inventarios", json={"unidade_id": unidade_id})
    assert resp.status_code == 201, resp.text
    return resp.json()["id"]


def test_inventario_lido_e_nao_encontrado(client_tecnico, unidade, criar_item):
    item = criar_item(revenda=unidade["nome"])
    id1 = _abrir(client_tecnico, unidade["id"])
    resp = client_tecnico.post(f"/api/inventarios/{id1}/leituras", json={"codigo": f"#{item['id']}"})
    assert resp.status_code in (200, 201), resp.text
    client_tecnico.post(f"/api/inventarios/{id1}/encerrar")
    id2 = _abrir(client_tecnico, unidade["id"])
    client_tecnico.post(f"/api/inventarios/{id2}/encerrar")

    eventos = [e for e in _montar(item["id"]).eventos if e.origem == "inventario"]
    assert [e.tipo for e in eventos] == ["inventario_leitura", "inventario_nao_encontrado"]
    assert eventos[0].titulo.startswith(f"Inventário #{id1} (") and eventos[0].titulo.endswith("Conferido")
    assert eventos[1].titulo.startswith(f"Inventário #{id2} (") and eventos[1].titulo.endswith("Não encontrado")
    assert unidade["nome"].upper() in eventos[0].titulo
    assert eventos[0].usuario == "tecnico.teste"


def test_api_traz_fontes_da_f7_sem_erro(client_gestor, criar_item):
    from app.db.termo_db import TermoDBManager

    item = criar_item()
    TermoDBManager().criar(
        item_id=item["id"], tipo="entrega", status="manual", solicitado_por="gestor.teste",
        solicitado_em=datetime(2026, 1, 1, 10, 0),
    )
    resp = client_gestor.get(f"/api/items/{item['id']}/linha-do-tempo")
    assert resp.status_code == 200, resp.text
    corpo = resp.json()
    assert corpo["fontes_com_erro"] == []
    assert any(e["origem"] == "termo" for e in corpo["eventos"])
