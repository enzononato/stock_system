"""Tarefa 4 (Fase 0): cada endpoint de escrita existente grava uma linha em
`audit_log` (depois do sucesso, sem alterar a resposta) e GET /api/history
ganha o filtro `item_id`."""
import json
from datetime import datetime

import pytest

from tests.dados_baixa import APAGAMENTO_OK

pytestmark = pytest.mark.integration

ARQUIVO_PDF = {"signed_pdf": ("termo.pdf", b"conteudo-fake-pdf", "application/pdf")}
CPF_VALIDO = "11144477735"
CPF_MASCARADO = "***.444.777-**"
CPF_FIXTURE_MASCARADO = "***.222.333-**"  # CPF da fixture item_pendente (11122233344)


def _linhas(entidade=None, acao=None):
    from app.db.database_mysql import get_cursor

    sql = "SELECT * FROM audit_log WHERE 1=1"
    params = []
    if entidade:
        sql += " AND entidade = %s"
        params.append(entidade)
    if acao:
        sql += " AND acao = %s"
        params.append(acao)
    with get_cursor(commit=False) as cur:
        cur.execute(sql + " ORDER BY id", params)
        linhas = list(cur.fetchall())
    for linha in linhas:
        linha["antes"] = json.loads(linha["antes"]) if linha["antes"] else None
        linha["depois"] = json.loads(linha["depois"]) if linha["depois"] else None
    return linhas


def _unica(entidade, acao):
    linhas = _linhas(entidade, acao)
    assert len(linhas) == 1, linhas
    return linhas[0]


def _corpo_emprestimo(item_id):
    return {
        "item_id": item_id,
        "usuario": "Fulano de Tal",
        "cpf": CPF_VALIDO,
        "center_cost": "101 - Puxada",
        "cargo": "Analista",
        "setor": "TI",
        "revenda": "Revalle Juazeiro",
        "date_issue": datetime.now().strftime("%d/%m/%Y"),
    }


# ── Itens ───────────────────────────────────────────────────────────────────

def test_criar_item_grava(client_gestor):
    resp = client_gestor.post(
        "/api/items",
        json={
            "tipo": "Notebook", "brand": "Dell", "model": "X1", "revenda": "Revalle Juazeiro",
            "date_registered": "01/02/2026", "identificador": "SN-AUD-1",
        },
    )
    assert resp.status_code == 200
    item_id = resp.json()["id"]
    linha = _unica("item", "criar")
    assert linha["entidade_id"] == str(item_id)
    assert linha["username"] == "gestor.teste"
    assert linha["papel"] == "Gestor"
    assert linha["antes"] is None
    assert linha["depois"]["identificador"] == "SN-AUD-1"


def test_editar_item_grava_diff(client_gestor, item_disponivel):
    resp = client_gestor.put(
        f"/api/items/{item_disponivel['id']}", json={"setor": "Financeiro", "model": "Novo Modelo"}
    )
    assert resp.status_code == 200
    linha = _unica("item", "editar")
    assert linha["entidade_id"] == str(item_disponivel["id"])
    assert linha["username"] == "gestor.teste"
    assert linha["antes"] == {"model": item_disponivel["model"], "setor": "TI"}
    assert linha["depois"] == {"model": "Novo Modelo", "setor": "Financeiro"}


def test_excluir_item_grava_motivo(client_gestor, item_disponivel):
    # MUDANÇA INTENCIONAL (F4): Notebook exige registro de apagamento (LGPD)
    resp = client_gestor.request(
        "DELETE", f"/api/items/{item_disponivel['id']}", data={"reason": "Doação", **APAGAMENTO_OK}
    )
    assert resp.status_code == 200
    linha = _unica("item", "excluir")
    assert linha["entidade_id"] == str(item_disponivel["id"])
    assert linha["depois"]["motivo"] == "Doação"
    assert linha["antes"]["identificador"] == item_disponivel["identificador"]


def test_falha_nao_grava_auditoria(client_gestor, item_pendente):
    resp = client_gestor.request(
        "DELETE", f"/api/items/{item_pendente['id']}", data={"reason": "Perda"}
    )
    assert resp.status_code == 400
    assert _linhas("item", "excluir") == []


# ── Empréstimo e devolução ──────────────────────────────────────────────────

def test_emprestimo_grava_auditoria_com_cpf_mascarado(client_gestor, item_disponivel):
    resp = client_gestor.post("/api/loans", json=_corpo_emprestimo(item_disponivel["id"]))
    assert resp.status_code == 200
    linha = _unica("emprestimo", "criar")
    assert linha["entidade_id"] == str(item_disponivel["id"])
    assert linha["username"] == "gestor.teste"
    assert linha["antes"]["status"] == "Disponível"
    assert linha["depois"]["status"] == "Pendente"
    assert linha["depois"]["assigned_to"] == "Fulano de Tal"
    assert linha["depois"]["cpf"] == CPF_MASCARADO
    assert CPF_VALIDO not in json.dumps(linha["depois"])


def test_confirmar_emprestimo_grava(client_gestor, item_pendente):
    resp = client_gestor.post(f"/api/loans/{item_pendente['id']}/confirm", files=ARQUIVO_PDF)
    assert resp.status_code == 200
    linha = _unica("emprestimo", "confirmar")
    assert linha["entidade_id"] == str(item_pendente["id"])
    assert linha["antes"] == {"status": "Pendente"}
    assert linha["depois"]["status"] == "Indisponível"
    assert linha["depois"]["termo_assinado"]


def test_devolucao_grava_criar_e_confirmar(client_gestor, item_indisponivel, termo_devolucao_disponivel):
    item_id = item_indisponivel["id"]
    resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate")
    if not termo_devolucao_disponivel:
        assert resp.status_code == 400
        assert _linhas("devolucao") == []
        pytest.skip("Modelo de termo de devolução ausente neste ambiente.")
    assert resp.status_code == 200
    criar = _unica("devolucao", "criar")
    assert criar["entidade_id"] == str(item_id)
    assert criar["antes"]["status"] == "Indisponível"
    assert criar["depois"]["status"] == "Pendente Devolução"

    resp = client_gestor.post(f"/api/loans/{item_id}/return/confirm", files=ARQUIVO_PDF)
    assert resp.status_code == 200
    confirmar = _unica("devolucao", "confirmar")
    assert confirmar["antes"]["status"] == "Pendente Devolução"
    assert confirmar["depois"]["status"] == "Disponível"
    assert confirmar["antes"]["cpf"] == CPF_FIXTURE_MASCARADO


def test_confirmar_devolucao_grava(client_gestor, item_indisponivel, forcar_devolucao_iniciada):
    forcar_devolucao_iniciada(item_indisponivel["id"])
    resp = client_gestor.post(f"/api/loans/{item_indisponivel['id']}/return/confirm", files=ARQUIVO_PDF)
    assert resp.status_code == 200
    linha = _unica("devolucao", "confirmar")
    assert linha["entidade_id"] == str(item_indisponivel["id"])
    assert linha["depois"]["status"] == "Disponível"


def test_estorno_grava_entidade_historico(
    client_gestor, inv_manager, senha_padrao_teste, item_pendente
):
    linhas, _ = inv_manager.list_history(item_id=item_pendente["id"])
    hid = next(h["id"] for h in linhas if h["operation"] == "Empréstimo")
    resp = client_gestor.post(f"/api/history/{hid}/reverse", json={"password": senha_padrao_teste})
    assert resp.status_code == 200
    linha = _unica("historico", "estornar")
    assert linha["entidade_id"] == str(hid)
    assert linha["username"] == "gestor.teste"
    assert linha["antes"]["is_reversed"] == 0
    assert linha["depois"]["is_reversed"] == 1
    assert linha["depois"]["operation"] == "Empréstimo"
    assert linha["depois"]["item_id"] == item_pendente["id"]
    assert "password" not in json.dumps(linha, default=str)


# ── Periféricos ─────────────────────────────────────────────────────────────

def test_criar_periferico_grava(client_gestor):
    resp = client_gestor.post(
        "/api/peripherals",
        json={"tipo": "Mouse", "brand": "Logitech", "model": "M90", "identificador": "PER-AUD-1"},
    )
    assert resp.status_code == 200
    linha = _unica("periferico", "criar")
    assert linha["depois"]["identificador"] == "PER-AUD-1"
    assert linha["antes"] is None
    assert linha["entidade_id"].isdigit()


def test_desativar_periferico_grava(client_gestor, periferico_disponivel):
    resp = client_gestor.delete(f"/api/peripherals/{periferico_disponivel['id']}")
    assert resp.status_code == 200
    linha = _unica("periferico", "desativar")
    assert linha["entidade_id"] == str(periferico_disponivel["id"])
    assert linha["antes"] == {"is_active": 1}
    assert linha["depois"] == {"is_active": 0}


def test_vincular_desvincular_periferico_grava(client_gestor, item_disponivel, periferico_disponivel):
    perif_id = periferico_disponivel["id"]
    resp = client_gestor.post(f"/api/items/{item_disponivel['id']}/peripherals/{perif_id}")
    assert resp.status_code == 200
    vincular = _unica("periferico", "vincular")
    assert vincular["entidade_id"] == str(perif_id)
    assert vincular["antes"]["status"] == "Disponível"
    assert vincular["depois"]["status"] == "Em Uso"
    assert vincular["depois"]["equipamento_id"] == item_disponivel["id"]

    link_id = client_gestor.get(f"/api/items/{item_disponivel['id']}/peripherals").json()[0]["link_id"]
    resp = client_gestor.delete(f"/api/peripherals/links/{link_id}")
    assert resp.status_code == 200
    desvincular = _unica("periferico", "desvincular")
    assert desvincular["entidade_id"] == str(perif_id)
    assert desvincular["antes"]["status"] == "Em Uso"
    assert desvincular["antes"]["equipamento_id"] == item_disponivel["id"]
    assert desvincular["depois"]["status"] == "Disponível"
    assert desvincular["depois"]["equipamento_id"] is None


def test_substituir_periferico_grava(client_gestor, item_disponivel, criar_periferico):
    antigo = criar_periferico(identificador="AUD-ANTIGO")
    novo = criar_periferico(identificador="AUD-NOVO")
    client_gestor.post(f"/api/items/{item_disponivel['id']}/peripherals/{antigo['id']}")
    resp = client_gestor.post(
        f"/api/items/{item_disponivel['id']}/peripherals/{antigo['id']}/replace",
        data={"new_peripheral_id": str(novo["id"]), "reason": "Defeito"},
    )
    assert resp.status_code == 200
    linha = _unica("periferico", "substituir")
    assert linha["entidade_id"] == str(antigo["id"])
    assert linha["antes"]["status"] == "Em Uso"
    assert linha["depois"]["status"] == "Substituido"
    assert linha["depois"]["motivo_substituicao"] == "Defeito"
    assert linha["depois"]["substituido_por"] == novo["id"]
    assert linha["antes"]["equipamento_id"] == item_disponivel["id"]
    assert linha["depois"]["equipamento_id"] is None


# ── Usuários ────────────────────────────────────────────────────────────────

def test_criar_e_excluir_usuario_grava_sem_senha(client_gestor):
    resp = client_gestor.post(
        "/api/users", json={"username": "novo.aud", "password": "Senha#123Abc", "role": "Técnico"}
    )
    assert resp.status_code == 200
    criar = _unica("usuario", "criar")
    assert criar["depois"] == {"id": int(criar["entidade_id"]), "username": "novo.aud", "role": "Técnico"}

    resp = client_gestor.delete(f"/api/users/{criar['entidade_id']}")
    assert resp.status_code == 200
    excluir = _unica("usuario", "excluir")
    assert excluir["entidade_id"] == criar["entidade_id"]
    assert excluir["antes"]["username"] == "novo.aud"
    assert excluir["depois"] is None
    for linha in _linhas("usuario"):
        texto = json.dumps(linha, default=str)
        assert "Senha#123Abc" not in texto
        assert "$2" not in texto


def test_trocar_senha_nao_grava_hash(client_gestor, usuario_tecnico):
    resp = client_gestor.put(
        f"/api/users/{usuario_tecnico['id']}/password", json={"new_password": "SenhaNova#456"}
    )
    assert resp.status_code == 200
    linha = _unica("usuario", "editar")
    assert linha["entidade_id"] == str(usuario_tecnico["id"])
    assert linha["depois"] == {"senha": "alterada"}
    assert linha["antes"] is None
    texto = json.dumps(linha, default=str)
    assert "SenhaNova#456" not in texto
    assert "$2b$" not in texto and "$2a$" not in texto


# ── Unidades ────────────────────────────────────────────────────────────────

def _payload_unidade(nome):
    return {
        "nome": nome,
        "razao_social": "REVENDA VALLE DA INTEGRAÇÃO LTDA",
        "cnpj": "04.690.106/0001-15",
        "endereco": "Centro Industrial São Francisco, Quadra QID, Lotes 5/6 s/n",
        "cep": "48905-630",
        "cidade": "Juazeiro",
        "uf": "BA",
    }


def test_unidade_criar_e_editar_grava(client_gestor):
    resp = client_gestor.post("/api/unidades", json=_payload_unidade("Unidade Aud"))
    assert resp.status_code == 200
    uid = resp.json()["id"]
    criar = _unica("unidade", "criar")
    assert criar["entidade_id"] == str(uid)
    assert criar["depois"]["nome"] == "Unidade Aud"

    resp = client_gestor.put(f"/api/unidades/{uid}", json={"cidade": "Petrolina"})
    assert resp.status_code == 200
    editar = _unica("unidade", "editar")
    assert editar["antes"] == {"cidade": "Juazeiro"}
    assert editar["depois"] == {"cidade": "Petrolina"}


def test_unidade_desativar_grava(client_gestor, unidade):
    resp = client_gestor.delete(f"/api/unidades/{unidade['id']}")
    assert resp.status_code == 200
    linha = _unica("unidade", "desativar")
    assert linha["entidade_id"] == str(unidade["id"])
    assert linha["username"] == "gestor.teste"
    assert linha["antes"] == {"is_active": 1}
    assert linha["depois"] == {"is_active": 0}


def test_unidade_reativar_grava(client_gestor, unidade, executar_sql):
    executar_sql("UPDATE unidades SET is_active=0 WHERE id=%s", (unidade["id"],))
    resp = client_gestor.post(f"/api/unidades/{unidade['id']}/reativar")
    assert resp.status_code == 200
    linha = _unica("unidade", "reativar")
    assert linha["antes"] == {"is_active": 0}
    assert linha["depois"] == {"is_active": 1}


# ── Histórico por item ──────────────────────────────────────────────────────

def test_history_filtra_por_item_id(client_gestor, criar_item):
    a = criar_item(identificador="HIST-A")
    b = criar_item(identificador="HIST-B")

    resp = client_gestor.get("/api/history", params={"item_id": a["id"]})
    assert resp.status_code == 200
    corpo = resp.json()
    assert corpo["total"] == 1
    assert {h["item_id"] for h in corpo["items"]} == {a["id"]}

    # Combina com search (AND).
    resp = client_gestor.get("/api/history", params={"item_id": a["id"], "search": "HIST-B"})
    assert resp.json()["total"] == 0
    resp = client_gestor.get("/api/history", params={"item_id": b["id"], "search": "HIST-B"})
    assert resp.json()["total"] == 1

    # Sem o filtro, o comportamento continua o mesmo.
    assert client_gestor.get("/api/history").json()["total"] == 2


# ── Leitura do estado "depois" não derruba a operação já confirmada ─────────

def _falhar_a_partir_da_segunda_chamada(monkeypatch, classe, nome):
    original = getattr(classe, nome)
    chamadas = {"n": 0}

    def _metodo(self, *args, **kwargs):
        chamadas["n"] += 1
        if chamadas["n"] > 1:
            raise RuntimeError("banco indisponível")
        return original(self, *args, **kwargs)

    monkeypatch.setattr(classe, nome, _metodo)


def test_falha_ao_ler_depois_nao_derruba_editar_item(client_gestor, item_disponivel, monkeypatch):
    from app.db.inventory_manager_db import InventoryDBManager

    _falhar_a_partir_da_segunda_chamada(monkeypatch, InventoryDBManager, "find")
    resp = client_gestor.put(f"/api/items/{item_disponivel['id']}", json={"setor": "Financeiro"})
    assert resp.status_code == 200
    assert resp.json() == {"detail": f"Item {item_disponivel['id']} atualizado."}
    assert _linhas("item", "editar") == []


def test_falha_ao_ler_depois_nao_derruba_desativar_unidade(client_gestor, unidade, monkeypatch):
    from app.db.unidade_db import UnidadeDBManager

    _falhar_a_partir_da_segunda_chamada(monkeypatch, UnidadeDBManager, "get_unidade")
    resp = client_gestor.delete(f"/api/unidades/{unidade['id']}")
    assert resp.status_code == 200
    assert _linhas("unidade", "desativar") == []


def test_record_diff_mescla_extra_so_no_depois(monkeypatch):
    from app.services import audit

    gravado = {}
    monkeypatch.setattr(audit, "record", lambda ctx, e, i, a, antes=None, depois=None, cur=None: gravado.update(
        entidade=e, id=i, acao=a, antes=antes, depois=depois))
    ctx = audit.AuditContext(1, "u", "Gestor", "1.2.3.4")
    audit.record_diff(
        ctx, "item", 5, "confirmar", {"status": "Pendente", "x": 1},
        obter_depois=lambda: {"status": "Indisponível", "x": 1}, extra={"termo": "a.pdf"},
    )
    assert gravado == {
        "entidade": "item", "id": 5, "acao": "confirmar",
        "antes": {"status": "Pendente"}, "depois": {"status": "Indisponível", "termo": "a.pdf"},
    }


def test_record_diff_nunca_levanta(monkeypatch):
    from app.services import audit

    chamou = []
    monkeypatch.setattr(audit, "record", lambda *a, **k: chamou.append(1))

    def _quebra():
        raise RuntimeError("falhou")

    ctx = audit.AuditContext(1, "u", "Gestor", None)
    audit.record_diff(ctx, "item", 1, "editar", {"a": 1}, obter_depois=_quebra)
    assert chamou == []
    # Sem obter_depois: usa `depois` direto (criação).
    audit.record_diff(ctx, "item", 1, "criar", None, depois={"a": 1})
    assert chamou == [1]
