"""Testes da Tarefa 3 (Fase 0) -- auditoria: registro (app/services/audit.py),
consulta (GET /api/audit) e contexto de auditoria (app/dependencies.py)."""
import json
from datetime import date, datetime
from decimal import Decimal

import pytest
from starlette.requests import Request

MSG_SO_GESTOR = "Acesso restrito ao Gestor."


def _contexto(**overrides):
    from app.services.audit import AuditContext

    dados = {"usuario_id": 7, "username": "gestor.teste", "papel": "Gestor", "ip": "10.1.2.3"}
    dados.update(overrides)
    return AuditContext(**dados)


def _requisicao(host, forwarded_for=None):
    headers = []
    if forwarded_for is not None:
        headers.append((b"x-forwarded-for", forwarded_for.encode()))
    return Request({"type": "http", "headers": headers, "client": (host, 50000)})


# ── Funções puras ───────────────────────────────────────────────────────────

def test_mascarar_cpf():
    from app.services.audit import mascarar

    assert mascarar({"cpf": "12345678909"}) == {"cpf": "***.456.789-**"}
    assert mascarar({"cpf": "123.456.789-09"}) == {"cpf": "***.456.789-**"}
    assert mascarar(None) is None


def test_mascarar_recursivo_e_sem_alterar_o_original():
    from app.services.audit import mascarar

    original = {"nome": "Fulano", "colaborador": {"cpf": "12345678909", "itens": [{"cpf": "98765432100"}]}}
    resultado = mascarar(original)
    assert resultado["nome"] == "Fulano"
    assert resultado["colaborador"]["cpf"] == "***.456.789-**"
    assert resultado["colaborador"]["itens"][0]["cpf"] == "***.654.321-**"
    assert original["colaborador"]["cpf"] == "12345678909"


def test_mascarar_cpf_malformado_nao_vaza():
    from app.services.audit import mascarar

    assert mascarar({"cpf": "123"}) == {"cpf": "***"}
    assert mascarar({"cpf": None}) == {"cpf": None}


def test_diff_so_mudancas():
    from app.services.audit import diff

    antes = {"nome": "A", "status": "Disponível", "marca": "Dell"}
    depois = {"nome": "A", "status": "Pendente", "marca": "Dell"}
    assert diff(antes, depois) == ({"status": "Disponível"}, {"status": "Pendente"})


def test_diff_serializa_datas_e_decimais():
    from app.services.audit import diff

    antes, depois = diff(
        {"valor": Decimal("10.50"), "data": date(2026, 1, 2), "quando": datetime(2026, 1, 2, 3, 4, 5)},
        {"valor": Decimal("12.00"), "data": date(2026, 1, 3), "quando": datetime(2026, 1, 2, 3, 4, 5)},
    )
    assert antes == {"valor": "10.50", "data": "2026-01-02"}
    assert depois == {"valor": "12.00", "data": "2026-01-03"}


def test_diff_com_um_lado_nulo():
    from app.services.audit import diff

    assert diff(None, {"a": 1}) == (None, {"a": 1})
    assert diff({"a": 1}, None) == ({"a": 1}, None)
    assert diff(None, None) == (None, None)


def test_sistema_context():
    from app.services.audit import AuditContext

    ctx = AuditContext.sistema()
    assert ctx.username == "sistema"
    assert ctx.usuario_id is None


# ── record ──────────────────────────────────────────────────────────────────

@pytest.mark.integration
def test_record_grava_linha(limpar_banco):
    from app.db.database_mysql import get_cursor
    from app.services.audit import record

    record(
        _contexto(), "item", 42, "editar",
        antes={"status": "Disponível", "cpf": "12345678909"},
        depois={"status": "Pendente", "obs": "ação"},
    )
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT * FROM audit_log")
        linhas = cur.fetchall()
    assert len(linhas) == 1
    linha = linhas[0]
    assert linha["entidade"] == "item"
    assert linha["entidade_id"] == "42"
    assert linha["acao"] == "editar"
    assert linha["usuario_id"] == 7
    assert linha["username"] == "gestor.teste"
    assert linha["papel"] == "Gestor"
    assert linha["ip"] == "10.1.2.3"
    assert linha["ocorrido_em"] is not None
    assert json.loads(linha["antes"]) == {"status": "Disponível", "cpf": "***.456.789-**"}
    assert json.loads(linha["depois"]) == {"status": "Pendente", "obs": "ação"}


@pytest.mark.integration
def test_record_com_cursor_usa_a_mesma_transacao(limpar_banco):
    from app.db.database_mysql import get_cursor
    from app.services.audit import record

    with pytest.raises(RuntimeError):
        with get_cursor() as cur:
            record(_contexto(), "unidade", 1, "criar", depois={"nome": "X"}, cur=cur)
            raise RuntimeError("falha depois do registro")
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT COUNT(*) AS n FROM audit_log")
        assert cur.fetchone()["n"] == 0


@pytest.mark.integration
def test_record_com_cursor_propaga_erro(limpar_banco):
    from app.db.database_mysql import get_cursor
    from app.services.audit import record

    with pytest.raises(Exception):
        with get_cursor() as cur:
            record(_contexto(), None, 1, "criar", cur=cur)  # entidade é NOT NULL


def test_record_sem_cursor_nao_propaga(monkeypatch):
    from app.services import audit

    def _quebra(*args, **kwargs):
        raise RuntimeError("banco fora do ar")

    monkeypatch.setattr(audit, "get_cursor", _quebra)
    audit.record(_contexto(), "item", 1, "criar")  # não levanta


# ── GET /api/audit ──────────────────────────────────────────────────────────

@pytest.mark.integration
def test_get_audit_exige_gestor(client_tecnico):
    resp = client_tecnico.get("/api/audit")
    assert resp.status_code == 403
    assert resp.json()["detail"] == MSG_SO_GESTOR


@pytest.mark.integration
def test_get_audit_exige_autenticacao(client):
    assert client.get("/api/audit").status_code == 401


@pytest.mark.integration
def test_get_audit_filtra_por_entidade(client_gestor):
    from app.services.audit import record

    record(_contexto(), "item", 1, "criar", depois={"tipo": "Notebook"})
    record(_contexto(), "item", 2, "criar", depois={"tipo": "Mouse"})
    record(_contexto(username="outro"), "unidade", 5, "editar", antes={"a": 1}, depois={"a": 2})

    resp = client_gestor.get("/api/audit", params={"entidade": "item"})
    assert resp.status_code == 200
    corpo = resp.json()
    assert corpo["total"] == 2
    assert {e["entidade"] for e in corpo["items"]} == {"item"}
    assert corpo["items"][0]["entidade_id"] == "2"  # mais recente primeiro
    assert corpo["items"][0]["depois"] == {"tipo": "Mouse"}

    resp = client_gestor.get("/api/audit", params={"entidade": "item", "entidade_id": "1"})
    assert resp.json()["total"] == 1

    resp = client_gestor.get("/api/audit", params={"username": "outro"})
    corpo = resp.json()
    assert corpo["total"] == 1
    assert corpo["items"][0]["antes"] == {"a": 1}
    assert corpo["items"][0]["ocorrido_em"].endswith(("+00:00", "Z"))


@pytest.mark.integration
def test_get_audit_filtra_por_periodo_e_pagina(client_gestor, executar_sql):
    from app.services.audit import record

    for i in range(3):
        record(_contexto(), "item", i, "criar")
    executar_sql("UPDATE audit_log SET ocorrido_em = '2026-01-10 12:00:00' WHERE entidade_id = '0'")
    executar_sql("UPDATE audit_log SET ocorrido_em = '2026-01-11 23:59:59' WHERE entidade_id = '1'")

    resp = client_gestor.get("/api/audit", params={"de": "2026-01-10", "ate": "2026-01-11"})
    assert resp.json()["total"] == 2

    resp = client_gestor.get("/api/audit", params={"ate": "2026-01-10"})
    assert resp.json()["total"] == 1

    resp = client_gestor.get("/api/audit", params={"limit": 1, "offset": 1})
    corpo = resp.json()
    assert corpo["total"] == 3
    assert len(corpo["items"]) == 1

    assert client_gestor.get("/api/audit", params={"de": "ontem"}).status_code == 422


# ── Contexto (IP) ───────────────────────────────────────────────────────────

def test_ip_x_forwarded_for_so_de_proxy_confiavel():
    from app.dependencies import CurrentUser, get_audit_context

    usuario = CurrentUser(id=3, username="tec", role="Técnico")

    # Cliente "testclient" (não é IP privado): o header é ignorado.
    ctx = get_audit_context(_requisicao("testclient", "1.2.3.4"), usuario)
    assert ctx.ip == "testclient"
    assert (ctx.usuario_id, ctx.username, ctx.papel) == (3, "tec", "Técnico")

    # IP público direto: o header é ignorado.
    assert get_audit_context(_requisicao("8.8.4.4", "1.2.3.4"), usuario).ip == "8.8.4.4"

    # Proxy do compose (rede privada) e loopback: usa o IP anexado pelo proxy.
    assert get_audit_context(_requisicao("172.18.0.5", "8.8.8.8"), usuario).ip == "8.8.8.8"
    assert get_audit_context(_requisicao("127.0.0.1", "8.8.8.8"), usuario).ip == "8.8.8.8"

    # Sem header: o IP da conexão.
    assert get_audit_context(_requisicao("172.18.0.5"), usuario).ip == "172.18.0.5"


@pytest.mark.parametrize(
    "peer, cadeia, esperado",
    [
        # Cadeia forjada pelo cliente: o nginx acrescenta o IP real à direita.
        ("172.18.0.5", "1.2.3.4, 8.8.4.4", "8.8.4.4"),
        # Proxies privados intermediários à direita são pulados.
        ("172.18.0.5", "1.2.3.4, 8.8.4.4, 10.0.0.7", "8.8.4.4"),
        # Entradas ilegíveis são puladas.
        ("172.18.0.5", "8.8.4.4, lixo", "8.8.4.4"),
        # Tudo confiável: vale a entrada válida mais à esquerda.
        ("172.18.0.5", "10.0.0.9, 192.168.1.4", "10.0.0.9"),
        # Header sem nenhum IP válido: o par direto.
        ("172.18.0.5", "lixo, outro", "172.18.0.5"),
        # Par público: o header é ignorado.
        ("8.8.4.4", "1.2.3.4, 8.8.8.8", "8.8.4.4"),
        ("testclient", "1.2.3.4", "testclient"),
    ],
)
def test_ip_real_do_cliente_nao_confia_no_inicio_da_cadeia(peer, cadeia, esperado):
    from app.dependencies import CurrentUser, get_audit_context, get_client_ip, ip_real_do_cliente

    req = _requisicao(peer, cadeia)
    assert ip_real_do_cliente(req) == esperado
    # O limitador de login e a auditoria seguem exatamente a mesma regra.
    assert get_client_ip(req) == esperado
    usuario = CurrentUser(id=1, username="u", role="Gestor")
    assert get_audit_context(req, usuario).ip == esperado


# ── Refinamentos da revisão ─────────────────────────────────────────────────

def test_mascarar_qualquer_chave_com_cpf():
    from app.services.audit import mascarar

    assert mascarar({"cpf_colaborador": "12345678909", "CPF": "98765432100"}) == {
        "cpf_colaborador": "***.456.789-**",
        "CPF": "***.654.321-**",
    }


def test_diff_chave_ausente_difere_de_none_e_ordem_estavel():
    from app.services.audit import diff

    antes, depois = diff({"b": None, "a": 1}, {"a": 1, "b": None, "c": None})
    assert antes == {"c": None}
    assert depois == {"c": None}

    antes, depois = diff({"z": 1, "y": 1, "x": 1}, {"z": 2, "y": 2, "x": 2})
    assert list(antes) == ["x", "y", "z"]
    assert list(depois) == ["x", "y", "z"]


@pytest.mark.integration
def test_record_grava_datas_em_iso_e_nao_propaga_erro_de_serializacao(limpar_banco):
    from app.db.database_mysql import get_cursor
    from app.services import audit

    audit.record(
        _contexto(), "item", 1, "editar",
        depois={"quando": datetime(2026, 1, 2, 3, 4, 5), "valor": Decimal("1.50")},
    )
    # Valor não serializável em JSON (chave não textual aninhada) não pode derrubar o chamador.
    audit.record(_contexto(), "item", 2, "editar", depois={"x": {(1, 2): 3}})
    with get_cursor(commit=False) as cur:
        cur.execute("SELECT depois FROM audit_log WHERE entidade_id = '1'")
        assert json.loads(cur.fetchone()["depois"]) == {"quando": "2026-01-02T03:04:05", "valor": "1.50"}
        cur.execute("SELECT COUNT(*) AS n FROM audit_log WHERE entidade_id = '2'")
        assert cur.fetchone()["n"] == 0


@pytest.mark.integration
def test_get_audit_periodo_invalido_retorna_422(client_gestor):
    for params in (
        {"ate": "9999-12-31"},
        {"de": "2026-02-01", "ate": "2026-01-01"},
        {"de": "0001-01-01"},
    ):
        resp = client_gestor.get("/api/audit", params=params)
        assert resp.status_code == 422, params
        assert resp.json()["detail"] == "Período inválido."
