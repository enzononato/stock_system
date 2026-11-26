"""Fixtures e ajudantes da Fase 3 (termos e RevSign).

Os arquivos de teste importam daqui com `# noqa: F401`; nada disto é registrado no conftest.
"""
import json
from datetime import datetime
from typing import Optional, Union

import pytest

SEGREDO_WEBHOOK = "segredo-webhook-de-teste-123"
URL_WEBHOOK = "/api/integrations/revsign/webhook"


@pytest.fixture
def revsign_fake(monkeypatch):
    """Liga a integração com o RevSign usando o cliente falso."""
    from app.core import config
    from app.services.revsign_client import FakeRevSignClient, usar_cliente

    monkeypatch.setattr(config.settings, "REVSIGN_ENABLED", True)
    monkeypatch.setattr(config.settings, "REVSIGN_WEBHOOK_SECRET", SEGREDO_WEBHOOK)
    # A conversão real exige o LibreOffice (lenta): nos testes do envio vai um PDF de mentira.
    monkeypatch.setattr("app.services.termos.docx_para_pdf", lambda conteudo, **kw: b"%PDF-1.4 teste")
    fake = FakeRevSignClient(webhook_secret=SEGREDO_WEBHOOK)
    usar_cliente(fake)
    yield fake
    usar_cliente(None)


@pytest.fixture
def criar_colaborador(_test_schema, limpar_banco):
    """Fábrica de colaboradores (cadastro manual) devolvendo a linha completa."""

    def _criar(**overrides):
        from app.db.colaborador_db import ColaboradorDBManager

        dados = {
            "cpf": "52998224725",
            "nome": "Maria Souza",
            "cargo": "Analista Financeiro",
            "setor": "Financeiro",
            "email": "maria.souza@gmail.com",
            "situacao": "ativo",
        }
        dados.update(overrides)
        db = ColaboradorDBManager()
        return db.obter(db.criar_manual(dados))

    return _criar


@pytest.fixture
def colaborador(criar_colaborador):
    return criar_colaborador()


@pytest.fixture
def unidade_juazeiro(criar_unidade):
    """Unidade 'Revalle Juazeiro' cadastrada: sem ela o termo de entrega não é gerado."""
    return criar_unidade(nome="Revalle Juazeiro")


def corpo_emprestimo_colaborador(item_id: int, colaborador_id: int, **extra) -> dict:
    corpo = {
        "item_id": item_id,
        "colaborador_id": colaborador_id,
        "center_cost": "101 - Puxada",
        "revenda": "Revalle Juazeiro",
        "date_issue": datetime.now().strftime("%d/%m/%Y"),
    }
    corpo.update(extra)
    return corpo


def evento_webhook(tipo: str, documento_id: str, termo_id: int) -> dict:
    """Corpo de webhook no formato do contrato (`termo.signed`, `termo.rejected`, ...)."""
    return {
        "eventId": "evt_teste",
        "event": tipo,
        "termoId": str(documento_id),
        "externalId": f"patrimonio-termo-{termo_id}",
        "status": "SIGNED",
        "occurredAt": "2026-11-12T10:00:00.000Z",
        "signer": {"name": "Maria Souza", "email": "maria.souza@gmail.com"},
    }


def postar_webhook(
    client,
    corpo: Union[dict, bytes],
    segredo: Optional[str] = SEGREDO_WEBHOOK,
    timestamp: Optional[int] = None,
):
    """Posta no webhook, sem login, assinado como o RevSign assina.

    Com `segredo=None` os cabeçalhos de assinatura não vão; `timestamp` (segundos Unix) é
    o de agora por padrão.
    """
    import time

    from app.services.revsign_client import assinar_webhook

    conteudo = corpo if isinstance(corpo, bytes) else json.dumps(corpo).encode()
    headers = {"Content-Type": "application/json"}
    if segredo is not None:
        headers.update(assinar_webhook(segredo, conteudo, int(time.time()) if timestamp is None else timestamp))
    return client.post(URL_WEBHOOK, content=conteudo, headers=headers)


def linhas_audit(entidade: str, acao: Optional[str] = None) -> list[dict]:
    from app.db.database_mysql import get_cursor

    sql = "SELECT * FROM audit_log WHERE entidade = %s"
    params: list = [entidade]
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
