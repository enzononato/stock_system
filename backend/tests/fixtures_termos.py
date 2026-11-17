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
    return {
        "event": tipo,
        "payload": {
            "id": int(documento_id),
            "externalId": f"patrimonio-termo-{termo_id}",
            "status": "COMPLETED",
        },
        "createdAt": "2026-11-12T10:00:00Z",
        "webhookEndpoint": "https://patrimonio.teste/api/integrations/revsign/webhook",
    }


def postar_webhook(client, corpo: Union[dict, bytes], segredo: Optional[str] = SEGREDO_WEBHOOK):
    """Posta no webhook, sem login. O cabeçalho do segredo só vai quando `segredo is not None`."""
    from app.services.revsign_client import CABECALHO_SEGREDO

    headers = {"Content-Type": "application/json"}
    if segredo is not None:
        headers[CABECALHO_SEGREDO] = segredo
    conteudo = corpo if isinstance(corpo, bytes) else json.dumps(corpo).encode()
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
