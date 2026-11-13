"""Fixtures das linhas telefônicas (F5).

Importadas pelos testes com `from tests.fixtures_linhas import *  # noqa: F401,F403`;
o `conftest.py` não muda.
"""
import itertools
import json
from datetime import datetime
from uuid import uuid4

import pytest

_contador_cpf = itertools.count(1)
_contador_linha = itertools.count(1)


@pytest.fixture
def linha_db(limpar_banco):
    from app.db.linha_db import LinhaDBManager

    return LinhaDBManager()


@pytest.fixture
def ctx_teste():
    from app.services.audit import AuditContext

    return AuditContext(usuario_id=None, username="teste", papel="Gestor", ip="127.0.0.1")


@pytest.fixture
def criar_colaborador(limpar_banco):
    from app.db.database_mysql import get_cursor

    def _criar(nome="Maria Souza", situacao="ativo") -> int:
        cpf = f"{next(_contador_cpf):011d}"
        agora = datetime.now()
        with get_cursor(dict_cursor=False) as cur:
            cur.execute(
                "INSERT INTO colaboradores (cpf, nome, situacao, origem, criado_em, atualizado_em) "
                "VALUES (%s, %s, %s, 'manual', %s, %s)",
                (cpf, nome, situacao, agora, agora),
            )
            return cur.lastrowid

    return _criar


@pytest.fixture
def criar_aparelho(criar_item):
    def _criar(tipo="Celular", brand="Samsung", model="A15", identificador=None) -> dict:
        return criar_item(
            tipo=tipo,
            brand=brand,
            model=model,
            identificador=identificador or f"APR-{uuid4().hex[:8].upper()}",
        )

    return _criar


@pytest.fixture
def dados_linha(unidade):
    def _dados(**overrides) -> dict:
        n = next(_contador_linha)
        dados = {
            "numero": f"7499998{n:04d}",
            "operadora": "Vivo",
            "plano": "Controle 20GB",
            "custo_mensal": "49.90",
            "unidade_id": unidade["id"],
        }
        dados.update(overrides)
        return dados

    return _dados


@pytest.fixture
def nova_linha(client_gestor, dados_linha):
    def _criar(**overrides) -> dict:
        resp = client_gestor.post("/api/linhas", json=dados_linha(**overrides))
        assert resp.status_code == 201, resp.text
        return resp.json()

    return _criar


@pytest.fixture
def eventos_no_banco(_test_schema):
    from app.db.database_mysql import get_cursor

    def _eventos(linha_id: int) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT * FROM linha_eventos WHERE linha_id = %s ORDER BY id", (linha_id,))
            return list(cur.fetchall())

    return _eventos


@pytest.fixture
def auditoria(_test_schema):
    from app.db.database_mysql import get_cursor

    def _auditoria(entidade: str, acao: str) -> list[dict]:
        with get_cursor(commit=False) as cur:
            cur.execute(
                "SELECT * FROM audit_log WHERE entidade = %s AND acao = %s ORDER BY id",
                (entidade, acao),
            )
            linhas = list(cur.fetchall())
        for linha in linhas:
            for chave in ("antes", "depois"):
                if isinstance(linha[chave], (str, bytes)):
                    linha[chave] = json.loads(linha[chave])
        return linhas

    return _auditoria
