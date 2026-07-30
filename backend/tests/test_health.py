"""T2 (novo, T11 no código): /api/health responde 200 sempre que o processo está de
pé, mesmo sem tocar no banco; /api/health/db reporta indisponibilidade (503) sem
derrubar o healthcheck principal quando o MySQL está inacessível."""
import pytest

pytestmark = pytest.mark.integration


class TestHealth:
    def test_health_responde_200_sem_tocar_banco(self, client):
        resp = client.get("/api/health")
        assert resp.status_code == 200
        assert resp.json() == {"status": "ok"}

    def test_health_nao_exige_autenticacao(self, client):
        # Nenhum header Authorization enviado, e ainda assim 200 (não é uma rota
        # protegida -- um healthcheck de infraestrutura não pode exigir login).
        resp = client.get("/api/health")
        assert resp.status_code == 200


class TestHealthDb:
    def test_health_db_responde_ok_quando_banco_disponivel(self, client):
        resp = client.get("/api/health/db")
        assert resp.status_code == 200
        body = resp.json()
        assert body["status"] == "ok"
        assert body["database"] == "reachable"

    def test_health_db_reporta_indisponibilidade_quando_banco_falha(self, client, monkeypatch):
        """Simula falha de conexão sem derrubar o banco de teste de verdade:
        monkeypatcha `get_inventory_db` no NAMESPACE de app.main (onde health_db()
        chama a função diretamente, sem Depends()) para levantar uma exceção,
        exatamente como aconteceria se o MySQL estivesse fora do ar."""
        import app.main as main_module

        def _levantar_erro_de_conexao():
            raise RuntimeError("Falha simulada de conexão com o banco de teste")

        monkeypatch.setattr(main_module, "get_inventory_db", _levantar_erro_de_conexao)

        resp = client.get("/api/health/db")
        assert resp.status_code == 503
        body = resp.json()
        assert body["status"] == "unavailable"
        assert body["database"] == "unreachable"
        assert "Falha simulada de conexão" in body["detail"]

        # O healthcheck principal continua respondendo 200 -- não foi derrubado
        # pela falha simulada do /api/health/db.
        assert client.get("/api/health").status_code == 200
