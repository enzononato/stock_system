"""
Caracterização de autenticação: /api/auth/login, /me, /refresh, /logout.

Ver docs/TESTES.md, seção "Bugs corrigidos com teste de regressão", item 1: este
módulo encontrou (e o Módulo 3 confirmou de forma independente) um bug crítico no
claim "sub" do JWT que quebrava toda autenticação. O bug já estava corrigido no
código-alvo desta refatoração no momento em que isso foi reconciliado com o
coordenador (login agora grava "sub" como string; get_current_user faz
int(payload["sub"])). Este arquivo caracteriza o comportamento CORRETO/atual (token
de login autentica de verdade) e inclui um teste de regressão explícito para que
ninguém reintroduza o "sub" como inteiro.
"""
import pytest
from jose import jwt as jose_jwt

pytestmark = pytest.mark.integration


class TestLogin:
    def test_login_com_credenciais_validas_retorna_access_token(self, client, usuario_gestor, senha_padrao_teste):
        resp = client.post(
            "/api/auth/login",
            json={"username": usuario_gestor["username"], "password": senha_padrao_teste},
        )
        assert resp.status_code == 200
        body = resp.json()
        assert isinstance(body["access_token"], str) and body["access_token"]
        assert body["token_type"] == "bearer"
        # Refresh token vai num cookie HttpOnly, não no corpo da resposta.
        assert "refresh_token" in resp.cookies

    def test_login_com_senha_errada(self, client, usuario_gestor):
        resp = client.post(
            "/api/auth/login",
            json={"username": usuario_gestor["username"], "password": "senha-errada"},
        )
        assert resp.status_code == 401
        assert resp.json()["detail"] == "Usuário ou senha inválidos."

    def test_login_com_usuario_inexistente(self, client):
        resp = client.post(
            "/api/auth/login",
            json={"username": "nao.existe", "password": "qualquer"},
        )
        assert resp.status_code == 401
        assert resp.json()["detail"] == "Usuário ou senha inválidos."

    def test_regressao_sub_do_jwt_emitido_pelo_login_deve_ser_string(
        self, client, usuario_gestor, senha_padrao_teste
    ):
        """
        Teste de regressão para o BUG CORRIGIDO nº 1 (ver docs/TESTES.md): o claim
        "sub" do access_token emitido por /api/auth/login precisa ser uma STRING. Se
        alguém voltar a colocar o id inteiro do usuário ali, `jose.jwt.decode()` (usado
        por decode_token(), chamado por toda rota autenticada via get_current_user)
        volta a rejeitar TODO token com 401 "Token inválido ou expirado." — foi
        exatamente isso que quebrava a autenticação inteira antes da correção.

        De propósito, este teste não pressupõe COMO a correção foi feita: ele só lê
        (sem verificar assinatura/expiração — get_unverified_claims) o payload de um
        token de verdade devolvido pelo login e confere o tipo do "sub". Continua
        válido não importa se a correção usa str(id), um UUID, ou outra estratégia,
        desde que "sub" seja uma string.
        """
        resp = client.post(
            "/api/auth/login",
            json={"username": usuario_gestor["username"], "password": senha_padrao_teste},
        )
        assert resp.status_code == 200
        access_token = resp.json()["access_token"]

        payload = jose_jwt.get_unverified_claims(access_token)
        assert isinstance(payload["sub"], str), (
            f"O claim 'sub' do JWT voltou a não ser string (veio {payload['sub']!r}, "
            f"tipo {type(payload['sub']).__name__}) — isso reintroduz o bug crítico nº 1 "
            "que quebrava toda a autenticação (ver docs/TESTES.md)."
        )


class TestMe:
    def test_me_sem_token_retorna_401(self, client):
        resp = client.get("/api/auth/me")
        assert resp.status_code == 401
        assert resp.json()["detail"] == "Token de autenticação não fornecido."

    def test_me_com_usuario_autenticado_retorna_seus_proprios_dados(self, client_gestor, usuario_gestor):
        """client_gestor usa um token JWT real (login de verdade — ver conftest.py)."""
        resp = client_gestor.get("/api/auth/me")
        assert resp.status_code == 200
        body = resp.json()
        assert body["username"] == usuario_gestor["username"]
        assert body["role"] == "Gestor"
        assert body["id"] == usuario_gestor["id"]

    def test_token_de_login_real_autentica_em_rotas_protegidas(self, client, usuario_gestor, senha_padrao_teste):
        """
        Caracterização ponta a ponta via HTTP, sem nenhum contorno: faz login de
        verdade, pega o access_token de verdade que a API devolveu, e usa esse token
        exato em /api/auth/me e em outra rota protegida qualquer. Antes da correção do
        bug nº 1 (ver docs/TESTES.md), isso devolvia 401 "Token inválido ou expirado."
        mesmo com um token recém-emitido e válido; hoje autentica corretamente.
        """
        login_resp = client.post(
            "/api/auth/login",
            json={"username": usuario_gestor["username"], "password": senha_padrao_teste},
        )
        assert login_resp.status_code == 200
        access_token = login_resp.json()["access_token"]

        me_resp = client.get("/api/auth/me", headers={"Authorization": f"Bearer {access_token}"})
        assert me_resp.status_code == 200
        assert me_resp.json()["username"] == usuario_gestor["username"]

        # O mesmo token também autentica em outra rota protegida qualquer — só uma
        # amostra aqui (a matriz completa de RBAC por role está em test_rbac.py).
        items_resp = client.get("/api/items", headers={"Authorization": f"Bearer {access_token}"})
        assert items_resp.status_code == 200


class TestRefresh:
    def test_refresh_sem_cookie_retorna_401(self, client):
        resp = client.post("/api/auth/refresh")
        assert resp.status_code == 401
        assert resp.json()["detail"] == "Refresh token não encontrado."

    def test_refresh_com_cookie_de_login_real_emite_novo_access_token(
        self, client, usuario_gestor, senha_padrao_teste
    ):
        """
        O refresh_token emitido pelo login também tem "sub" como string (mesma
        correção do bug nº 1 — ver docs/TESTES.md), então /api/auth/refresh —
        que decodifica esse cookie com decode_token(..., expected_type="refresh") —
        funciona corretamente: devolve um novo access_token, também utilizável.
        """
        login_resp = client.post(
            "/api/auth/login",
            json={"username": usuario_gestor["username"], "password": senha_padrao_teste},
        )
        assert login_resp.status_code == 200
        assert "refresh_token" in client.cookies  # o TestClient guarda o cookie automaticamente

        refresh_resp = client.post("/api/auth/refresh")
        assert refresh_resp.status_code == 200
        novo_access_token = refresh_resp.json()["access_token"]
        assert isinstance(novo_access_token, str) and novo_access_token

        # O token novo também autentica de verdade.
        me_resp = client.get("/api/auth/me", headers={"Authorization": f"Bearer {novo_access_token}"})
        assert me_resp.status_code == 200


class TestLogout:
    def test_logout_nao_exige_autenticacao_e_sempre_retorna_sucesso(self, client):
        """logout() não tem nenhuma dependency de autenticação — funciona (e devolve
        sucesso) mesmo sem nunca ter havido login nesta sessão."""
        resp = client.post("/api/auth/logout")
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Logout realizado com sucesso."
