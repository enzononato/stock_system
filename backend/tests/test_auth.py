"""
Caracterização de autenticação: /api/auth/login, /me, /refresh, /logout.

Ver conftest.py para a nota extensa sobre o BUG CONHECIDO (CRÍTICO) do claim "sub" do
JWT. Este arquivo contém a demonstração ponta a ponta, via HTTP real (sem nenhum
contorno), desse bug — é a evidência mais direta e importante de toda a suíte.
"""
import pytest

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


class TestMe:
    def test_me_sem_token_retorna_401(self, client):
        resp = client.get("/api/auth/me")
        assert resp.status_code == 401
        assert resp.json()["detail"] == "Token de autenticação não fornecido."

    def test_me_com_usuario_autenticado_via_dependency_override(self, client_gestor, usuario_gestor):
        """
        Caracteriza a LÓGICA do endpoint /me (dado um usuário autenticado, devolve seus
        dados) isolada do bug de JWT — ver test_token_de_login_real_nao_autentica_em_nenhuma_rota
        logo abaixo para o comportamento com um token *real*.
        """
        resp = client_gestor.get("/api/auth/me")
        assert resp.status_code == 200
        body = resp.json()
        assert body["username"] == usuario_gestor["username"]
        assert body["role"] == "Gestor"
        assert body["id"] == usuario_gestor["id"]

    def test_token_de_login_real_nao_autentica_em_nenhuma_rota(self, client, usuario_gestor, senha_padrao_teste):
        """
        BUG CONHECIDO (CRÍTICO) — caracterização ponta a ponta via HTTP, sem nenhum
        contorno: faz login de verdade, pega o access_token de verdade que a API
        devolveu, e usa esse token exato em /api/auth/me. O resultado hoje é 401,
        não 200 — ou seja, o próprio fluxo de login da aplicação gera tokens que a
        aplicação em seguida rejeita.

        Causa raiz (ver detalhes em test_unit_seguranca.py e conftest.py):
        app/routers/auth.py::login() monta `{"sub": user["id"], ...}` com um INT; a
        validação padrão de claims do python-jose (verify_sub=True) exige que "sub"
        seja uma string e rejeita com JWTClaimsError, que decode_token() converte em
        401 "Token inválido ou expirado.".

        Isso significa que, no sistema como está hoje, NENHUM usuário consegue de fato
        usar a API além do login — toda rota protegida por get_current_user (ou seja,
        praticamente toda a API) devolve 401 mesmo com um token recém-emitido e válido.
        """
        login_resp = client.post(
            "/api/auth/login",
            json={"username": usuario_gestor["username"], "password": senha_padrao_teste},
        )
        assert login_resp.status_code == 200
        access_token = login_resp.json()["access_token"]

        me_resp = client.get("/api/auth/me", headers={"Authorization": f"Bearer {access_token}"})
        assert me_resp.status_code == 401
        assert me_resp.json()["detail"] == "Token inválido ou expirado."

        # O mesmo token também não autentica em nenhuma outra rota protegida — só uma
        # amostra aqui (a matriz completa de RBAC está em test_rbac.py, usando o
        # contorno via dependency_override para caracterizar a lógica de roles em si).
        items_resp = client.get("/api/items", headers={"Authorization": f"Bearer {access_token}"})
        assert items_resp.status_code == 401
        assert items_resp.json()["detail"] == "Token inválido ou expirado."


class TestRefresh:
    def test_refresh_sem_cookie_retorna_401(self, client):
        resp = client.post("/api/auth/refresh")
        assert resp.status_code == 401
        assert resp.json()["detail"] == "Refresh token não encontrado."

    def test_refresh_com_cookie_de_login_real_tambem_falha(self, client, usuario_gestor, senha_padrao_teste):
        """
        Mesmo BUG CONHECIDO do "sub" inteiro (ver acima): o refresh_token emitido pelo
        login também tem `sub=user["id"]` (int), então /api/auth/refresh — que decodifica
        esse cookie com decode_token(..., expected_type="refresh") — falha do mesmo jeito.
        """
        login_resp = client.post(
            "/api/auth/login",
            json={"username": usuario_gestor["username"], "password": senha_padrao_teste},
        )
        assert login_resp.status_code == 200
        assert "refresh_token" in client.cookies  # o TestClient guarda o cookie automaticamente

        refresh_resp = client.post("/api/auth/refresh")
        assert refresh_resp.status_code == 401
        assert refresh_resp.json()["detail"] == "Token inválido ou expirado."


class TestLogout:
    def test_logout_nao_exige_autenticacao_e_sempre_retorna_sucesso(self, client):
        """logout() não tem nenhuma dependency de autenticação — funciona (e devolve
        sucesso) mesmo sem nunca ter havido login nesta sessão."""
        resp = client.post("/api/auth/logout")
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Logout realizado com sucesso."
