"""Caracterização de backend/app/routers/users.py + UserDBManager (todas as rotas são
restritas ao Gestor — ver a matriz completa em test_rbac.py::TestRbacUsuarios)."""
import pytest

pytestmark = pytest.mark.integration


class TestCriarUsuario:
    def test_criar_usuario_com_sucesso(self, client_gestor):
        resp = client_gestor.post(
            "/api/users", json={"username": "novo.usuario", "password": "Senha#123", "role": "Técnico"}
        )
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Usuário 'novo.usuario' cadastrado com sucesso."

    def test_criar_usuario_com_username_duplicado(self, client_gestor, usuario_gestor):
        resp = client_gestor.post(
            "/api/users",
            json={"username": usuario_gestor["username"], "password": "OutraSenha#1", "role": "Técnico"},
        )
        assert resp.status_code == 400
        assert resp.json()["detail"] == f"O nome de usuário '{usuario_gestor['username']}' já existe."

    def test_criar_usuario_com_role_fora_da_lista_permitida_e_422(self, client_gestor):
        """UserCreate.role é Literal["Gestor", "Técnico", "Jovem Aprendiz"]: qualquer
        outro valor é rejeitado pela validação do Pydantic antes de chegar no banco."""
        resp = client_gestor.post(
            "/api/users", json={"username": "x", "password": "Senha#123", "role": "Estagiário"}
        )
        assert resp.status_code == 422

    def test_listar_usuarios_nao_expoe_hash_de_senha(self, client_gestor, usuario_gestor):
        resp = client_gestor.get("/api/users")
        assert resp.status_code == 200
        usuario = next(u for u in resp.json() if u["username"] == usuario_gestor["username"])
        assert "password" not in usuario
        assert set(usuario.keys()) == {"id", "username", "role"}


class TestRemoverUsuario:
    def test_remover_usuario_existente(self, client_gestor, usuario_tecnico):
        resp = client_gestor.delete(f"/api/users/{usuario_tecnico['id']}")
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Usuário removido com sucesso."
        ids_restantes = {u["id"] for u in client_gestor.get("/api/users").json()}
        assert usuario_tecnico["id"] not in ids_restantes

    def test_remover_usuario_inexistente_tambem_retorna_sucesso(self, client_gestor):
        """
        BUG CONHECIDO: UserDBManager.remove_user() executa um DELETE sem checar antes se
        o id existe e sem olhar `cur.rowcount` — um DELETE que não afeta nenhuma linha não
        levanta exceção no MySQL, então a função sempre devolve (True, "Usuário removido
        com sucesso."), mesmo quando nenhum usuário foi de fato removido. A API não tem
        como o cliente distinguir "removi o usuário X" de "usuário X nunca existiu".
        """
        resp = client_gestor.delete("/api/users/999999")
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Usuário removido com sucesso."


class TestAtualizarSenha:
    def test_atualizar_senha_com_sucesso_permite_login_com_a_nova_senha(
        self, client, client_gestor, usuario_tecnico
    ):
        resp = client_gestor.put(
            f"/api/users/{usuario_tecnico['id']}/password", json={"new_password": "SenhaNova#456"}
        )
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Senha alterada com sucesso."

        login_com_senha_antiga = client.post(
            "/api/auth/login", json={"username": usuario_tecnico["username"], "password": "SenhaForte#123"}
        )
        assert login_com_senha_antiga.status_code == 401

        login_com_senha_nova = client.post(
            "/api/auth/login", json={"username": usuario_tecnico["username"], "password": "SenhaNova#456"}
        )
        assert login_com_senha_nova.status_code == 200

    def test_atualizar_senha_em_branco(self, client_gestor, usuario_tecnico):
        resp = client_gestor.put(f"/api/users/{usuario_tecnico['id']}/password", json={"new_password": ""})
        assert resp.status_code == 400
        assert resp.json()["detail"] == "A nova senha não pode estar em branco."

    def test_atualizar_senha_de_usuario_inexistente_tambem_retorna_sucesso(self, client_gestor):
        """Mesma classe de BUG CONHECIDO do remove_user(): update_password() não checa
        se o id existe nem o rowcount, então sempre reporta sucesso."""
        resp = client_gestor.put("/api/users/999999/password", json={"new_password": "Qualquer#123"})
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Senha alterada com sucesso."
