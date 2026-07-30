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

    def test_remover_usuario_inexistente_retorna_404(self, client_gestor):
        """
        CORRIGIDO NESTE CICLO (era o BUG CONHECIDO nº 6 para remove_user — ver
        docs/TESTES.md): `UserDBManager.remove_user()` em si ainda não checa
        existência/rowcount (esse remendo específico não mudou), mas o router
        (`app/routers/users.py::remove_user`) passou a buscar o usuário-alvo com
        `user_db.get_user_by_id(user_id)` ANTES de chamar remove_user() — parte da
        proteção de lockout (T5: Gestor não remove a si mesmo nem o último Gestor).
        Efeito colateral correto: um id inexistente agora é barrado nessa checagem
        e devolve 404 antes mesmo de chegar no DELETE sem side effect. Este teste
        afirma o comportamento CORRIGIDO e serve de regressão. (O mesmo bug em
        `update_password()`/`PUT /api/users/{id}/password` NÃO foi tocado — ver
        `test_atualizar_senha_de_usuario_inexistente_tambem_retorna_sucesso` logo
        abaixo, que continua caracterizando o comportamento antigo.)
        """
        resp = client_gestor.delete("/api/users/999999")
        assert resp.status_code == 404
        assert resp.json()["detail"] == "Usuário não encontrado."


class TestProtecaoDeLockout:
    """T2 (novo, T5 no código): app/routers/users.py::remove_user() tem duas
    proteções contra lockout administrativo -- um Gestor não pode remover a si
    mesmo, nem remover o último Gestor do sistema (o que deixaria ninguém capaz de
    administrar usuários). Com dois Gestores, remover um deles funciona
    normalmente."""

    def test_gestor_nao_remove_a_si_mesmo(self, client_gestor, usuario_gestor):
        resp = client_gestor.delete(f"/api/users/{usuario_gestor['id']}")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Você não pode remover seu próprio usuário."
        # Não removeu de verdade: o usuário continua listado.
        ids = {u["id"] for u in client_gestor.get("/api/users").json()}
        assert usuario_gestor["id"] in ids

    def test_com_dois_gestores_remover_um_funciona(self, client_gestor, usuario_gestor, criar_usuario, senha_padrao_teste):
        segundo_gestor = criar_usuario("segundo.gestor", senha_padrao_teste, "Gestor")

        resp = client_gestor.delete(f"/api/users/{segundo_gestor['id']}")
        assert resp.status_code == 200
        assert resp.json()["detail"] == "Usuário removido com sucesso."
        ids = {u["id"] for u in client_gestor.get("/api/users").json()}
        assert segundo_gestor["id"] not in ids
        assert usuario_gestor["id"] in ids  # o Gestor que fez a chamada continua existindo

    def test_nao_remove_o_ultimo_gestor(self, client_gestor, usuario_gestor, criar_usuario, senha_padrao_teste, executar_sql):
        """
        A checagem "não remove o último Gestor" (`count_gestores() <= 1`) só pode
        disparar quando o ALVO da remoção é um Gestor diferente de quem está
        logado -- mas como remover a si mesmo já é bloqueado ANTES dessa checagem
        (`user_id == current_user.id`), e qualquer outro Gestor distinto do
        logado implica pelo menos 2 Gestores no sistema (o logado + o alvo), a
        única forma de exercitar este ramo isoladamente é uma sessão cujo token
        ainda diz "Gestor" (o JWT não é revalidado contra o banco a cada request)
        mas cujo role no banco já não é mais Gestor -- daí o uso de `executar_sql`
        para rebaixar o role do usuário logado diretamente no banco, simulando
        esse cenário e comprovando que a proteção dedicada ao "último Gestor"
        (não apenas a de "não remove a si mesmo") também funciona.
        """
        outro_gestor = criar_usuario("outro.gestor", senha_padrao_teste, "Gestor")
        # Rebaixa o usuário logado (usuario_gestor) para Técnico diretamente no
        # banco -- o token JWT de client_gestor já foi emitido e continua
        # afirmando role=Gestor (get_current_user só decodifica o JWT, não
        # revalida contra o banco), então `gestor_only` ainda deixa passar, mas
        # `count_gestores()` agora enxerga só 1 Gestor de verdade (outro_gestor).
        executar_sql("UPDATE usuarios SET role='Técnico' WHERE id=%s", (usuario_gestor["id"],))

        resp = client_gestor.delete(f"/api/users/{outro_gestor['id']}")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Não é possível remover o último Gestor do sistema."


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

    def test_atualizar_senha_de_usuario_inexistente_retorna_404(self, client_gestor):
        """
        CORRIGIDO NESTE CICLO: `update_password()` não checava o id nem o `rowcount`,
        então respondia 200 "Senha alterada com sucesso" sem ter alterado nada. Agora
        o router devolve 404, coerente com remove_user, e a camada de dados também
        confere `rowcount` como defesa em profundidade.
        """
        resp = client_gestor.put("/api/users/999999/password", json={"new_password": "Qualquer#123"})
        assert resp.status_code == 404
        assert resp.json()["detail"] == "Usuário não encontrado."
