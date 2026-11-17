"""Caracterização de backend/app/routers/users.py + UserDBManager (todas as rotas são
restritas ao Gestor — ver a matriz completa em test_rbac.py::TestRbacUsuarios)."""
import re

import pytest

from app.services.mail import MailError, limpar_outbox, outbox
from tests.fixtures_auth import ativar_2fa_direto, chave_totp, gestor_2fa, linhas_auditoria  # noqa: F401

pytestmark = pytest.mark.integration

MSG_GENERICA = "Usuário ou senha inválidos."
MSG_FALHA_CONVITE = "Conta criada, mas não foi possível enviar o e-mail. Use “Reenviar convite”."
MSG_LINK_INVALIDO = "Link inválido ou expirado. Peça um novo."


@pytest.fixture(autouse=True)
def _outbox_limpo():
    limpar_outbox()
    yield
    limpar_outbox()


def _token_do_email(indice: int = -1) -> str:
    m = re.search(r"token=([A-Za-z0-9_\-]+)", outbox[indice].text)
    assert m, outbox[indice].text
    return m.group(1)


def _criar(client_gestor, email: str, role: str = "Técnico", **extra):
    return client_gestor.post("/api/users", json={"email": email, "role": role, **extra})


def _conta(email: str) -> dict:
    from app.db.user_manager_db import UserDBManager

    return UserDBManager().get_user_para_login(email)


def _concluir(client, token: str, senha: str = "MinhaSenha#1"):
    return client.post("/api/auth/reset", json={"token": token, "nova_senha": senha})


def _login(client, identificador: str, senha: str = "MinhaSenha#1"):
    return client.post("/api/auth/login", json={"username": identificador, "password": senha})


class TestCriarUsuario:
    def test_criar_so_com_email_gera_username_e_envia_convite(self, client_gestor, monkeypatch):
        monkeypatch.setattr("app.routers.auth_senha.settings.APP_PUBLIC_URL", "https://patrimonio.exemplo/")
        resp = _criar(client_gestor, " Joao.Silva@Revalle.com.br ")

        assert resp.status_code == 200, resp.text
        assert resp.json()["convite_enviado"] is True
        conta = _conta("joao.silva@revalle.com.br")
        assert conta["username"] == "joao.silva"
        assert conta["role"] == "Técnico"
        assert conta["senha_definida"] == 0
        assert conta["email_confirmado"] == 0

        assert len(outbox) == 1
        convite = outbox[0]
        assert convite.to == ["joao.silva@revalle.com.br"]
        assert convite.subject == "Seu acesso ao Patrimônio Revalle"
        assert (
            "Você recebeu acesso ao sistema de Patrimônio da Revalle. Crie sua senha pelo botão abaixo. "
            "O link vale por 72 horas." in convite.text
        )
        assert "Criar minha senha" in convite.text
        assert re.search(r"https://patrimonio\.exemplo/redefinir-senha\?token=[A-Za-z0-9_\-]+&convite=1", convite.text)

        auditoria = linhas_auditoria("usuario", "enviar_convite")
        assert len(auditoria) == 1
        assert auditoria[0]["entidade_id"] == str(conta["id"])
        assert _token_do_email() not in str(auditoria[0])

    def test_criar_ignora_senha_enviada(self, client, client_gestor):
        """Não existe mais senha definida pelo Gestor: um `password` no corpo é ignorado."""
        resp = _criar(client_gestor, "maria@revalle.com.br", password="Abc#12345")
        assert resp.status_code == 200, resp.text
        assert _login(client, "maria@revalle.com.br", "Abc#12345").json()["detail"] == MSG_GENERICA

    def test_username_gerado_ganha_numero_quando_ja_existe(self, client_gestor, user_manager, senha_padrao_teste):
        user_manager.add_user("maria", senha_padrao_teste, "Técnico")
        user_manager.add_user("maria2", senha_padrao_teste, "Técnico")
        assert _criar(client_gestor, "maria@revalle.com.br").status_code == 200
        assert _conta("maria@revalle.com.br")["username"] == "maria3"

    def test_criar_com_email_de_fora_422(self, client_gestor):
        resp = _criar(client_gestor, "joao@gmail.com")
        assert resp.status_code == 422
        assert "Use o seu e-mail @revalle.com.br." in resp.json()["detail"]
        assert outbox == []

    def test_criar_sem_email_422(self, client_gestor):
        assert client_gestor.post("/api/users", json={"role": "Técnico"}).status_code == 422

    def test_criar_com_email_duplicado_400(self, client_gestor, usuario_tecnico):
        resp = _criar(client_gestor, "TECNICO.TESTE@revalle.com.br")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Este e-mail já está em uso por outro usuário."
        assert outbox == []

    def test_criar_usuario_com_role_fora_da_lista_permitida_e_422(self, client_gestor):
        """UserCreate.role é Literal["Gestor", "Técnico", "Jovem Aprendiz"]: qualquer
        outro valor é rejeitado pela validação do Pydantic antes de chegar no banco."""
        assert _criar(client_gestor, "x@revalle.com.br", role="Estagiário").status_code == 422

    def test_falha_no_envio_do_convite_mantem_conta(self, client_gestor, monkeypatch):
        def _falha(*args, **kwargs):
            raise MailError("SMTP fora do ar")

        monkeypatch.setattr("app.routers.auth_senha.send_mail", _falha)
        resp = _criar(client_gestor, "joao@revalle.com.br")

        assert resp.status_code == 200, resp.text
        assert resp.json() == {"detail": MSG_FALHA_CONVITE, "convite_enviado": False}
        assert _conta("joao@revalle.com.br")["senha_definida"] == 0

    def test_listar_usuarios_nao_expoe_hash_de_senha(self, client_gestor, usuario_gestor):
        resp = client_gestor.get("/api/users")
        assert resp.status_code == 200
        usuario = next(u for u in resp.json() if u["username"] == usuario_gestor["username"])
        assert "password" not in usuario
        assert set(usuario.keys()) == {
            "id", "username", "role", "email", "email_confirmado", "senha_definida", "totp_ativo",
        }

    def test_lista_traz_email_confirmado_e_senha_definida(self, client_gestor, usuario_gestor):
        _criar(client_gestor, "novo@revalle.com.br")
        lista = {u["username"]: u for u in client_gestor.get("/api/users").json()}
        assert lista["gestor.teste"]["email_confirmado"] is True
        assert lista["gestor.teste"]["senha_definida"] is True
        assert lista["novo"]["email_confirmado"] is False
        assert lista["novo"]["senha_definida"] is False

    def test_editar_email_zera_confirmacao(self, client_gestor, usuario_tecnico):
        resp = client_gestor.put(f"/api/users/{usuario_tecnico['id']}", json={"email": "Tec.Novo@revalle.com.br"})
        assert resp.status_code == 200, resp.text
        assert resp.json()["email"] == "tec.novo@revalle.com.br"
        assert resp.json()["email_confirmado"] is False

    def test_editar_email_de_fora_422(self, client_gestor, usuario_tecnico):
        resp = client_gestor.put(f"/api/users/{usuario_tecnico['id']}", json={"email": "tec@gmail.com"})
        assert resp.status_code == 422


class TestConvite:
    def test_login_antes_de_concluir_o_convite_401_generico(self, client, client_gestor, user_manager):
        _criar(client_gestor, "joao@revalle.com.br")
        conta = _conta("joao@revalle.com.br")
        # Mesmo que alguém acerte a senha (aqui forçada direto no banco), a conta sem
        # senha definida não entra.
        user_manager.update_password(conta["id"], "Forcada#123")
        for identificador in ("joao@revalle.com.br", "joao"):
            resp = _login(client, identificador, "Forcada#123")
            assert resp.status_code == 401
            assert resp.json()["detail"] == MSG_GENERICA

    def test_concluir_convite_confirma_email_e_entra(self, client, client_gestor):
        _criar(client_gestor, "joao@revalle.com.br")
        resp = _concluir(client, _token_do_email())
        assert resp.status_code == 200, resp.text

        conta = _conta("joao@revalle.com.br")
        assert conta["senha_definida"] == 1
        assert conta["email_confirmado"] == 1
        login = _login(client, "joao@revalle.com.br")
        assert login.status_code == 200, login.text
        assert login.json()["etapa"] == "ok"
        assert login.json()["access_token"]

    def test_gestor_conclui_convite_e_vai_para_cadastro_2fa(self, client, gestor_2fa):  # noqa: F811
        client_gestor = gestor_2fa["client"]  # com a chave TOTP ligada, o Gestor da sessão tem 2FA
        _criar(client_gestor, "chefe@revalle.com.br", role="Gestor")
        assert _concluir(client, _token_do_email()).status_code == 200

        login = _login(client, "chefe@revalle.com.br")
        assert login.status_code == 200, login.text
        assert login.json()["etapa"] == "cadastro_2fa"
        assert login.json()["access_token"] is None

    def test_convite_vale_72h(self, client, client_gestor, executar_sql):
        from app.db.database_mysql import get_cursor

        _criar(client_gestor, "joao@revalle.com.br")
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT TIMESTAMPDIFF(MINUTE, criado_em, expira_em) AS minutos FROM password_reset_tokens")
            assert cur.fetchone()["minutos"] == 72 * 60

        executar_sql("UPDATE password_reset_tokens SET expira_em = expira_em - INTERVAL 72 HOUR - INTERVAL 1 MINUTE")
        resp = _concluir(client, _token_do_email())
        assert resp.status_code == 400
        assert resp.json()["detail"] == MSG_LINK_INVALIDO
        assert _conta("joao@revalle.com.br")["senha_definida"] == 0

    def test_reenviar_invalida_link_anterior(self, client, client_gestor):
        _criar(client_gestor, "joao@revalle.com.br")
        antigo = _token_do_email()
        conta = _conta("joao@revalle.com.br")

        resp = client_gestor.post(f"/api/users/{conta['id']}/enviar-link-senha")
        assert resp.status_code == 200, resp.text
        assert len(outbox) == 2
        assert outbox[-1].subject == "Seu acesso ao Patrimônio Revalle"  # ainda é convite
        novo = _token_do_email()
        assert novo != antigo

        assert _concluir(client, antigo).status_code == 400
        assert _concluir(client, novo).status_code == 200
        assert len(linhas_auditoria("usuario", "enviar_convite")) == 2


class TestEnviarLinkSenha:
    def test_conta_com_senha_recebe_link_normal(self, client, client_gestor, usuario_tecnico):
        resp = client_gestor.post(f"/api/users/{usuario_tecnico['id']}/enviar-link-senha")
        assert resp.status_code == 200, resp.text
        assert resp.json()["detail"] == "Enviamos o link para tecnico.teste@revalle.com.br."

        assert len(outbox) == 1
        assert outbox[0].to == ["tecnico.teste@revalle.com.br"]
        assert outbox[0].subject == "Redefinição de senha — Patrimônio Revalle"
        assert "convite=1" not in outbox[0].text

        linhas = linhas_auditoria("usuario", "enviar_link_senha")
        assert len(linhas) == 1
        assert linhas[0]["entidade_id"] == str(usuario_tecnico["id"])
        assert "tecnico.teste@revalle.com.br" not in str(linhas[0])

        assert _concluir(client, _token_do_email(), "SenhaNova#456").status_code == 200
        assert _login(client, "tecnico.teste@revalle.com.br", "SenhaForte#123").status_code == 401
        assert _login(client, "tecnico.teste@revalle.com.br", "SenhaNova#456").json()["etapa"] == "ok"

    def test_link_novo_invalida_os_anteriores(self, client, client_gestor, usuario_tecnico):
        client.post("/api/auth/forgot", json={"identificador": "tecnico.teste@revalle.com.br"})
        assert len(outbox) == 1
        do_esqueci = _token_do_email()

        client_gestor.post(f"/api/users/{usuario_tecnico['id']}/enviar-link-senha")
        assert len(outbox) == 2
        assert _concluir(client, do_esqueci).status_code == 400
        assert _concluir(client, _token_do_email()).status_code == 200

    def test_link_para_email_trocado_pelo_gestor_confirma_o_email(self, client, client_gestor, usuario_tecnico):
        uid = usuario_tecnico["id"]
        client_gestor.put(f"/api/users/{uid}", json={"email": "tec.novo@revalle.com.br"})
        assert _conta("tec.novo@revalle.com.br")["email_confirmado"] == 0

        client_gestor.post(f"/api/users/{uid}/enviar-link-senha")
        assert outbox[-1].to == ["tec.novo@revalle.com.br"]
        assert _concluir(client, _token_do_email()).status_code == 200
        assert _conta("tec.novo@revalle.com.br")["email_confirmado"] == 1

    def test_trocar_email_invalida_links_abertos(self, client, client_gestor, usuario_tecnico):
        uid = usuario_tecnico["id"]
        client_gestor.post(f"/api/users/{uid}/enviar-link-senha")
        token = _token_do_email()
        client_gestor.put(f"/api/users/{uid}", json={"email": "tec.novo@revalle.com.br"})

        resp = _concluir(client, token)
        assert resp.status_code == 400
        assert _conta("tec.novo@revalle.com.br")["email_confirmado"] == 0

    def test_enviar_link_sem_email_400(self, client_gestor, usuario_tecnico, executar_sql):
        executar_sql(
            "UPDATE usuarios SET email = NULL, email_confirmado = 0 WHERE id = %s", (usuario_tecnico["id"],)
        )
        resp = client_gestor.post(f"/api/users/{usuario_tecnico['id']}/enviar-link-senha")
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Cadastre o e-mail desta conta primeiro."
        assert outbox == []

    def test_enviar_link_usuario_inexistente_404(self, client_gestor):
        resp = client_gestor.post("/api/users/999999/enviar-link-senha")
        assert resp.status_code == 404
        assert resp.json()["detail"] == "Usuário não encontrado."

    def test_enviar_link_exige_gestor(self, client, client_tecnico, client_aprendiz, usuario_gestor):
        url = f"/api/users/{usuario_gestor['id']}/enviar-link-senha"
        assert client_tecnico.post(url).status_code == 403
        assert client_aprendiz.post(url).status_code == 403
        assert client.post(url).status_code == 401
        assert outbox == []

    def test_falha_no_envio_503(self, client_gestor, usuario_tecnico, monkeypatch):
        def _falha(*args, **kwargs):
            raise MailError("SMTP fora do ar")

        monkeypatch.setattr("app.routers.auth_senha.send_mail", _falha)
        resp = client_gestor.post(f"/api/users/{usuario_tecnico['id']}/enviar-link-senha")
        assert resp.status_code == 503
        assert resp.json()["detail"] == "Não foi possível enviar o e-mail agora. Tente de novo em instantes."

    def test_endpoint_de_senha_do_gestor_removido(self, client_gestor, usuario_tecnico):
        resp = client_gestor.put(
            f"/api/users/{usuario_tecnico['id']}/password", json={"new_password": "SenhaNova#456"}
        )
        assert resp.status_code in (404, 405)


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
        afirma o comportamento CORRIGIDO e serve de regressão. (A troca de senha pelo
        Gestor, `PUT /api/users/{id}/password`, foi removida: hoje o Gestor só envia
        um link, ver `TestEnviarLinkSenha`.)
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


