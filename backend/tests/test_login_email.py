"""Login por e-mail, Task 2: entrar com o e-mail e etapa `cadastro_email`."""
import re

import pytest

from app.services.mail import limpar_outbox, outbox
from tests.fixtures_auth import ativar_2fa_direto, chave_totp, codigo_totp  # noqa: F401

pytestmark = pytest.mark.integration

SENHA = "SenhaForte#123"
MSG_GENERICA = "Usuário ou senha inválidos."
MSG_USE_EMAIL = "Esta conta agora entra com o e-mail. Use o seu e-mail @revalle.com.br."
MSG_ETAPA_EXPIRADA = "Etapa de verificação expirada. Entre novamente."


@pytest.fixture(autouse=True)
def _outbox_limpo():
    limpar_outbox()
    yield
    limpar_outbox()


@pytest.fixture
def criar_conta_antiga(user_manager):
    """Conta como as de produção antes da mudança: sem e-mail confirmado."""

    def _criar(username: str, role: str = "Técnico", email: str | None = None) -> dict:
        ok, msg = user_manager.add_user(username, SENHA, role, email)
        assert ok, msg
        return user_manager.get_user_by_username(username)

    return _criar


def _login(client, identificador, senha=SENHA):
    return client.post("/api/auth/login", json={"username": identificador, "password": senha})


def _iniciar(client, mfa_token, email):
    return client.post("/api/auth/email/iniciar", json={"mfa_token": mfa_token, "email": email})


def _confirmar(client, mfa_token, codigo):
    return client.post("/api/auth/email/confirmar", json={"mfa_token": mfa_token, "codigo": codigo})


def _codigo_do_email() -> str:
    m = re.search(r"\b(\d{6})\b", outbox[-1].text)
    assert m, outbox[-1].text
    return m.group(1)


def _confirmar_email_pelo_login(client, identificador, email) -> dict:
    """Login + iniciar + confirmar; devolve o JSON da resposta do confirmar."""
    resp = _login(client, identificador)
    assert resp.status_code == 200, resp.text
    assert resp.json()["etapa"] == "cadastro_email", resp.text
    token = resp.json()["mfa_token"]
    ini = _iniciar(client, token, email)
    assert ini.status_code == 200, ini.text
    conf = _confirmar(client, token, _codigo_do_email())
    assert conf.status_code == 200, conf.text
    return conf.json()


def test_conta_antiga_por_usuario_vai_para_cadastro_email(client, criar_conta_antiga):
    criar_conta_antiga("joao", email="joao.sugerido@revalle.com.br")

    resp = _login(client, "joao")

    assert resp.status_code == 200, resp.text
    corpo = resp.json()
    assert corpo["etapa"] == "cadastro_email"
    assert corpo["access_token"] is None
    assert corpo["mfa_token"]
    assert corpo["email_sugerido"] == "joao.sugerido@revalle.com.br"
    assert corpo["expira_em_segundos"] == 15 * 60
    assert "refresh_token" not in resp.cookies


def test_conta_antiga_com_email_nao_confirmado_entra_pelo_email(client, criar_conta_antiga):
    """O e-mail cadastrado pelo Gestor (ainda não confirmado) também leva à etapa."""
    criar_conta_antiga("joao", email="joao@revalle.com.br")

    resp = _login(client, " Joao@Revalle.com.br ")

    assert resp.status_code == 200, resp.text
    assert resp.json()["etapa"] == "cadastro_email"
    assert "refresh_token" not in resp.cookies


def test_fluxo_tecnico_confirma_e_entra(client, criar_conta_antiga):
    criar_conta_antiga("joao")

    resp = _login(client, "joao")
    token = resp.json()["mfa_token"]
    ini = _iniciar(client, token, " Joao@Revalle.com.br ")
    assert ini.status_code == 200, ini.text
    assert ini.json() == {"detail": "Enviamos um código para joao@revalle.com.br."}
    assert outbox[-1].to == ["joao@revalle.com.br"]

    conf = _confirmar(client, token, _codigo_do_email())
    assert conf.status_code == 200, conf.text
    assert conf.json()["etapa"] == "ok"
    assert conf.json()["access_token"]
    assert "refresh_token" in conf.cookies

    # O token da etapa é de uso único.
    assert _confirmar(client, token, "000000").json()["detail"] == MSG_ETAPA_EXPIRADA

    depois = _login(client, "JOAO@REVALLE.COM.BR")
    assert depois.status_code == 200, depois.text
    assert depois.json()["etapa"] == "ok"
    assert depois.json()["access_token"]


def test_gestor_confirma_email_e_cai_em_cadastro_2fa(client, criar_conta_antiga, chave_totp):  # noqa: F811
    criar_conta_antiga("chefe", role="Gestor")

    corpo = _confirmar_email_pelo_login(client, "chefe", "chefe@revalle.com.br")

    assert corpo["etapa"] == "cadastro_2fa"
    assert corpo["access_token"] is None
    assert corpo["mfa_token"]
    # O token novo serve no cadastro do 2FA.
    cad = client.post("/api/auth/mfa/cadastro/iniciar", json={"mfa_token": corpo["mfa_token"]})
    assert cad.status_code == 200, cad.text


def test_gestor_com_2fa_confirma_email_e_segue_para_o_codigo(
    client, criar_conta_antiga, ativar_2fa_direto  # noqa: F811
):
    gestor = criar_conta_antiga("chefe", role="Gestor")
    ativar_2fa_direto(gestor["id"])

    resp = _login(client, "chefe")
    assert resp.json()["etapa"] == "cadastro_email"  # o e-mail vem antes do 2FA
    assert "refresh_token" not in resp.cookies

    token = resp.json()["mfa_token"]
    _iniciar(client, token, "chefe@revalle.com.br")
    conf = _confirmar(client, token, _codigo_do_email())
    assert conf.status_code == 200, conf.text
    assert conf.json()["etapa"] == "mfa"
    assert conf.json()["access_token"] is None
    assert "refresh_token" not in conf.cookies


def test_usuario_antigo_apos_confirmar_recebe_401_especifico(client, criar_conta_antiga):
    criar_conta_antiga("joao")
    _confirmar_email_pelo_login(client, "joao", "joao@revalle.com.br")

    certa = _login(client, "joao")
    assert certa.status_code == 401
    assert certa.json()["detail"] == MSG_USE_EMAIL

    errada = _login(client, "joao", "senha-errada")
    assert errada.status_code == 401
    assert errada.json()["detail"] == MSG_GENERICA

    for identificador in ("joao@revalle.com.br", "ninguem@revalle.com.br", "ninguem"):
        resp = _login(client, identificador, "senha-errada")
        assert resp.status_code == 401
        assert resp.json()["detail"] == MSG_GENERICA


def test_token_de_email_nao_serve_no_2fa_e_vice_versa(
    client, criar_conta_antiga, usuario_gestor, ativar_2fa_direto  # noqa: F811
):
    criar_conta_antiga("joao")
    token_email = _login(client, "joao").json()["mfa_token"]
    for rota, corpo in (
        ("/api/auth/mfa/verificar", {"mfa_token": token_email, "codigo": "123456"}),
        ("/api/auth/mfa/cadastro/iniciar", {"mfa_token": token_email}),
        ("/api/auth/mfa/cadastro/confirmar", {"mfa_token": token_email, "codigo": "123456"}),
    ):
        resp = client.post(rota, json=corpo)
        assert resp.status_code == 400, (rota, resp.text)
        assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA

    ativar_2fa_direto(usuario_gestor["id"])
    resp = _login(client, usuario_gestor["email"])
    assert resp.json()["etapa"] == "mfa"
    token_mfa = resp.json()["mfa_token"]
    ini = _iniciar(client, token_mfa, "gestor.teste@revalle.com.br")
    assert ini.status_code == 400
    assert ini.json()["detail"] == MSG_ETAPA_EXPIRADA
    conf = _confirmar(client, token_mfa, "123456")
    assert conf.status_code == 400
    assert conf.json()["detail"] == MSG_ETAPA_EXPIRADA
    assert outbox == []


def test_access_token_nao_serve_como_token_da_etapa(client, criar_conta_antiga, usuario_tecnico):
    access = _login(client, usuario_tecnico["email"]).json()["access_token"]
    resp = _iniciar(client, access, "tecnico.teste@revalle.com.br")
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA


def test_confirmar_esgotado_revoga_token(client, criar_conta_antiga):
    criar_conta_antiga("joao")
    token = _login(client, "joao").json()["mfa_token"]
    _iniciar(client, token, "joao@revalle.com.br")
    codigo = _codigo_do_email()
    errado = f"{(int(codigo) + 1) % 10**6:06d}"

    for _ in range(4):
        resp = _confirmar(client, token, errado)
        assert resp.status_code == 400
        assert resp.json()["detail"] == "Código inválido ou expirado."
    resp = _confirmar(client, token, errado)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Muitas tentativas. Entre de novo para receber outro código."

    resp = _iniciar(client, token, "joao@revalle.com.br")
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA
    resp = _confirmar(client, token, codigo)
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_ETAPA_EXPIRADA


def test_iniciar_erros_da_etapa(client, criar_conta_antiga):
    criar_conta_antiga("joao")
    criar_conta_antiga("maria", email="maria@revalle.com.br")
    token = _login(client, "joao").json()["mfa_token"]

    fora = _iniciar(client, token, "joao@gmail.com")
    assert fora.status_code == 422
    assert fora.json()["detail"] == "Use o seu e-mail @revalle.com.br."

    em_uso = _iniciar(client, token, "maria@revalle.com.br")
    assert em_uso.status_code == 409
    assert em_uso.json()["detail"] == "Este e-mail já está em uso por outra conta."

    assert _iniciar(client, token, "joao@revalle.com.br").status_code == 200
    cedo = _iniciar(client, token, "joao@revalle.com.br")
    assert cedo.status_code == 429
    assert cedo.json()["detail"] == "Aguarde 1 minuto para pedir outro código."


def test_me_traz_email(client, usuario_tecnico):
    access = _login(client, usuario_tecnico["email"]).json()["access_token"]
    resp = client.get("/api/auth/me", headers={"Authorization": f"Bearer {access}"})
    assert resp.status_code == 200
    assert resp.json()["email"] == "tecnico.teste@revalle.com.br"
    assert resp.json()["username"] == "tecnico.teste"


def test_forgot_ignora_email_nao_confirmado(client, criar_conta_antiga):
    criar_conta_antiga("joao", email="joao@revalle.com.br")

    for identificador in ("joao@revalle.com.br", "joao"):
        resp = client.post("/api/auth/forgot", json={"identificador": identificador})
        assert resp.status_code == 200
        assert resp.json() == {
            "detail": "Se o usuário existir e tiver e-mail cadastrado, enviaremos um link para redefinir a senha."
        }
    assert outbox == []
