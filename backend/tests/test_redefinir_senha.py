"""F2 Task 3: esqueci minha senha (forgot/reset)."""
import json
import logging
import re
from datetime import datetime, timedelta, timezone

import pytest

from app.services.mail import MailError, limpar_outbox, outbox
from tests.fixtures_auth import cliente_com_rate_limit, linhas_auditoria, usuario_com_email  # noqa: F401

pytestmark = pytest.mark.integration

MSG_FORGOT = "Se o usuário existir e tiver e-mail cadastrado, enviaremos um link para redefinir a senha."
MSG_LINK_INVALIDO = "Link inválido ou expirado. Peça um novo."
MSG_SENHA_REDEFINIDA = "Senha redefinida. Entre com a nova senha."
MSG_SESSAO = "Sessão encerrada. Faça login novamente."


@pytest.fixture(autouse=True)
def _outbox_limpo():
    limpar_outbox()
    yield
    limpar_outbox()


def _agora_utc() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def _forgot(client, identificador):
    return client.post("/api/auth/forgot", json={"identificador": identificador})


def _token_do_email() -> str:
    m = re.search(r"token=([A-Za-z0-9_\-]+)", outbox[-1].text)
    assert m, outbox[-1].text
    return m.group(1)


def _linhas_tokens() -> list[dict]:
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute("SELECT * FROM password_reset_tokens ORDER BY id")
        return list(cur.fetchall())


def test_forgot_envia_email_com_link(client, usuario_com_email, monkeypatch):
    from app.core.security import sha256_hex

    monkeypatch.setattr("app.routers.auth_senha.settings.APP_PUBLIC_URL", "https://patrimonio.exemplo/")
    resp = _forgot(client, "tecnico.email")
    assert resp.status_code == 200
    assert resp.json() == {"detail": MSG_FORGOT}

    assert outbox[-1].to == ["tecnico@revalle.com.br"]
    assert "https://patrimonio.exemplo/redefinir-senha?token=" in outbox[-1].text
    token = _token_do_email()

    linhas = _linhas_tokens()
    assert len(linhas) == 1
    assert linhas[0]["token_hash"] == sha256_hex(token)
    assert token not in json.dumps(linhas, default=str)
    assert linhas[0]["expira_em"] - linhas[0]["criado_em"] == timedelta(minutes=30)

    auditoria = linhas_auditoria("usuario", "solicitar_redefinicao_senha")
    assert len(auditoria) == 1
    assert token not in json.dumps(auditoria, default=str)
    assert auditoria[0]["entidade_id"] == str(usuario_com_email["id"])


def test_forgot_por_email_sem_diferenciar_caixa(client, usuario_com_email):
    resp = _forgot(client, "TECNICO@Revalle.com.br")
    assert resp.status_code == 200
    assert len(outbox) == 1
    assert outbox[-1].to == ["tecnico@revalle.com.br"]


def test_forgot_usuario_inexistente_resposta_igual(client, usuario_com_email):
    existente = _forgot(client, "tecnico.email")
    limpar_outbox()
    from app.db.database_mysql import get_cursor

    with get_cursor() as cur:
        cur.execute("DELETE FROM password_reset_tokens")

    inexistente = _forgot(client, "ninguem.assim")
    assert (inexistente.status_code, inexistente.json()) == (existente.status_code, existente.json())
    assert outbox == []
    assert _linhas_tokens() == []


def test_forgot_usuario_sem_email_nao_envia(client, usuario_tecnico):
    resp = _forgot(client, "tecnico.teste")
    assert resp.status_code == 200
    assert resp.json() == {"detail": MSG_FORGOT}
    assert outbox == []
    assert _linhas_tokens() == []


def test_forgot_limite_de_3_por_hora(client, usuario_com_email):
    respostas = [_forgot(client, "tecnico.email") for _ in range(4)]
    assert {(r.status_code, json.dumps(r.json())) for r in respostas} == {(200, json.dumps({"detail": MSG_FORGOT}))}
    assert len(outbox) == 3


def test_forgot_falha_de_envio_responde_igual(client, usuario_com_email, monkeypatch, caplog):
    def _falha(*args, **kwargs):
        raise MailError("Não foi possível enviar o e-mail.")

    monkeypatch.setattr("app.routers.auth_senha.send_mail", _falha)
    with caplog.at_level(logging.DEBUG):
        resp = _forgot(client, "tecnico.email")
    assert resp.status_code == 200
    assert resp.json() == {"detail": MSG_FORGOT}

    token = _linhas_tokens()[0]["token_hash"]
    assert "tecnico@revalle.com.br" not in caplog.text
    assert token not in caplog.text
    assert "token=" not in caplog.text


def _pedir_token(client) -> str:
    assert _forgot(client, "tecnico.email").status_code == 200
    return _token_do_email()


def test_reset_troca_a_senha(client, usuario_com_email, senha_padrao_teste):
    token = _pedir_token(client)
    resp = client.post("/api/auth/reset", json={"token": token, "nova_senha": "NovaSenha#9"})
    assert resp.status_code == 200
    assert resp.json() == {"detail": MSG_SENHA_REDEFINIDA}

    antiga = client.post("/api/auth/login", json={"username": "tecnico.email", "password": senha_padrao_teste})
    assert antiga.status_code == 401
    nova = client.post("/api/auth/login", json={"username": "tecnico.email", "password": "NovaSenha#9"})
    assert nova.status_code == 200


def test_reset_token_reutilizado_recusado(client, usuario_com_email):
    token = _pedir_token(client)
    assert client.post("/api/auth/reset", json={"token": token, "nova_senha": "NovaSenha#9"}).status_code == 200
    resp = client.post("/api/auth/reset", json={"token": token, "nova_senha": "OutraSenha#9"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_LINK_INVALIDO


def test_reset_token_expirado_recusado(client, usuario_com_email, executar_sql):
    token = _pedir_token(client)
    executar_sql(
        "UPDATE password_reset_tokens SET expira_em = %s", (_agora_utc() - timedelta(minutes=1),)
    )
    resp = client.post("/api/auth/reset", json={"token": token, "nova_senha": "NovaSenha#9"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_LINK_INVALIDO


def test_reset_token_desconhecido_recusado(client, usuario_com_email):
    resp = client.post("/api/auth/reset", json={"token": "nao-existe", "nova_senha": "NovaSenha#9"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_LINK_INVALIDO


def test_reset_senha_em_branco_nao_gasta_o_token(client, usuario_com_email):
    token = _pedir_token(client)
    resp = client.post("/api/auth/reset", json={"token": token, "nova_senha": ""})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "A nova senha não pode estar em branco."
    assert client.post("/api/auth/reset", json={"token": token, "nova_senha": "NovaSenha#9"}).status_code == 200


def test_reset_invalida_outros_links_abertos(client, usuario_com_email):
    primeiro = _pedir_token(client)
    segundo = _pedir_token(client)
    assert primeiro != segundo
    assert client.post("/api/auth/reset", json={"token": segundo, "nova_senha": "NovaSenha#9"}).status_code == 200
    resp = client.post("/api/auth/reset", json={"token": primeiro, "nova_senha": "OutraSenha#9"})
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_LINK_INVALIDO


def test_reset_revoga_sessoes(client, usuario_com_email, senha_padrao_teste):
    login = client.post("/api/auth/login", json={"username": "tecnico.email", "password": senha_padrao_teste})
    access = login.json()["access_token"]
    cookie = login.cookies["refresh_token"]

    token = _pedir_token(client)
    assert client.post("/api/auth/reset", json={"token": token, "nova_senha": "NovaSenha#9"}).status_code == 200

    me = client.get("/api/auth/me", headers={"Authorization": f"Bearer {access}"})
    assert me.status_code == 401
    assert me.json()["detail"] == MSG_SESSAO

    client.cookies.clear()
    client.cookies.set("refresh_token", cookie)
    ref = client.post("/api/auth/refresh")
    assert ref.status_code == 401
    assert ref.json()["detail"] == MSG_SESSAO


def test_reset_grava_auditoria_sem_segredos(client, usuario_com_email):
    token = _pedir_token(client)
    assert client.post("/api/auth/reset", json={"token": token, "nova_senha": "NovaSenha#9"}).status_code == 200

    linhas = linhas_auditoria("usuario", "redefinir_senha")
    assert len(linhas) == 1
    assert linhas[0]["entidade_id"] == str(usuario_com_email["id"])
    assert linhas[0]["username"] == "tecnico.email"
    texto = json.dumps(linhas, default=str)
    assert token not in texto
    assert "NovaSenha#9" not in texto
    assert "$2b$" not in texto


def test_forgot_tem_rate_limit(usuario_com_email):
    with cliente_com_rate_limit("3/minute") as c:
        for _ in range(3):
            assert c.post("/api/auth/forgot", json={"identificador": "x"}).status_code == 200
        assert c.post("/api/auth/forgot", json={"identificador": "x"}).status_code == 429


def test_forgot_responde_antes_de_processar(client, usuario_com_email, monkeypatch):
    """Busca, limite, gravação, auditoria e e-mail ficam todos na BackgroundTask: a resposta
    não depende de nada disso (mesmo tempo para qualquer entrada)."""
    recebidos = []
    monkeypatch.setattr("app.routers.auth_senha._processar_pedido", lambda *a, **k: recebidos.append(a[0]))

    for ident in ("tecnico.email", "ninguem.assim"):
        resp = _forgot(client, ident)
        assert (resp.status_code, resp.json()) == (200, {"detail": MSG_FORGOT})

    assert recebidos == ["tecnico.email", "ninguem.assim"]
    assert outbox == [] and _linhas_tokens() == []


def test_limite_de_3_por_hora_e_atomico(usuario_com_email):
    from concurrent.futures import ThreadPoolExecutor

    from app.core.security import sha256_hex
    from app.db.password_reset_db import PasswordResetDBManager

    db = PasswordResetDBManager()
    agora = _agora_utc()

    def tentar(i):
        return db.criar_se_dentro_do_limite(
            usuario_com_email["id"], sha256_hex(f"tok-{i}"), agora, agora + timedelta(minutes=30),
            None, agora - timedelta(hours=1), 3,
        )

    with ThreadPoolExecutor(max_workers=8) as ex:
        resultados = list(ex.map(tentar, range(8)))

    assert resultados.count(True) == 3
    assert len(_linhas_tokens()) == 3


def test_aviso_de_boot_com_url_publica_local_em_producao(caplog, monkeypatch):
    import app.main as main_module

    monkeypatch.setattr("app.main.settings.ENVIRONMENT", "production")
    monkeypatch.setattr("app.main.settings.APP_PUBLIC_URL", "http://localhost:5173")
    with caplog.at_level(logging.WARNING):
        main_module.avisar_url_publica_local()
    assert "APP_PUBLIC_URL aponta para localhost em produção" in caplog.text

    caplog.clear()
    monkeypatch.setattr("app.main.settings.APP_PUBLIC_URL", "https://patrimonio.revalle.com.br")
    with caplog.at_level(logging.WARNING):
        main_module.avisar_url_publica_local()
    assert "APP_PUBLIC_URL" not in caplog.text

    monkeypatch.setattr("app.main.settings.APP_PUBLIC_URL", "http://localhost:5173")
    monkeypatch.setattr("app.main.settings.ENVIRONMENT", "development")
    caplog.clear()
    with caplog.at_level(logging.WARNING):
        main_module.avisar_url_publica_local()
    assert "APP_PUBLIC_URL" not in caplog.text
