"""F2 Task 8: Gestor zera o 2FA de outro usuário e comando `reset_2fa` no servidor."""
import json

import pytest

from tests.fixtures_auth import (  # noqa: F401
    ativar_2fa_direto,
    chave_totp,
    codigo_totp,
    gestor_2fa,
    linhas_auditoria,
    login_com_2fa,
)

pytestmark = pytest.mark.integration

MSG_SESSAO = "Sessão encerrada. Faça login novamente."
MSG_MFA_SEM_2FA = "Este usuário não tem a verificação em duas etapas ativa."


def _linhas(sql, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


def _login(client, username, senha):
    return client.post("/api/auth/login", json={"username": username, "password": senha})


def test_gestor_zera_2fa_de_outro(gestor_2fa, client, usuario_tecnico, ativar_2fa_direto, senha_padrao_teste):
    dados = ativar_2fa_direto(usuario_tecnico["id"])
    alvo = usuario_tecnico["id"]

    resp = gestor_2fa["client"].delete(f"/api/users/{alvo}/2fa")
    assert resp.status_code == 200, resp.text
    assert resp.json() == {"detail": "Verificação em duas etapas zerada."}

    linha = _linhas("SELECT totp_ativo, totp_secret_cifrado FROM usuarios WHERE id = %s", (alvo,))[0]
    assert linha["totp_ativo"] == 0 and linha["totp_secret_cifrado"] is None
    assert _linhas("SELECT 1 FROM mfa_recovery_codes WHERE usuario_id = %s", (alvo,)) == []
    assert _login(client, "tecnico.teste", senha_padrao_teste).json()["etapa"] == "ok"

    auditoria = linhas_auditoria("usuario", "zerar_2fa")
    assert len(auditoria) == 1
    assert auditoria[0]["entidade_id"] == str(alvo)
    assert auditoria[0]["username"] == "gestor.teste"
    assert auditoria[0]["antes"] == {"totp_ativo": True}
    assert auditoria[0]["depois"] == {"totp_ativo": False, "origem": "tela de usuários"}
    texto = json.dumps(auditoria, default=str)
    assert dados.segredo not in texto and not any(c in texto for c in dados.codigos)


def test_zerar_2fa_encerra_sessoes_do_alvo(
    gestor_2fa, client, usuario_tecnico, ativar_2fa_direto, senha_padrao_teste
):
    dados = ativar_2fa_direto(usuario_tecnico["id"])
    access = login_com_2fa(client, "tecnico.teste", senha_padrao_teste, codigo_totp(dados.segredo))
    assert client.get("/api/auth/me", headers={"Authorization": f"Bearer {access}"}).status_code == 200

    assert gestor_2fa["client"].delete(f"/api/users/{usuario_tecnico['id']}/2fa").status_code == 200

    resp = client.get("/api/auth/me", headers={"Authorization": f"Bearer {access}"})
    assert resp.status_code == 401
    assert resp.json()["detail"] == MSG_SESSAO


def test_gestor_zerado_volta_ao_cadastro(gestor_2fa, client, criar_usuario, ativar_2fa_direto, senha_padrao_teste):
    outro = criar_usuario("gestor.dois", senha_padrao_teste, "Gestor")
    ativar_2fa_direto(outro["id"])
    assert _login(client, "gestor.dois", senha_padrao_teste).json()["etapa"] == "mfa"

    assert gestor_2fa["client"].delete(f"/api/users/{outro['id']}/2fa").status_code == 200

    assert _login(client, "gestor.dois", senha_padrao_teste).json()["etapa"] == "cadastro_2fa"


def test_zerar_2fa_exige_gestor(client_tecnico, usuario_gestor):
    resp = client_tecnico.delete(f"/api/users/{usuario_gestor['id']}/2fa")
    assert resp.status_code == 403
    assert resp.json()["detail"] == "Acesso restrito ao Gestor."


def test_zerar_proprio_2fa_recusado(gestor_2fa):
    resp = gestor_2fa["client"].delete(f"/api/users/{gestor_2fa['usuario']['id']}/2fa")
    assert resp.status_code == 400
    assert resp.json()["detail"] == (
        "Você não pode zerar a sua própria verificação em duas etapas. "
        "Peça a outro Gestor ou use o comando no servidor."
    )


def test_zerar_2fa_usuario_sem_2fa_400(gestor_2fa, usuario_tecnico):
    resp = gestor_2fa["client"].delete(f"/api/users/{usuario_tecnico['id']}/2fa")
    assert resp.status_code == 400
    assert resp.json()["detail"] == MSG_MFA_SEM_2FA


def test_zerar_2fa_usuario_inexistente_404(gestor_2fa):
    resp = gestor_2fa["client"].delete("/api/users/99999/2fa")
    assert resp.status_code == 404
    assert resp.json()["detail"] == "Usuário não encontrado."


def test_listagem_mostra_2fa_zerado(gestor_2fa, usuario_tecnico, ativar_2fa_direto):
    ativar_2fa_direto(usuario_tecnico["id"])
    c = gestor_2fa["client"]

    def _alvo():
        return next(u for u in c.get("/api/users").json() if u["id"] == usuario_tecnico["id"])

    assert _alvo()["totp_ativo"] is True
    assert c.delete(f"/api/users/{usuario_tecnico['id']}/2fa").status_code == 200
    assert _alvo()["totp_ativo"] is False


def test_script_reset_2fa_zera_e_audita(gestor_2fa, monkeypatch, capsys):
    from app.core import config
    from scripts.reset_2fa import main

    uid = gestor_2fa["usuario"]["id"]
    antes = _linhas("SELECT sessao_versao FROM usuarios WHERE id = %s", (uid,))[0]["sessao_versao"]
    # O script não lê a chave: sem ela funciona igual.
    monkeypatch.setattr(config.settings, "TOTP_ENCRYPTION_KEY", "")

    assert main(["gestor.teste"]) == 0

    linha = _linhas("SELECT totp_ativo, totp_secret_cifrado, sessao_versao FROM usuarios WHERE id = %s", (uid,))[0]
    assert linha["totp_ativo"] == 0 and linha["totp_secret_cifrado"] is None
    assert linha["sessao_versao"] == antes + 1
    assert _linhas("SELECT 1 FROM mfa_recovery_codes WHERE usuario_id = %s", (uid,)) == []

    auditoria = linhas_auditoria("usuario", "zerar_2fa")
    assert len(auditoria) == 1
    assert auditoria[0]["username"] == "sistema"
    assert auditoria[0]["entidade_id"] == str(uid)
    assert auditoria[0]["depois"]["origem"] == "script reset_2fa"
    assert "zerado" in capsys.readouterr().out


def test_script_reset_2fa_usuario_inexistente(client, capsys):
    from scripts.reset_2fa import main

    assert main(["nao.existe"]) == 1
    assert "Usuário 'nao.existe' não encontrado." in capsys.readouterr().out


def test_script_reset_2fa_sem_2fa(client, usuario_tecnico, capsys):
    from scripts.reset_2fa import main

    assert main(["tecnico.teste"]) == 0
    assert "não tem 2FA ativo" in capsys.readouterr().out
    assert linhas_auditoria("usuario", "zerar_2fa") == []
