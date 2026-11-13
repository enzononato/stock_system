"""Catálogo de sistemas e pacotes por função (/api/acessos/sistemas e /pacotes)."""
import pytest

from app.schemas.acessos import MSG_SEM_SENHA
from tests.acessos_helpers import auditoria, contar, criar_sistema

pytestmark = pytest.mark.integration


def _nomes(resposta):
    return [s["nome"] for s in resposta]


def test_tecnico_lista_so_sistemas_ativos(client_tecnico):
    criar_sistema("AD")
    criar_sistema("Promax", ativo=False)

    r = client_tecnico.get("/api/acessos/sistemas")
    assert r.status_code == 200
    assert _nomes(r.json()) == ["AD"]

    r = client_tecnico.get("/api/acessos/sistemas?include_inactive=true")
    assert _nomes(r.json()) == ["AD", "Promax"]
    assert r.json()[1]["ativo"] is False


def test_aprendiz_nao_acessa_o_catalogo(client_aprendiz):
    r = client_aprendiz.get("/api/acessos/sistemas")
    assert r.status_code == 403
    assert r.json()["detail"] == "Acesso restrito a Gestor ou Técnico."


def test_criar_sistema_exige_gestor(client_tecnico):
    r = client_tecnico.post("/api/acessos/sistemas", json={"nome": "GLPI"})
    assert r.status_code == 403
    assert r.json()["detail"] == "Acesso restrito ao Gestor."


def test_criar_sistema(client_gestor):
    r = client_gestor.post("/api/acessos/sistemas", json={"nome": " GLPI ", "descricao": ""})
    assert r.status_code == 201
    corpo = r.json()
    assert corpo["nome"] == "GLPI"
    assert corpo["descricao"] is None
    assert corpo["ativo"] is True
    linhas = auditoria("acesso_sistema", "criar")
    assert len(linhas) == 1
    assert linhas[0]["username"] == "gestor.teste"


def test_sistema_duplicado_409(client_gestor):
    client_gestor.post("/api/acessos/sistemas", json={"nome": "GLPI"})
    r = client_gestor.post("/api/acessos/sistemas", json={"nome": "glpi"})
    assert r.status_code == 409
    assert r.json()["detail"] == "Já existe um sistema com esse nome."


def test_desativar_e_reativar_sistema(client_gestor):
    sid = criar_sistema("AD")
    r = client_gestor.put(f"/api/acessos/sistemas/{sid}", json={"ativo": False})
    assert r.status_code == 200
    assert r.json()["ativo"] is False
    assert len(auditoria("acesso_sistema", "desativar")) == 1
    assert client_gestor.get("/api/acessos/sistemas").json() == []

    r = client_gestor.put(f"/api/acessos/sistemas/{sid}", json={"ativo": True})
    assert r.json()["ativo"] is True
    assert len(auditoria("acesso_sistema", "reativar")) == 1


def test_editar_sistema_audita_diff(client_gestor):
    sid = criar_sistema("AD")
    r = client_gestor.put(f"/api/acessos/sistemas/{sid}", json={"descricao": "Rede"})
    assert r.json()["descricao"] == "Rede"
    linha = auditoria("acesso_sistema", "editar")[0]
    assert linha["depois"] == {"descricao": "Rede"}


def test_editar_sistema_404_e_corpo_vazio(client_gestor):
    r = client_gestor.put("/api/acessos/sistemas/999", json={"ativo": False})
    assert r.status_code == 404
    assert r.json()["detail"] == "Sistema não encontrado."
    sid = criar_sistema("AD")
    r = client_gestor.put(f"/api/acessos/sistemas/{sid}", json={})
    assert r.status_code == 400
    assert r.json()["detail"] == "Nenhum campo para atualizar."


def test_senha_no_corpo_do_sistema_422(client_gestor):
    r = client_gestor.post("/api/acessos/sistemas", json={"nome": "X", "senha": "123"})
    assert r.status_code == 422
    assert r.json()["detail"] == MSG_SEM_SENHA
    assert contar("sistemas", "nome='X'") == 0


def test_definir_pacote_normaliza_e_substitui(client_gestor):
    ad, email, vpn = criar_sistema("AD"), criar_sistema("E-mail"), criar_sistema("VPN")
    r = client_gestor.put(
        "/api/acessos/pacotes", json={"cargo": " Vendedor ", "sistema_ids": [ad, email, ad]}
    )
    assert r.status_code == 200
    assert r.json()["cargo"] == "Vendedor"
    assert _nomes(r.json()["sistemas"]) == ["AD", "E-mail"]

    client_gestor.put("/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": [vpn]})
    pacotes = client_gestor.get("/api/acessos/pacotes").json()
    assert len(pacotes) == 1
    assert pacotes[0]["cargo"] == "Vendedor"
    assert _nomes(pacotes[0]["sistemas"]) == ["VPN"]


def test_pacote_com_cargo_em_outra_caixa_substitui_o_mesmo_pacote(client_gestor):
    ad, vpn = criar_sistema("AD"), criar_sistema("VPN")
    client_gestor.put("/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": [ad]})
    client_gestor.put("/api/acessos/pacotes", json={"cargo": "VENDEDOR ", "sistema_ids": [vpn]})
    pacotes = client_gestor.get("/api/acessos/pacotes").json()
    assert len(pacotes) == 1
    assert _nomes(pacotes[0]["sistemas"]) == ["VPN"]


def test_pacote_vazio_remove_o_cargo(client_gestor):
    ad = criar_sistema("AD")
    client_gestor.put("/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": [ad]})
    r = client_gestor.put("/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": []})
    assert r.status_code == 200
    assert client_gestor.get("/api/acessos/pacotes").json() == []


def test_pacote_com_sistema_desativado_ou_inexistente_422(client_gestor):
    ad = criar_sistema("AD")
    inativo = criar_sistema("Promax", ativo=False)
    client_gestor.put("/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": [ad]})
    for ruim in (inativo, 9999):
        r = client_gestor.put(
            "/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": [ad, ruim]}
        )
        assert r.status_code == 422
        assert r.json()["detail"] == "Sistema inválido ou desativado."
    pacotes = client_gestor.get("/api/acessos/pacotes").json()
    assert _nomes(pacotes[0]["sistemas"]) == ["AD"]


def test_pacote_tecnico_le_mas_nao_edita(client_tecnico):
    assert client_tecnico.get("/api/acessos/pacotes").status_code == 200
    r = client_tecnico.put("/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": []})
    assert r.status_code == 403
    assert r.json()["detail"] == "Acesso restrito ao Gestor."


def test_pacote_audita_antes_e_depois(client_gestor):
    ad, email, vpn = criar_sistema("AD"), criar_sistema("E-mail"), criar_sistema("VPN")
    client_gestor.put("/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": [ad]})
    client_gestor.put(
        "/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": [vpn, email]}
    )
    linha = auditoria("acesso_pacote", "editar")[-1]
    assert linha["entidade_id"] is None
    assert linha["antes"] == {"cargo": "Vendedor", "sistemas": ["AD"]}
    assert linha["depois"] == {"cargo": "Vendedor", "sistemas": ["E-mail", "VPN"]}


def test_pacote_lista_sistema_desativado_depois(client_gestor):
    from app.db.acesso_db import AcessoDBManager

    ad = criar_sistema("AD")
    client_gestor.put("/api/acessos/pacotes", json={"cargo": "Vendedor", "sistema_ids": [ad]})
    client_gestor.put(f"/api/acessos/sistemas/{ad}", json={"ativo": False})
    pacotes = client_gestor.get("/api/acessos/pacotes").json()
    assert pacotes[0]["sistemas"][0]["nome"] == "AD"
    assert pacotes[0]["sistemas"][0]["ativo"] is False
    assert AcessoDBManager().sistemas_do_pacote("vendedor") == []
