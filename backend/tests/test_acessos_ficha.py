"""Ficha de acessos do colaborador e equipamentos em posse dele (/api/acessos/ficha)."""
import pytest

from app.schemas.acessos import MSG_SEM_SENHA
from tests.acessos_helpers import (
    CPF_JOAO,
    auditoria,
    contar,
    criar_acesso,
    criar_colaborador,
    criar_sistema,
    vincular_item,
)

pytestmark = pytest.mark.integration


def _corpo(c, s, login="maria.silva", item_id=None):
    return {"colaborador_id": c, "sistema_id": s, "login": login, "item_id": item_id}


def test_registrar_acesso_com_equipamento(client_tecnico, criar_item):
    c = criar_colaborador()
    ad = criar_sistema("AD")
    item = criar_item()
    r = client_tecnico.post("/api/acessos/ficha", json=_corpo(c, ad, item_id=item["id"]))
    assert r.status_code == 201
    corpo = r.json()
    assert corpo["status"] == "ativo"
    assert corpo["criado_por"] == "tecnico.teste"
    assert corpo["colaborador_nome"] == "Maria Silva"
    assert corpo["sistema_nome"] == "AD"
    assert corpo["item_id"] == item["id"]
    assert corpo["item_descricao"].startswith("Notebook Dell Latitude 5420 · SN-")
    assert len(auditoria("acesso", "criar")) == 1


def test_login_ativo_repetido_409(client_tecnico):
    c, ad = criar_colaborador(), criar_sistema("AD")
    client_tecnico.post("/api/acessos/ficha", json=_corpo(c, ad))
    r = client_tecnico.post("/api/acessos/ficha", json=_corpo(c, ad, "MARIA.silva"))
    assert r.status_code == 409
    assert r.json()["detail"] == "Este login já está ativo para o colaborador neste sistema."


def test_criar_acesso_reativa_login_desativado_ignorando_caixa_e_espacos(client_tecnico):
    c, ad = criar_colaborador(), criar_sistema("AD")
    acesso_id = criar_acesso(c, ad, "Maria.Silva", status="desativado")
    r = client_tecnico.post("/api/acessos/ficha", json=_corpo(c, ad, "  maria.silva "))
    assert r.status_code == 201
    corpo = r.json()
    assert corpo["id"] == acesso_id
    assert corpo["login"] == "maria.silva"
    assert corpo["status"] == "ativo"
    assert corpo["desativado_em"] is None
    assert corpo["desativado_por"] is None
    assert corpo["criado_por"] == "tecnico.teste"
    assert contar("colaborador_acessos") == 1
    assert len(auditoria("acesso", "reativar")) == 1


def test_sistema_desativado_422(client_tecnico):
    c, ad = criar_colaborador(), criar_sistema("AD", ativo=False)
    r = client_tecnico.post("/api/acessos/ficha", json=_corpo(c, ad))
    assert r.status_code == 422
    assert r.json()["detail"] == "Sistema desativado: não é possível registrar acesso nele."


def test_colaborador_inexistente_404(client_tecnico):
    ad = criar_sistema("AD")
    r = client_tecnico.post("/api/acessos/ficha", json=_corpo(999, ad))
    assert r.status_code == 404
    assert r.json()["detail"] == "Colaborador não encontrado."


def test_sistema_inexistente_404(client_tecnico):
    c = criar_colaborador()
    r = client_tecnico.post("/api/acessos/ficha", json=_corpo(c, 999))
    assert r.status_code == 404
    assert r.json()["detail"] == "Sistema não encontrado."


def test_equipamento_inexistente_404(client_tecnico):
    c, ad = criar_colaborador(), criar_sistema("AD")
    r = client_tecnico.post("/api/acessos/ficha", json=_corpo(c, ad, item_id=999))
    assert r.status_code == 404
    assert r.json()["detail"] == "Equipamento não encontrado."
    assert contar("colaborador_acessos") == 0


def test_senha_em_qualquer_grafia_422(client_tecnico):
    c, ad = criar_colaborador(), criar_sistema("AD")
    r = client_tecnico.post("/api/acessos/ficha", json={**_corpo(c, ad), "Senha": "123"})
    assert r.status_code == 422
    assert r.json()["detail"] == MSG_SEM_SENHA
    assert contar("colaborador_acessos") == 0

    acesso_id = criar_acesso(c, ad, "maria.silva")
    r = client_tecnico.put(f"/api/acessos/ficha/{acesso_id}", json={"senha_vpn": "x"})
    assert r.status_code == 422
    assert r.json()["detail"] == MSG_SEM_SENHA


def test_listar_ficha_exige_filtro_e_ordena(client_tecnico, criar_item):
    c = criar_colaborador()
    ad, vpn = criar_sistema("AD"), criar_sistema("VPN")
    item = criar_item()
    criar_acesso(c, vpn, "msilva")
    criar_acesso(c, ad, "maria.antiga", status="desativado")
    criar_acesso(c, ad, "maria.silva", item_id=item["id"])

    r = client_tecnico.get("/api/acessos/ficha")
    assert r.status_code == 422
    assert r.json()["detail"] == "Informe colaborador_id ou item_id."

    r = client_tecnico.get(f"/api/acessos/ficha?colaborador_id={c}")
    assert [(a["sistema_nome"], a["status"]) for a in r.json()] == [
        ("AD", "ativo"), ("VPN", "ativo"), ("AD", "desativado"),
    ]

    r = client_tecnico.get(f"/api/acessos/ficha?item_id={item['id']}")
    assert [a["login"] for a in r.json()] == ["maria.silva"]

    r = client_tecnico.get(f"/api/acessos/ficha?colaborador_id={c}&status=desativado")
    assert [a["login"] for a in r.json()] == ["maria.antiga"]


def test_editar_login_e_limpar_equipamento(client_tecnico, criar_item):
    c, ad = criar_colaborador(), criar_sistema("AD")
    item = criar_item()
    acesso_id = criar_acesso(c, ad, "maria.silva", item_id=item["id"])
    criar_acesso(c, ad, "outro.login")

    r = client_tecnico.put(f"/api/acessos/ficha/{acesso_id}", json={"item_id": None})
    assert r.status_code == 200
    assert r.json()["item_id"] is None
    assert r.json()["login"] == "maria.silva"

    r = client_tecnico.put(f"/api/acessos/ficha/{acesso_id}", json={"login": "Outro.Login"})
    assert r.status_code == 409
    assert r.json()["detail"] == (
        "Já existe um acesso com esse login para o colaborador neste sistema."
    )

    r = client_tecnico.put(f"/api/acessos/ficha/{acesso_id}", json={"login": " m.silva "})
    assert r.json()["login"] == "m.silva"
    assert len(auditoria("acesso", "editar")) == 2

    r = client_tecnico.put(f"/api/acessos/ficha/{acesso_id}", json={})
    assert r.status_code == 400
    assert r.json()["detail"] == "Nenhum campo para atualizar."

    r = client_tecnico.put("/api/acessos/ficha/999", json={"login": "x"})
    assert r.status_code == 404
    assert r.json()["detail"] == "Acesso não encontrado."


def test_desativar_acesso(client_tecnico):
    c, ad = criar_colaborador(), criar_sistema("AD")
    acesso_id = criar_acesso(c, ad, "maria.silva")
    r = client_tecnico.post(f"/api/acessos/ficha/{acesso_id}/desativar")
    assert r.status_code == 200
    corpo = r.json()
    assert corpo["status"] == "desativado"
    assert corpo["desativado_por"] == "tecnico.teste"
    assert corpo["desativado_em"] is not None

    r = client_tecnico.post(f"/api/acessos/ficha/{acesso_id}/desativar")
    assert r.status_code == 409
    assert r.json()["detail"] == "O acesso já está desativado."
    assert len(auditoria("acesso", "desativar")) == 1


def test_aprendiz_nao_ve_a_ficha(client_aprendiz):
    r = client_aprendiz.get("/api/acessos/ficha?colaborador_id=1")
    assert r.status_code == 403
    assert r.json()["detail"] == "Acesso restrito a Gestor ou Técnico."


def test_auditoria_da_ficha_nao_tem_senha(client_tecnico):
    c, ad = criar_colaborador(), criar_sistema("AD")
    client_tecnico.post("/api/acessos/ficha", json=_corpo(c, ad))
    for linha in auditoria("acesso"):
        for lado in (linha["antes"], linha["depois"]):
            assert not any("senha" in k.lower() or "pass" in k.lower() for k in (lado or {}))


def test_equipamentos_inclui_legado_por_cpf_com_mascara(client_tecnico, criar_item):
    c = criar_colaborador()
    joao = criar_colaborador("João Souza", cpf=CPF_JOAO)
    a, b, c_item, d, e = (criar_item() for _ in range(5))
    vincular_item(a["id"], colaborador_id=c)
    vincular_item(b["id"], cpf="111.444.777-35", status="Pendente Devolução")
    vincular_item(c_item["id"], colaborador_id=c, status="Disponível")
    vincular_item(d["id"], cpf=CPF_JOAO)
    vincular_item(e["id"], colaborador_id=c)
    from app.db.database_mysql import get_cursor

    with get_cursor() as cur:
        cur.execute("UPDATE items SET is_active = 0 WHERE id = %s", (e["id"],))

    r = client_tecnico.get(f"/api/acessos/colaboradores/{c}/equipamentos")
    assert r.status_code == 200
    assert [i["id"] for i in r.json()] == [a["id"], b["id"]]
    assert r.json()[1]["status"] == "Pendente Devolução"
    assert [i["id"] for i in client_tecnico.get(
        f"/api/acessos/colaboradores/{joao}/equipamentos").json()] == [d["id"]]

    r = client_tecnico.get("/api/acessos/colaboradores/999/equipamentos")
    assert r.status_code == 404
    assert r.json()["detail"] == "Colaborador não encontrado."


def test_editar_com_login_nulo_ou_em_branco_e_400(client_tecnico):
    c, ad = criar_colaborador(), criar_sistema("AD")
    acesso_id = criar_acesso(c, ad, "maria.silva")
    for corpo in ({"login": None}, {"login": "   "}):
        r = client_tecnico.put(f"/api/acessos/ficha/{acesso_id}", json=corpo)
        assert r.status_code == 400
        assert r.json()["detail"] == "Nenhum campo para atualizar."
    assert client_tecnico.get(f"/api/acessos/ficha?colaborador_id={c}").json()[0]["login"] == "maria.silva"
