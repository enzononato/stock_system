"""Resolução de tarefas, conclusão e cancelamento de eventos de acesso."""
import pytest

from tests.acessos_helpers import (
    auditoria,
    contar,
    criar_acesso,
    criar_colaborador,
    criar_sistema,
    definir_pacote,
    vincular_item,
)

pytestmark = pytest.mark.integration


def _abrir(cliente, colaborador_id, tipo, **extra):
    r = cliente.post(
        "/api/acessos/eventos", json={"colaborador_id": colaborador_id, "tipo": tipo, **extra}
    )
    assert r.status_code == 201, r.text
    return r.json()


def _tarefa(evento, sistema_nome):
    return next(t for t in evento["tarefas"] if t["sistema_nome"] == sistema_nome)


def _resolver(cliente, tarefa_id, **corpo):
    return cliente.post(f"/api/acessos/tarefas/{tarefa_id}/resolver", json=corpo)


@pytest.fixture
def admissao(client_tecnico):
    ad, email = criar_sistema("AD"), criar_sistema("E-mail")
    definir_pacote("Vendedor", ad, email)
    c = criar_colaborador()
    return c, ad, email, _abrir(client_tecnico, c, "admissao")


def test_criar_feito_registra_login_e_cria_acesso(client_tecnico, admissao, criar_item):
    c, ad, _, ev = admissao
    item = criar_item()
    r = _resolver(
        client_tecnico, _tarefa(ev, "AD")["id"],
        status="feito", login="  maria.silva ", item_id=item["id"],
    )
    assert r.status_code == 200
    novo = r.json()
    tarefa = _tarefa(novo, "AD")
    assert tarefa["status"] == "feito"
    assert tarefa["login"] == "maria.silva"
    assert tarefa["feito_por"] == "tecnico.teste"
    assert tarefa["acesso_id"] is not None
    assert novo["status"] == "aberto"

    ficha = client_tecnico.get(f"/api/acessos/ficha?colaborador_id={c}").json()
    assert [(a["sistema_nome"], a["login"], a["status"], a["item_id"]) for a in ficha] == [
        ("AD", "maria.silva", "ativo", item["id"])
    ]
    assert len(auditoria("acesso", "criar")) == 1
    assert len(auditoria("acesso_evento", "resolver_tarefa")) == 1
    assert auditoria("acesso_evento", "resolver_tarefa")[0]["depois"]["login"] == "maria.silva"


def test_ultima_tarefa_conclui_a_admissao(client_tecnico, admissao):
    _, _, _, ev = admissao
    _resolver(client_tecnico, _tarefa(ev, "AD")["id"], status="feito", login="maria.silva")
    r = _resolver(
        client_tecnico, _tarefa(ev, "E-mail")["id"],
        status="nao_se_aplica", observacao="Sem licença",
    )
    assert r.status_code == 200
    final = r.json()
    assert final["status"] == "concluido"
    assert final["concluido_por"] == "tecnico.teste"
    assert final["pendencias"] == []
    assert _tarefa(final, "E-mail")["observacao"] == "Sem licença"
    assert len(auditoria("acesso_evento", "concluir")) == 1


def test_criar_feito_sem_login_422(client_tecnico, admissao):
    _, _, _, ev = admissao
    tarefa_id = _tarefa(ev, "AD")["id"]
    r = _resolver(client_tecnico, tarefa_id, status="feito")
    assert r.status_code == 422
    assert r.json()["detail"] == "Informe o login criado no sistema."
    evento = client_tecnico.get(f"/api/acessos/eventos/{ev['id']}").json()
    assert _tarefa(evento, "AD")["status"] == "pendente"


def test_criar_feito_reativa_login_desativado(client_tecnico, admissao):
    c, ad, _, ev = admissao
    acesso_id = criar_acesso(c, ad, "maria.silva", status="desativado")
    r = _resolver(client_tecnico, _tarefa(ev, "AD")["id"], status="feito", login="Maria.Silva")
    assert r.status_code == 200
    assert _tarefa(r.json(), "AD")["acesso_id"] == acesso_id
    ficha = client_tecnico.get(f"/api/acessos/ficha?colaborador_id={c}").json()
    assert [(a["id"], a["status"]) for a in ficha] == [(acesso_id, "ativo")]
    assert contar("colaborador_acessos") == 1


def test_criar_feito_com_login_ja_ativo_so_vincula(client_tecnico, admissao):
    c, ad, _, ev = admissao
    acesso_id = criar_acesso(c, ad, "maria.silva")  # registrado à mão depois de abrir o evento
    r = _resolver(client_tecnico, _tarefa(ev, "AD")["id"], status="feito", login="MARIA.silva")
    assert r.status_code == 200
    assert _tarefa(r.json(), "AD")["acesso_id"] == acesso_id
    assert contar("colaborador_acessos") == 1


def test_criar_feito_em_sistema_desativado_depois_422(client_tecnico, admissao, client_gestor):
    _, ad, _, ev = admissao
    client_gestor.put(f"/api/acessos/sistemas/{ad}", json={"ativo": False})
    tarefa_id = _tarefa(ev, "AD")["id"]
    r = _resolver(client_tecnico, tarefa_id, status="feito", login="maria.silva")
    assert r.status_code == 422
    assert r.json()["detail"] == "Sistema desativado: não é possível registrar acesso nele."
    r = _resolver(client_tecnico, tarefa_id, status="nao_se_aplica", observacao="Sistema saiu")
    assert r.status_code == 200
    r = _resolver(
        client_tecnico, _tarefa(ev, "E-mail")["id"], status="feito", login="maria@revalle"
    )
    assert r.json()["status"] == "concluido"


def test_nao_se_aplica_exige_observacao(client_tecnico, admissao):
    _, _, _, ev = admissao
    r = _resolver(client_tecnico, _tarefa(ev, "AD")["id"], status="nao_se_aplica")
    assert r.status_code == 422
    assert r.json()["detail"] == "Informe a observação para marcar como não se aplica."


def test_resolver_duas_vezes_409_e_um_so_acesso(client_tecnico, admissao):
    c, _, _, ev = admissao
    tarefa_id = _tarefa(ev, "AD")["id"]
    assert _resolver(client_tecnico, tarefa_id, status="feito", login="maria.silva").status_code == 200
    r = _resolver(client_tecnico, tarefa_id, status="feito", login="maria.silva")
    assert r.status_code == 409
    assert r.json()["detail"] == "Esta tarefa já foi resolvida."
    assert contar("colaborador_acessos") == 1


def test_resolver_ao_mesmo_tempo_cria_um_so_acesso(_fastapi_app, client_tecnico, admissao):
    """Dois técnicos (ou dois cliques) resolvendo a mesma tarefa em paralelo."""
    from concurrent.futures import ThreadPoolExecutor

    _, _, _, ev = admissao
    tarefa_id = _tarefa(ev, "AD")["id"]
    cabecalho = client_tecnico.headers

    def chamar(_):
        from fastapi.testclient import TestClient

        outro = TestClient(_fastapi_app)
        return outro.post(
            f"/api/acessos/tarefas/{tarefa_id}/resolver",
            json={"status": "feito", "login": "maria.silva"},
            headers={"Authorization": cabecalho["Authorization"]},
        ).status_code

    with ThreadPoolExecutor(2) as pool:
        codigos = sorted(pool.map(chamar, range(2)))
    assert codigos == [200, 409]
    assert contar("colaborador_acessos") == 1


def test_desativar_feito_desativa_o_acesso(client_tecnico):
    ad = criar_sistema("AD")
    c = criar_colaborador()
    acesso_id = criar_acesso(c, ad, "maria.silva")
    ev = _abrir(client_tecnico, c, "desligamento")
    r = _resolver(client_tecnico, _tarefa(ev, "AD")["id"], status="feito")
    assert r.status_code == 200
    assert r.json()["status"] == "concluido"
    ficha = client_tecnico.get(f"/api/acessos/ficha?colaborador_id={c}").json()
    assert ficha[0]["id"] == acesso_id
    assert ficha[0]["status"] == "desativado"
    assert ficha[0]["desativado_por"] == "tecnico.teste"


def test_desativar_acesso_ja_desativado_so_marca_a_tarefa(client_tecnico):
    ad = criar_sistema("AD")
    c = criar_colaborador()
    acesso_id = criar_acesso(c, ad, "maria.silva")
    ev = _abrir(client_tecnico, c, "desligamento")
    client_tecnico.post(f"/api/acessos/ficha/{acesso_id}/desativar")
    r = _resolver(client_tecnico, _tarefa(ev, "AD")["id"], status="feito")
    assert r.status_code == 200
    assert _tarefa(r.json(), "AD")["status"] == "feito"
    assert len(auditoria("acesso", "desativar")) == 1


def test_desligamento_nao_conclui_com_equipamento(client_tecnico, criar_item, executar_sql):
    ad = criar_sistema("AD")
    c = criar_colaborador()
    criar_acesso(c, ad, "maria.silva")
    item = criar_item()
    vincular_item(item["id"], colaborador_id=c)
    ev = _abrir(client_tecnico, c, "desligamento")

    r = _resolver(client_tecnico, _tarefa(ev, "AD")["id"], status="feito")
    assert r.status_code == 200
    final = r.json()
    assert final["status"] == "aberto"
    assert final["pendencias"] == ["1 equipamento ainda com o colaborador"]
    assert final["pode_concluir"] is False

    r = client_tecnico.post(f"/api/acessos/eventos/{ev['id']}/concluir")
    assert r.status_code == 409
    assert r.json()["detail"] == "Não é possível concluir: 1 equipamento ainda com o colaborador."

    executar_sql(
        "UPDATE items SET status='Disponível', colaborador_id=NULL, cpf=NULL WHERE id=%s",
        (item["id"],),
    )
    r = client_tecnico.post(f"/api/acessos/eventos/{ev['id']}/concluir")
    assert r.status_code == 200
    assert r.json()["status"] == "concluido"
    r = client_tecnico.post(f"/api/acessos/eventos/{ev['id']}/concluir")
    assert r.status_code == 409
    assert r.json()["detail"] == "O evento não está aberto."


def _mudanca(client_tecnico):
    ad, promax, vpn = criar_sistema("AD"), criar_sistema("Promax"), criar_sistema("VPN")
    definir_pacote("Vendedor", ad, promax)
    definir_pacote("Supervisor", ad, vpn)
    c = criar_colaborador(cargo="Supervisor")
    promax_acesso = criar_acesso(c, promax, "MSILVA")
    criar_acesso(c, ad, "maria.silva")
    ev = _abrir(client_tecnico, c, "mudanca_funcao", cargo_anterior="Vendedor")
    return c, promax_acesso, ev


def test_manter_vira_nao_se_aplica_com_observacao(client_tecnico):
    c, promax_acesso, ev = _mudanca(client_tecnico)
    r = _resolver(
        client_tecnico, _tarefa(ev, "Promax")["id"],
        status="manter", observacao="continua atendendo a rota antiga",
    )
    assert r.status_code == 200
    tarefa = _tarefa(r.json(), "Promax")
    assert tarefa["status"] == "nao_se_aplica"
    assert tarefa["observacao"] == "Acesso mantido: continua atendendo a rota antiga"
    ficha = client_tecnico.get(f"/api/acessos/ficha?colaborador_id={c}").json()
    assert next(a for a in ficha if a["id"] == promax_acesso)["status"] == "ativo"


def test_manter_fora_de_desativar_em_mudanca_422(client_tecnico):
    _, _, ev = _mudanca(client_tecnico)
    r = _resolver(
        client_tecnico, _tarefa(ev, "VPN")["id"], status="manter", observacao="porque sim"
    )
    assert r.status_code == 422
    assert r.json()["detail"] == "Manter só vale para tarefas de desativar em mudança de função."
    r = _resolver(client_tecnico, _tarefa(ev, "Promax")["id"], status="manter")
    assert r.status_code == 422
    assert r.json()["detail"] == "Informe o motivo para manter o acesso."


def test_manter_em_desligamento_422(client_tecnico):
    ad = criar_sistema("AD")
    c = criar_colaborador()
    criar_acesso(c, ad, "maria.silva")
    ev = _abrir(client_tecnico, c, "desligamento")
    r = _resolver(client_tecnico, _tarefa(ev, "AD")["id"], status="manter", observacao="x")
    assert r.status_code == 422
    assert r.json()["detail"] == "Manter só vale para tarefas de desativar em mudança de função."


def test_cancelar_evento(client_tecnico, admissao):
    c, _, _, ev = admissao
    url = f"/api/acessos/eventos/{ev['id']}/cancelar"
    r = client_tecnico.post(url, json={"motivo_cancelamento": "  "})
    assert r.status_code == 422
    assert r.json()["detail"] == "Motivo do cancelamento: Informe o motivo do cancelamento."

    r = client_tecnico.post(url, json={"motivo_cancelamento": "Aberto por engano"})
    assert r.status_code == 200
    corpo = r.json()
    assert corpo["status"] == "cancelado"
    assert corpo["cancelado_por"] == "tecnico.teste"
    assert corpo["motivo_cancelamento"] == "Aberto por engano"
    assert auditoria("acesso_evento", "cancelar")[0]["depois"] == {
        "motivo_cancelamento": "Aberto por engano"
    }

    r = _resolver(client_tecnico, _tarefa(ev, "AD")["id"], status="feito", login="maria.silva")
    assert r.status_code == 409
    assert r.json()["detail"] == "O evento não está aberto."
    assert client_tecnico.post(url, json={"motivo_cancelamento": "de novo"}).status_code == 409
    assert contar("colaborador_acessos") == 0

    r = client_tecnico.post(
        "/api/acessos/eventos", json={"colaborador_id": c, "tipo": "admissao"}
    )
    assert r.status_code == 201


def test_senha_no_corpo_422(client_tecnico, admissao):
    _, _, _, ev = admissao
    r = _resolver(
        client_tecnico, _tarefa(ev, "AD")["id"], status="feito", login="maria", password="x"
    )
    assert r.status_code == 422
    assert "não guarda senhas" in r.json()["detail"]
    assert contar("colaborador_acessos") == 0


def test_tarefa_inexistente_404_e_aprendiz_403(client_tecnico, client_aprendiz):
    r = _resolver(client_tecnico, 999, status="feito", login="x")
    assert r.status_code == 404
    assert r.json()["detail"] == "Tarefa não encontrada."
    r = client_aprendiz.post("/api/acessos/tarefas/1/resolver", json={"status": "feito"})
    assert r.status_code == 403
    r = client_aprendiz.post("/api/acessos/eventos/1/concluir")
    assert r.status_code == 403


def test_resolver_tarefas_diferentes_ao_mesmo_tempo_conclui_o_evento(_fastapi_app, client_tecnico):
    """As duas últimas tarefas do evento resolvidas em paralelo: nenhuma das duas pode ler o
    estado antigo, então o evento termina `concluido` (e nunca fica aberto sem pendência)."""
    from concurrent.futures import ThreadPoolExecutor
    from threading import Barrier

    from fastapi.testclient import TestClient

    ad, email = criar_sistema("AD"), criar_sistema("E-mail")
    definir_pacote("Vendedor", ad, email)
    cabecalho = {"Authorization": client_tecnico.headers["Authorization"]}

    for rodada in range(4):
        c = criar_colaborador(f"Pessoa {rodada}", cpf=f"9000000000{rodada}")
        ev = _abrir(client_tecnico, c, "admissao")
        ids = [_tarefa(ev, "AD")["id"], _tarefa(ev, "E-mail")["id"]]
        largada = Barrier(2, timeout=10)

        def chamar(i, ids=ids, largada=largada):
            outro = TestClient(_fastapi_app)
            largada.wait()
            return outro.post(
                f"/api/acessos/tarefas/{ids[i]}/resolver",
                json={"status": "feito", "login": f"login{i}"},
                headers=cabecalho,
            ).status_code

        with ThreadPoolExecutor(2) as pool:
            assert list(pool.map(chamar, range(2))) == [200, 200]
        final = client_tecnico.get(f"/api/acessos/eventos/{ev['id']}").json()
        assert final["status"] == "concluido", f"rodada {rodada}: {final['pendencias']}"
        assert final["tarefas_pendentes"] == 0


def test_manter_com_observacao_longa_422(client_tecnico):
    _, _, ev = _mudanca(client_tecnico)
    tarefa_id = _tarefa(ev, "Promax")["id"]
    r = _resolver(client_tecnico, tarefa_id, status="manter", observacao="x" * 485)
    assert r.status_code == 422
    assert "máximo 484 caracteres" in r.json()["detail"]
    r = _resolver(client_tecnico, tarefa_id, status="manter", observacao="x" * 484)
    assert r.status_code == 200
    assert len(_tarefa(r.json(), "Promax")["observacao"]) == 500


def test_resolver_criar_e_registrar_acesso_ao_mesmo_tempo_nao_da_500(_fastapi_app, client_tecnico):
    """Resolver uma tarefa "criar" enquanto POST /ficha registra acesso do mesmo colaborador no
    mesmo sistema: as duas rotas travam na mesma ordem, então nunca há 500 (deadlock)."""
    from concurrent.futures import ThreadPoolExecutor
    from threading import Barrier

    from fastapi.testclient import TestClient

    ad, email = criar_sistema("AD"), criar_sistema("E-mail")
    definir_pacote("Vendedor", ad, email)
    cabecalho = {"Authorization": client_tecnico.headers["Authorization"]}

    for rodada in range(6):
        c = criar_colaborador(f"Pessoa {rodada}", cpf=f"8000000000{rodada}")
        ev = _abrir(client_tecnico, c, "admissao")
        tarefa_id = _tarefa(ev, "AD")["id"]
        largada = Barrier(2, timeout=10)
        login_ficha = "mesmo.login" if rodada % 2 else "outro.login"

        def resolver():
            largada.wait()
            return TestClient(_fastapi_app).post(
                f"/api/acessos/tarefas/{tarefa_id}/resolver",
                json={"status": "feito", "login": "mesmo.login"},
                headers=cabecalho,
            ).status_code

        def registrar():
            largada.wait()
            return TestClient(_fastapi_app).post(
                "/api/acessos/ficha",
                json={"colaborador_id": c, "sistema_id": ad, "login": login_ficha, "item_id": None},
                headers=cabecalho,
            ).status_code

        with ThreadPoolExecutor(2) as pool:
            futuros = [pool.submit(resolver), pool.submit(registrar)]
            codigos = [f.result() for f in futuros]
        assert all(codigo in (200, 201, 409) for codigo in codigos), f"rodada {rodada}: {codigos}"
        # Nunca dois acessos ativos com o mesmo login (qualquer caixa) no mesmo sistema.
        assert contar(
            "colaborador_acessos", "colaborador_id=%s AND LOWER(login)='mesmo.login'", (c,)
        ) <= 1
