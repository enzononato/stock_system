"""Empréstimo e devolução com o colaborador do cadastro e a solicitação de termo automática."""
from datetime import datetime

import pytest

from app.services.audit import AuditContext
from tests.fixtures_termos import (  # noqa: F401
    colaborador,
    corpo_emprestimo_colaborador,
    criar_colaborador,
    revsign_fake,
    unidade_juazeiro,
)

pytestmark = pytest.mark.integration

CPF_VALIDO_TESTE = "11144477735"
CPF_ITEM_FIXTURE = "11122233344"  # CPF do responsável nas fixtures item_pendente/item_indisponivel
ARQUIVO_PDF = {"signed_pdf": ("termo.pdf", b"conteudo-fake-pdf", "application/pdf")}
CTX = AuditContext(1, "gestor.teste", "Gestor", "127.0.0.1")


def _corpo_antigo(item_id: int) -> dict:
    return {
        "item_id": item_id,
        "usuario": "Fulano de Tal",
        "cpf": CPF_VALIDO_TESTE,
        "center_cost": "101 - Puxada",
        "cargo": "Analista",
        "setor": "TI",
        "revenda": "Revalle Juazeiro",
        "date_issue": datetime.now().strftime("%d/%m/%Y"),
    }


def _item(client, item_id):
    return client.get(f"/api/items/{item_id}").json()


def _ultimo_history(inv_manager, item_id, operacao):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT * FROM history WHERE item_id=%s AND operation=%s ORDER BY id DESC LIMIT 1",
            (item_id, operacao),
        )
        return cur.fetchone()


def test_emprestimo_com_colaborador_preenche_dados_do_cadastro(
    client_gestor, inv_manager, item_disponivel, colaborador
):
    from app.services.termos import MOTIVO_DESLIGADO

    item_id = item_disponivel["id"]
    corpo = corpo_emprestimo_colaborador(
        item_id, colaborador["id"], cargo="Outro", setor="TI", usuario="Ignorado"
    )
    resp = client_gestor.post("/api/loans", json=corpo)
    assert resp.status_code == 200, resp.text
    item = _item(client_gestor, item_id)
    assert item["assigned_to"] == "Maria Souza"
    assert item["cpf"] == "52998224725"
    assert item["colaborador_id"] == colaborador["id"]
    historico = _ultimo_history(inv_manager, item_id, "Empréstimo")
    assert historico["cargo"] == "Analista Financeiro"
    assert historico["setor"] == "Financeiro"
    termo = resp.json()["termo"]
    assert termo["status"] == "manual"
    assert termo["motivo"] == MOTIVO_DESLIGADO


def test_colaborador_sem_email_400(client_gestor, item_disponivel, criar_colaborador):
    sem_email = criar_colaborador(email=None)
    resp = client_gestor.post("/api/loans", json=corpo_emprestimo_colaborador(item_disponivel["id"], sem_email["id"]))
    assert resp.status_code == 400
    assert resp.json()["detail"] == "O colaborador precisa ter e-mail cadastrado para receber o termo."
    assert _item(client_gestor, item_disponivel["id"])["status"] == "Disponível"


def test_colaborador_desligado_400(client_gestor, item_disponivel, criar_colaborador):
    desligado = criar_colaborador(situacao="desligado")
    resp = client_gestor.post("/api/loans", json=corpo_emprestimo_colaborador(item_disponivel["id"], desligado["id"]))
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Colaborador desligado não pode receber equipamento."
    assert _item(client_gestor, item_disponivel["id"])["status"] == "Disponível"


def test_colaborador_inexistente_400(client_gestor, item_disponivel):
    resp = client_gestor.post("/api/loans", json=corpo_emprestimo_colaborador(item_disponivel["id"], 9999))
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Colaborador não encontrado."
    assert _item(client_gestor, item_disponivel["id"])["status"] == "Disponível"


def test_payload_antigo_continua_aceito(client_gestor, item_disponivel):
    item_id = item_disponivel["id"]
    resp = client_gestor.post("/api/loans", json=_corpo_antigo(item_id))
    assert resp.status_code == 200, resp.text
    assert resp.json()["detail"] == f"Empréstimo do item {item_id} para Fulano de Tal iniciado. Status: Pendente."
    termo = resp.json()["termo"]
    assert termo["tipo"] == "entrega"
    assert termo["colaborador_id"] is None
    assert _item(client_gestor, item_id)["colaborador_id"] is None


def test_sem_colaborador_e_sem_campos_antigos_422(client_gestor, item_disponivel):
    corpo = {
        "item_id": item_disponivel["id"],
        "center_cost": "101 - Puxada",
        "revenda": "Revalle Juazeiro",
        "date_issue": datetime.now().strftime("%d/%m/%Y"),
    }
    resp = client_gestor.post("/api/loans", json=corpo)
    assert resp.status_code == 422
    assert resp.json()["detail"] == "Informe o colaborador do cadastro ou nome, CPF, cargo e setor."


def test_emprestimo_com_revsign_envia_para_email_pessoal(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    resp = client_gestor.post(
        "/api/loans", json=corpo_emprestimo_colaborador(item_disponivel["id"], colaborador["id"])
    )
    assert resp.status_code == 200, resp.text
    assert resp.json()["termo"]["status"] == "enviado"
    assert revsign_fake.criados[0].signatario_email == "maria.souza@gmail.com"


def test_falha_ao_solicitar_termo_nao_derruba_o_emprestimo(
    client_gestor, item_disponivel, colaborador, monkeypatch
):
    from app.services import termos

    def quebra(*a, **k):
        raise RuntimeError("falha inesperada")

    monkeypatch.setattr(termos, "solicitar", quebra)
    resp = client_gestor.post(
        "/api/loans", json=corpo_emprestimo_colaborador(item_disponivel["id"], colaborador["id"])
    )
    assert resp.status_code == 200
    assert resp.json()["termo"] is None
    assert _item(client_gestor, item_disponivel["id"])["status"] == "Pendente"


def test_devolucao_confirmada_limpa_colaborador_id(
    client_gestor, item_disponivel, colaborador, forcar_devolucao_iniciada
):
    item_id = item_disponivel["id"]
    assert client_gestor.post(
        "/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador["id"])
    ).status_code == 200
    assert client_gestor.post(f"/api/loans/{item_id}/confirm", files=ARQUIVO_PDF).status_code == 200
    forcar_devolucao_iniciada(item_id)
    assert _item(client_gestor, item_id)["colaborador_id"] == colaborador["id"]

    resp = client_gestor.post(f"/api/loans/{item_id}/return/confirm", files=ARQUIVO_PDF)
    assert resp.status_code == 200
    assert _item(client_gestor, item_id)["colaborador_id"] is None


def test_devolucao_de_item_legado_vincula_colaborador_escolhido(
    client_gestor, item_indisponivel, criar_colaborador, termo_devolucao_disponivel
):
    if not termo_devolucao_disponivel:
        pytest.skip("Modelo do termo de devolução ausente neste ambiente.")
    colaborador = criar_colaborador(cpf=CPF_ITEM_FIXTURE)
    item_id = item_indisponivel["id"]
    resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate", json={"colaborador_id": colaborador["id"]})
    assert resp.status_code == 200, resp.text
    assert _item(client_gestor, item_id)["colaborador_id"] == colaborador["id"]
    termo = resp.json()["termo"]
    assert termo["tipo"] == "devolucao"
    assert termo["tem_documento"] is True
    assert resp.json()["detail"] == "Termo de devolução gerado."
    assert resp.json()["download_url"]


def test_devolucao_sem_corpo_continua_funcionando(client_gestor, item_indisponivel, termo_devolucao_disponivel):
    if not termo_devolucao_disponivel:
        pytest.skip("Modelo do termo de devolução ausente neste ambiente.")
    resp = client_gestor.post(f"/api/loans/{item_indisponivel['id']}/return/initiate")
    assert resp.status_code == 200, resp.text
    assert resp.json()["termo"]["tipo"] == "devolucao"


def test_devolucao_com_colaborador_diferente_do_vinculado_400(
    client_gestor, item_disponivel, colaborador, criar_colaborador, forcar_devolucao_iniciada
):
    item_id = item_disponivel["id"]
    assert client_gestor.post(
        "/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador["id"])
    ).status_code == 200
    assert client_gestor.post(f"/api/loans/{item_id}/confirm", files=ARQUIVO_PDF).status_code == 200
    outro = criar_colaborador(cpf="11144477735", nome="Outro Colaborador", email="outro@gmail.com")
    resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate", json={"colaborador_id": outro["id"]})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "O equipamento está vinculado a outro colaborador."
    assert _item(client_gestor, item_id)["status"] == "Indisponível"


def test_devolucao_com_colaborador_inexistente_ou_sem_email_400(
    client_gestor, item_indisponivel, criar_colaborador
):
    item_id = item_indisponivel["id"]
    resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate", json={"colaborador_id": 9999})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Colaborador não encontrado."
    sem_email = criar_colaborador(email=None)
    resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate", json={"colaborador_id": sem_email["id"]})
    assert resp.status_code == 400
    assert resp.json()["detail"] == "O colaborador precisa ter e-mail cadastrado para receber o termo."
    assert _item(client_gestor, item_id)["status"] == "Indisponível"


def test_devolucao_aceita_colaborador_desligado(
    client_gestor, item_indisponivel, criar_colaborador, termo_devolucao_disponivel
):
    if not termo_devolucao_disponivel:
        pytest.skip("Modelo do termo de devolução ausente neste ambiente.")
    desligado = criar_colaborador(situacao="desligado", cpf=CPF_ITEM_FIXTURE)
    resp = client_gestor.post(
        f"/api/loans/{item_indisponivel['id']}/return/initiate", json={"colaborador_id": desligado["id"]}
    )
    assert resp.status_code == 200, resp.text


def test_devolucao_cancela_uso_pendente(
    client_gestor, inv_manager, item_indisponivel, colaborador, revsign_fake, unidade_juazeiro,
    termo_devolucao_disponivel,
):
    """Troca de usuário/aparelho: a devolução cancela o termo de uso que ainda estava pendente."""
    if not termo_devolucao_disponivel:
        pytest.skip("Modelo do termo de devolução ausente neste ambiente.")
    from app.services import termos

    item_id = item_indisponivel["id"]
    inv_manager.vincular_colaborador(item_id, colaborador["id"])
    uso = termos.solicitar(item_id, "uso", CTX)
    assert uso["status"] == "enviado" and uso["revsign_documento_id"] == "1001"

    resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate")
    assert resp.status_code == 200, resp.text
    uso = termos.get_termo_db().obter(uso["id"])
    assert uso["status"] == "cancelado"
    assert "Substituída pela solicitação #" in uso["motivo"]
    assert "1001" in revsign_fake.cancelados
    assert resp.json()["termo"]["status"] == "enviado"
    assert resp.json()["termo"]["tipo"] == "devolucao"


def test_estorno_do_emprestimo_cancela_solicitacao_e_no_revsign(
    client_gestor, inv_manager, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    from app.services.termos import MOTIVO_ESTORNO

    item_id = item_disponivel["id"]
    resp = client_gestor.post("/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador["id"]))
    termo_id = resp.json()["termo"]["id"]
    history_id = inv_manager.ultimo_history_id(item_id, "Empréstimo")

    resp = client_gestor.post(f"/api/history/{history_id}/reverse", json={"password": "SenhaForte#123"})
    assert resp.status_code == 200, resp.text
    from app.dependencies import get_termo_db

    termo = get_termo_db().obter(termo_id)
    assert termo["status"] == "cancelado"
    assert termo["motivo"] == MOTIVO_ESTORNO
    assert revsign_fake.cancelados == ["1001"]
    item = _item(client_gestor, item_id)
    assert item["status"] == "Disponível"
    assert item["colaborador_id"] is None


def test_estorno_nao_falha_se_cancelar_solicitacoes_quebrar(
    client_gestor, inv_manager, item_disponivel, monkeypatch
):
    from app.services import termos

    item_id = item_disponivel["id"]
    assert client_gestor.post("/api/loans", json=_corpo_antigo(item_id)).status_code == 200
    history_id = inv_manager.ultimo_history_id(item_id, "Empréstimo")

    def quebra(*a, **k):
        raise RuntimeError("falha inesperada")

    monkeypatch.setattr(termos, "cancelar_por_estorno", quebra)
    resp = client_gestor.post(f"/api/history/{history_id}/reverse", json={"password": "SenhaForte#123"})
    assert resp.status_code == 200


def test_devolucao_com_cpf_diferente_do_responsavel_422(
    client_gestor, item_indisponivel, colaborador, termo_devolucao_disponivel
):
    """O termo sai com o nome/CPF do equipamento: o colaborador escolhido tem de ser essa pessoa."""
    item_id = item_indisponivel["id"]
    resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate", json={"colaborador_id": colaborador["id"]})
    assert resp.status_code == 422
    assert resp.json()["detail"] == "O CPF do colaborador escolhido não confere com o do responsável pelo equipamento."
    item = _item(client_gestor, item_id)
    assert item["status"] == "Indisponível" and item["colaborador_id"] is None


def test_devolucao_sem_cpf_no_item_aceita_qualquer_colaborador(
    client_gestor, item_indisponivel, colaborador, executar_sql, termo_devolucao_disponivel
):
    if not termo_devolucao_disponivel:
        pytest.skip("Modelo do termo de devolução ausente neste ambiente.")
    executar_sql("UPDATE items SET cpf=NULL WHERE id=%s", (item_indisponivel["id"],))
    resp = client_gestor.post(
        f"/api/loans/{item_indisponivel['id']}/return/initiate", json={"colaborador_id": colaborador["id"]}
    )
    assert resp.status_code == 200, resp.text


def test_estorno_da_devolucao_restaura_colaborador_id_pelo_cpf(
    client_gestor, inv_manager, item_disponivel, colaborador, forcar_devolucao_iniciada
):
    item_id = item_disponivel["id"]
    assert client_gestor.post(
        "/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador["id"])
    ).status_code == 200
    assert client_gestor.post(f"/api/loans/{item_id}/confirm", files=ARQUIVO_PDF).status_code == 200
    forcar_devolucao_iniciada(item_id)
    assert client_gestor.post(f"/api/loans/{item_id}/return/confirm", files=ARQUIVO_PDF).status_code == 200
    assert _item(client_gestor, item_id)["colaborador_id"] is None

    devolucao_id = inv_manager.ultimo_history_id(item_id, "Devolução")
    resp = client_gestor.post(f"/api/history/{devolucao_id}/reverse", json={"password": "SenhaForte#123"})
    assert resp.status_code == 200, resp.text
    item = _item(client_gestor, item_id)
    assert item["status"] == "Indisponível"
    assert item["colaborador_id"] == colaborador["id"]


def test_estorno_da_confirmacao_reabre_a_solicitacao_assinada(
    client_gestor, inv_manager, item_disponivel, colaborador, unidade_juazeiro
):
    from app.dependencies import get_termo_db
    from app.services.termos import MOTIVO_CONFIRMACAO_ESTORNADA

    item_id = item_disponivel["id"]
    termo_id = client_gestor.post(
        "/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador["id"])
    ).json()["termo"]["id"]
    assert client_gestor.post(f"/api/loans/{item_id}/confirm", files=ARQUIVO_PDF).status_code == 200
    assert get_termo_db().obter(termo_id)["status"] == "assinado"

    confirmacao_id = inv_manager.ultimo_history_id(item_id, "Confirmação Empréstimo")
    resp = client_gestor.post(f"/api/history/{confirmacao_id}/reverse", json={"password": "SenhaForte#123"})
    assert resp.status_code == 200, resp.text

    termo = get_termo_db().obter(termo_id)
    assert termo["status"] == "manual"
    assert termo["motivo"] == MOTIVO_CONFIRMACAO_ESTORNADA
    assert termo["assinado_em"] is None and termo["arquivo_assinado_key"] is None
    assert _item(client_gestor, item_id)["status"] == "Pendente"
    # o card volta a oferecer "Anexar assinado": o anexo conclui de novo
    assert client_gestor.post(f"/api/termos/{termo_id}/assinado", files=ARQUIVO_PDF).status_code == 200
    assert _item(client_gestor, item_id)["status"] == "Indisponível"
    assert get_termo_db().obter(termo_id)["status"] == "assinado"


def test_devolucao_de_item_vinculado_nao_confere_cpf(
    client_gestor, item_disponivel, colaborador, executar_sql, termo_devolucao_disponivel
):
    """A regra do CPF vale só para item legado (sem vínculo): o vinculado já está ligado ao colaborador."""
    if not termo_devolucao_disponivel:
        pytest.skip("Modelo do termo de devolução ausente neste ambiente.")
    item_id = item_disponivel["id"]
    assert client_gestor.post(
        "/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador["id"])
    ).status_code == 200
    assert client_gestor.post(f"/api/loans/{item_id}/confirm", files=ARQUIVO_PDF).status_code == 200
    executar_sql("UPDATE items SET cpf=%s WHERE id=%s", ("11144477735", item_id))

    resp = client_gestor.post(f"/api/loans/{item_id}/return/initiate", json={"colaborador_id": colaborador["id"]})
    assert resp.status_code == 200, resp.text


def test_estorno_da_devolucao_normaliza_o_cpf_do_historico(
    client_gestor, inv_manager, item_disponivel, colaborador, executar_sql, forcar_devolucao_iniciada
):
    item_id = item_disponivel["id"]
    assert client_gestor.post(
        "/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador["id"])
    ).status_code == 200
    assert client_gestor.post(f"/api/loans/{item_id}/confirm", files=ARQUIVO_PDF).status_code == 200
    forcar_devolucao_iniciada(item_id)
    assert client_gestor.post(f"/api/loans/{item_id}/return/confirm", files=ARQUIVO_PDF).status_code == 200
    # CPF gravado com máscara no histórico do empréstimo
    executar_sql(
        "UPDATE history SET cpf=%s WHERE item_id=%s AND operation='Empréstimo'", ("529.982.247-25", item_id)
    )

    devolucao_id = inv_manager.ultimo_history_id(item_id, "Devolução")
    resp = client_gestor.post(f"/api/history/{devolucao_id}/reverse", json={"password": "SenhaForte#123"})
    assert resp.status_code == 200, resp.text
    assert _item(client_gestor, item_id)["colaborador_id"] == colaborador["id"]
