"""F4 / Task 6: baixa com registro de apagamento de dados (LGPD)."""
import json
from datetime import date

import pytest

from app.services.baixa import (
    NAO_APLICA,
    Apagamento,
    motivo_bloqueio_baixa,
    validar_apagamento,
)
from tests.dados_baixa import APAGAMENTO_OK

HOJE = date(2026, 11, 12)


def _validar(**kw):
    base = dict(
        tipo="Notebook", motivo="Doação", dados_apagados=None, metodo=None,
        apagado_por=None, apagado_em=None, hoje=HOJE,
    )
    base.update(kw)
    return validar_apagamento(**base)


# ── Unidade ─────────────────────────────────────────────────────────────────

@pytest.mark.unit
def test_tipo_sem_dados_vira_nao_aplica():
    assert _validar(tipo="Impressora", dados_apagados="sim", metodo="xx") == NAO_APLICA
    assert _validar(tipo="Impressora") == NAO_APLICA


@pytest.mark.unit
def test_notebook_sem_resposta_recusa():
    with pytest.raises(ValueError) as exc:
        _validar()
    assert str(exc.value) == "Informe se os dados do equipamento foram apagados."


@pytest.mark.unit
def test_dados_apagados_invalido_recusa():
    with pytest.raises(ValueError) as exc:
        _validar(dados_apagados="talvez")
    assert str(exc.value) == "Valor inválido para dados apagados."


@pytest.mark.unit
def test_nao_aplica_para_doacao_recusado():
    with pytest.raises(ValueError) as exc:
        _validar(dados_apagados="nao_aplica", motivo="Doação")
    assert str(exc.value) == (
        'Para Notebook, a baixa exige o apagamento dos dados (LGPD). '
        '"Não se aplica" só vale para Roubo ou Perda.'
    )


@pytest.mark.unit
@pytest.mark.parametrize("motivo", ["Roubo", "Perda"])
def test_roubo_ou_perda_aceita_nao_aplica(motivo):
    assert _validar(dados_apagados="nao_aplica", motivo=motivo) == NAO_APLICA


@pytest.mark.unit
def test_celular_com_reset_valido():
    r = _validar(
        tipo="Celular", motivo="Venda", dados_apagados="sim", metodo="reset_fabrica",
        apagado_por="Ana", apagado_em="01/11/2026",
    )
    assert r == Apagamento("sim", "reset_fabrica", "Ana", date(2026, 11, 1))


@pytest.mark.unit
def test_data_futura_recusa():
    with pytest.raises(ValueError) as exc:
        _validar(dados_apagados="sim", metodo="reset_fabrica", apagado_por="Ana", apagado_em="13/11/2026")
    assert str(exc.value) == "A data do apagamento não pode ser no futuro."


@pytest.mark.unit
def test_data_invalida_ou_ausente_recusa():
    for data in (None, "", "31/02/2026", "2026-11-01"):
        with pytest.raises(ValueError) as exc:
            _validar(dados_apagados="sim", metodo="reset_fabrica", apagado_por="Ana", apagado_em=data)
        assert str(exc.value) == "Informe a data do apagamento no formato dd/mm/aaaa."


@pytest.mark.unit
def test_metodo_invalido_recusa():
    for metodo in (None, "", "nao_aplica", "martelo"):
        with pytest.raises(ValueError) as exc:
            _validar(dados_apagados="sim", metodo=metodo, apagado_por="Ana", apagado_em="01/11/2026")
        assert str(exc.value) == "Informe o método de apagamento dos dados."


@pytest.mark.unit
def test_quem_apagou_obrigatorio_e_limitado():
    with pytest.raises(ValueError) as exc:
        _validar(dados_apagados="sim", metodo="reset_fabrica", apagado_por="  ", apagado_em="01/11/2026")
    assert str(exc.value) == "Informe quem apagou os dados."
    with pytest.raises(ValueError) as exc:
        _validar(dados_apagados="sim", metodo="reset_fabrica", apagado_por="x" * 101, apagado_em="01/11/2026")
    assert str(exc.value) == "Quem apagou: máximo de 100 caracteres."


@pytest.mark.unit
def test_motivo_bloqueio_baixa():
    assert motivo_bloqueio_baixa({"status": "Disponível"}, False) is None
    assert motivo_bloqueio_baixa({"status": "Em Manutenção"}, True) == (
        "Conclua ou cancele a manutenção em andamento antes da baixa."
    )
    assert motivo_bloqueio_baixa({"status": "Em Manutenção"}, False) is None
    assert motivo_bloqueio_baixa({"status": "Pendente"}, False) == "Não é possível remover produto emprestado."
    assert motivo_bloqueio_baixa({"status": "Indisponível"}, False) == (
        "Não é possível remover produto emprestado."
    )


# ── Integração ──────────────────────────────────────────────────────────────

def _baixar(client, item_id, **dados):
    arquivos = dados.pop("files", None)
    return client.request("DELETE", f"/api/items/{item_id}", data=dados, files=arquivos)


def _consultar(sql, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


@pytest.mark.integration
def test_baixa_notebook_sem_apagamento_recusada(client_gestor, item_disponivel):
    resp = _baixar(client_gestor, item_disponivel["id"], reason="Doação")
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Informe se os dados do equipamento foram apagados."
    assert client_gestor.get(f"/api/items/{item_disponivel['id']}").status_code == 200


@pytest.mark.integration
def test_baixa_notebook_com_apagamento_grava_registro(client_gestor, item_disponivel):
    resp = _baixar(
        client_gestor, item_disponivel["id"], reason="Doação", **APAGAMENTO_OK,
        files={"comprovante_apagamento": ("c.pdf", b"%PDF-1.4 comprovante", "application/pdf")},
    )
    assert resp.status_code == 200, resp.text
    hist = _consultar(
        "SELECT id FROM history WHERE item_id=%s AND operation='Exclusão'", (item_disponivel["id"],)
    )
    reg = _consultar("SELECT * FROM baixas_apagamento WHERE item_id=%s", (item_disponivel["id"],))
    assert len(reg) == 1
    assert reg[0]["history_id"] == hist[0]["id"]
    assert reg[0]["dados_apagados"] == "sim"
    assert reg[0]["metodo"] == "formatacao_segura"
    assert reg[0]["apagado_por"] == "Técnico TI"
    assert reg[0]["apagado_em"] == date.today()
    assert reg[0]["comprovante_key"].startswith("apagamentos/")
    assert reg[0]["registrado_por"] == "gestor.teste"


@pytest.mark.integration
def test_baixa_impressora_grava_nao_aplica(client_gestor, criar_item):
    item = criar_item(tipo="Impressora")
    resp = _baixar(client_gestor, item["id"], reason="Obsolescência")
    assert resp.status_code == 200, resp.text
    reg = _consultar("SELECT * FROM baixas_apagamento WHERE item_id=%s", (item["id"],))
    assert len(reg) == 1
    assert reg[0]["dados_apagados"] == "nao_aplica"
    assert reg[0]["metodo"] == "nao_aplica"
    assert reg[0]["apagado_por"] is None
    assert reg[0]["apagado_em"] is None


@pytest.mark.integration
def test_baixa_notebook_roubo_nao_aplica(client_gestor, item_disponivel):
    resp = _baixar(client_gestor, item_disponivel["id"], reason="Roubo", dados_apagados="nao_aplica")
    assert resp.status_code == 200, resp.text
    reg = _consultar("SELECT * FROM baixas_apagamento WHERE item_id=%s", (item_disponivel["id"],))
    assert reg[0]["dados_apagados"] == "nao_aplica"


@pytest.mark.integration
def test_baixa_notebook_nao_aplica_para_doacao_recusada(client_gestor, item_disponivel):
    resp = _baixar(client_gestor, item_disponivel["id"], reason="Doação", dados_apagados="nao_aplica")
    assert resp.status_code == 400
    assert resp.json()["detail"].startswith("Para Notebook, a baixa exige o apagamento dos dados (LGPD).")


@pytest.mark.integration
def test_baixa_item_sem_conserto_permitida(client_gestor, item_disponivel):
    mid = client_gestor.post(
        "/api/manutencoes", json={"item_id": item_disponivel["id"], "defeito": "Placa queimada"}
    ).json()["id"]
    resp = client_gestor.post(f"/api/manutencoes/{mid}/retorno", data={"resultado": "sem_conserto"})
    assert resp.status_code == 200
    resp = _baixar(client_gestor, item_disponivel["id"], reason="Obsolescência", **APAGAMENTO_OK)
    assert resp.status_code == 200, resp.text
    assert client_gestor.get(f"/api/items/{item_disponivel['id']}").status_code == 404


@pytest.mark.integration
def test_baixa_com_manutencao_em_andamento_recusada(client_gestor, item_disponivel):
    client_gestor.post("/api/manutencoes", json={"item_id": item_disponivel["id"], "defeito": "x"})
    resp = _baixar(client_gestor, item_disponivel["id"], reason="Obsolescência", **APAGAMENTO_OK)
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Conclua ou cancele a manutenção em andamento antes da baixa."
    assert client_gestor.get(f"/api/items/{item_disponivel['id']}").status_code == 200


@pytest.mark.integration
def test_comprovante_invalido_nao_baixa(client_gestor, item_disponivel):
    resp = _baixar(
        client_gestor, item_disponivel["id"], reason="Doação", **APAGAMENTO_OK,
        files={"comprovante_apagamento": ("c.pdf", b"MZ", "application/pdf")},
    )
    assert resp.status_code == 400
    assert resp.json()["detail"] == "Envie um PDF ou uma imagem PNG/JPG."
    assert client_gestor.get(f"/api/items/{item_disponivel['id']}").status_code == 200
    assert _consultar("SELECT * FROM baixas_apagamento") == []


@pytest.mark.integration
def test_validacao_vem_antes_de_gravar_anexos(client_gestor, item_disponivel):
    from app.core.storage import storage
    import os

    antes = set(os.listdir(os.path.join(storage.base_path, "notas_remocao")))
    resp = _baixar(
        client_gestor, item_disponivel["id"], reason="Doação",
        files={"attachment": ("n.pdf", b"conteudo", "application/pdf")},
    )
    assert resp.status_code == 400
    assert set(os.listdir(os.path.join(storage.base_path, "notas_remocao"))) == antes


@pytest.mark.integration
def test_auditoria_da_baixa_inclui_apagamento(client_gestor, item_disponivel):
    assert _baixar(client_gestor, item_disponivel["id"], reason="Doação", **APAGAMENTO_OK).status_code == 200
    linhas = _consultar("SELECT * FROM audit_log WHERE entidade='item' AND acao='excluir'")
    assert len(linhas) == 1
    depois = json.loads(linhas[0]["depois"])
    assert depois["motivo"] == "Doação"
    assert depois["apagamento"] == {
        "dados_apagados": "sim",
        "metodo": "formatacao_segura",
        "apagado_por": "Técnico TI",
        "apagado_em": date.today().isoformat(),
    }


@pytest.mark.integration
def test_constants_traz_regras_de_apagamento(client_gestor):
    corpo = client_gestor.get("/api/constants").json()
    assert corpo["tipos_exigem_apagamento"] == ["Notebook", "Desktop", "Celular", "Tablet", "HD"]
    assert corpo["motivos_sem_posse"] == ["Roubo", "Perda"]
    assert corpo["metodos_apagamento"] == {
        "formatacao_segura": "Formatação segura",
        "reset_fabrica": "Reset de fábrica",
        "destruicao_fisica": "Destruição física",
    }
