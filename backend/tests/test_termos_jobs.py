"""Jobs do agendador: envio agendado de termos e consulta periódica ao RevSign."""
from datetime import timedelta

import pytest

from tests.fixtures_termos import (  # noqa: F401
    colaborador,
    corpo_emprestimo_colaborador,
    criar_colaborador,
    revsign_fake,
    unidade_juazeiro,
)

pytestmark = pytest.mark.integration

ARQUIVO_PDF = {"signed_pdf": ("termo.pdf", b"%PDF-1.4 conteudo", "application/pdf")}


def _emprestar(client_gestor, item_id, colaborador_id):
    resp = client_gestor.post("/api/loans", json=corpo_emprestimo_colaborador(item_id, colaborador_id))
    assert resp.status_code == 200, resp.text
    return resp.json()["termo"]


def _indisponivel(client_gestor, item_id, colaborador_id):
    termo = _emprestar(client_gestor, item_id, colaborador_id)
    assert client_gestor.post(f"/api/loans/{item_id}/confirm", files=ARQUIVO_PDF).status_code == 200
    return termo


def _termo(termo_id):
    from app.dependencies import get_termo_db

    return get_termo_db().obter(termo_id)


def _agendar(client_gestor, item_id, quando):
    resp = client_gestor.post(
        f"/api/items/{item_id}/termos", json={"agendado_para": quando.strftime("%Y-%m-%dT%H:%M")}
    )
    assert resp.status_code == 201, resp.text
    assert resp.json()["status"] == "agendado"
    return resp.json()


def test_envio_agendado_sai_no_tique_certo(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    from app.core.tempo import agora_local
    from app.services.termos_retorno import enviar_agendados

    item_id = item_disponivel["id"]
    _indisponivel(client_gestor, item_id, colaborador["id"])
    agora = agora_local()
    agendado = _agendar(client_gestor, item_id, agora + timedelta(hours=1))
    assert len(revsign_fake.criados) == 1  # só a entrega

    assert enviar_agendados(agora + timedelta(minutes=30)) == 0
    assert _termo(agendado["id"])["status"] == "agendado"

    assert enviar_agendados(agora + timedelta(minutes=61)) == 1
    assert _termo(agendado["id"])["status"] == "enviado"
    assert len(revsign_fake.criados) == 2
    # rodar de novo (estado do job não gravado) não envia outra vez
    assert enviar_agendados(agora + timedelta(minutes=62)) == 0
    assert len(revsign_fake.criados) == 2


def test_agendado_de_item_que_mudou_de_situacao_e_cancelado(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake, executar_sql
):
    from app.core.tempo import agora_local
    from app.services.termos import MOTIVO_SITUACAO_MUDOU
    from app.services.termos_retorno import enviar_agendados

    item_id = item_disponivel["id"]
    _indisponivel(client_gestor, item_id, colaborador["id"])
    agora = agora_local()
    agendado = _agendar(client_gestor, item_id, agora + timedelta(hours=1))
    executar_sql("UPDATE items SET status='Disponível' WHERE id=%s", (item_id,))

    assert enviar_agendados(agora + timedelta(minutes=61)) == 1
    atual = _termo(agendado["id"])
    assert atual["status"] == "cancelado"
    assert atual["motivo"] == MOTIVO_SITUACAO_MUDOU
    assert len(revsign_fake.criados) == 1


def test_erro_em_uma_solicitacao_nao_para_as_outras(
    client_gestor, item_disponivel, criar_item, colaborador, unidade_juazeiro, revsign_fake, monkeypatch
):
    from app.core.tempo import agora_local
    from app.services import termos, termos_retorno

    primeiro, segundo = item_disponivel["id"], criar_item()["id"]
    for item_id in (primeiro, segundo):
        _indisponivel(client_gestor, item_id, colaborador["id"])
    agora = agora_local()
    a = _agendar(client_gestor, primeiro, agora + timedelta(hours=1))
    b = _agendar(client_gestor, segundo, agora + timedelta(hours=1))

    original = termos.enviar

    def enviar_quebrando_o_primeiro(termo_id, ctx, **kwargs):
        if termo_id == a["id"]:
            raise RuntimeError("falha inesperada")
        return original(termo_id, ctx, **kwargs)

    monkeypatch.setattr(termos, "enviar", enviar_quebrando_o_primeiro)
    assert termos_retorno.enviar_agendados(agora + timedelta(hours=2)) == 1
    assert _termo(a["id"])["status"] == "agendado"
    assert _termo(b["id"])["status"] == "enviado"


def test_consulta_periodica_conclui_enviado(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    from app.services.termos_retorno import consultar_enviados

    item_id = item_disponivel["id"]
    termo = _emprestar(client_gestor, item_id, colaborador["id"])
    assert consultar_enviados() == 0  # ainda pendente no RevSign
    revsign_fake.marcar_assinado("1001")

    assert consultar_enviados() == 1
    assert _termo(termo["id"])["status"] == "assinado"
    assert client_gestor.get(f"/api/items/{item_id}").json()["status"] == "Indisponível"
    assert consultar_enviados() == 0


def test_falha_no_download_mantem_enviado_e_tenta_de_novo(
    client_gestor, item_disponivel, colaborador, unidade_juazeiro, revsign_fake
):
    from app.services.termos_retorno import consultar_enviados

    termo = _emprestar(client_gestor, item_disponivel["id"], colaborador["id"])
    revsign_fake.marcar_assinado("1001")
    revsign_fake.falhar_em = {"baixar_assinado"}

    assert consultar_enviados() == 0
    atual = _termo(termo["id"])
    assert atual["status"] == "enviado"
    assert atual["motivo"].startswith("Falha ao baixar o PDF assinado")

    revsign_fake.falhar_em = set()
    assert consultar_enviados() == 1
    atual = _termo(termo["id"])
    assert atual["status"] == "assinado"
    assert atual["motivo"] is None


def test_jobs_desligados_nao_fazem_nada(client_gestor, item_indisponivel):
    from app.services.termos_retorno import consultar_enviados, enviar_agendados

    assert enviar_agendados() == 0
    assert consultar_enviados() == 0


def test_registrar_jobs_so_com_revsign_ligado(_test_schema, monkeypatch):
    from app import jobs
    from app.core import config
    from app.services import scheduler

    nomes = {"termos_envios_agendados", "termos_consulta_revsign"}
    scheduler.limpar_jobs()
    try:
        monkeypatch.setattr(config.settings, "REVSIGN_ENABLED", False)
        jobs.registrar_jobs()
        assert not nomes & set(scheduler.jobs_registrados())

        scheduler.limpar_jobs()
        monkeypatch.setattr(config.settings, "REVSIGN_ENABLED", True)
        jobs.registrar_jobs()
        assert nomes <= set(scheduler.jobs_registrados())
        assert scheduler.jobs_registrados()["termos_envios_agendados"].a_cada_segundos == 60
        assert scheduler.jobs_registrados()["termos_consulta_revsign"].a_cada_segundos == 600
    finally:
        scheduler.limpar_jobs()
