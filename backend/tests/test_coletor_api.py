"""Rotas do coletor de dados da máquina."""
import hashlib
import re

import pytest

pytestmark = pytest.mark.integration

MSG_VENCIDO = "Este código venceu. Gere um novo na página do Patrimônio."

MAC_DELL = "A4:BB:6D:11:22:33"

PAYLOAD_DELL = {
    "versao_coletor": "1",
    "fabricante": "Dell Inc.",
    "modelo": "Latitude 3420",
    "serial": "ABC1234",
    "host": "NB-TI-01",
    "cpu": "11th Gen Intel(R) Core(TM) i5-1135G7 @ 2.40GHz",
    "sistema": "Microsoft Windows 11 Pro",
    "ram_bytes": 17179869184,
    "chassi": [10],
    "no_dominio": True,
    "discos": [{"tipo": "SSD", "bytes": 512110190592, "barramento": "NVMe"}],
    "adaptadores": [
        {
            "nome": "Ethernet",
            "descricao": "Intel(R) Ethernet",
            "mac": MAC_DELL,
            "tipo": "ethernet",
            "ativo": True,
            "gateway": True,
        }
    ],
}


def _consultar(sql, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return cur.fetchall()


def _criar(cliente):
    resp = cliente.post("/api/coletor")
    assert resp.status_code == 201, resp.text
    corpo = resp.json()
    codigo = re.search(r"/api/coletor/([A-Za-z0-9_-]+)'", corpo["comando"]).group(1)
    return corpo, codigo


def _enviar(client, codigo, payload=None):
    return client.post(f"/api/coletor/{codigo}/dados", json=payload or PAYLOAD_DELL)


def test_criar_coleta_tecnico(client_tecnico):
    corpo, codigo = _criar(client_tecnico)

    assert "/api/coletor/" in corpo["comando"]
    assert codigo in corpo["comando"]
    assert corpo["url_arquivo"] == f"/api/coletor/{codigo}/coletor.cmd"
    assert corpo["expira_em"].endswith("Z")
    assert isinstance(corpo["id"], int)

    linhas = _consultar("SELECT codigo_hash FROM coletas_maquina WHERE id = %s", (corpo["id"],))
    assert linhas[0]["codigo_hash"] == hashlib.sha256(codigo.encode()).hexdigest()
    assert not _consultar("SELECT id FROM coletas_maquina WHERE codigo_hash = %s", (codigo,))


def test_aprendiz_nao_cria(client_aprendiz):
    assert client_aprendiz.post("/api/coletor").status_code == 403


def test_sem_login_nao_cria(client):
    assert client.post("/api/coletor").status_code == 401


def test_script_publico_e_vencido(client, client_tecnico):
    _, codigo = _criar(client_tecnico)

    resp = client.get(f"/api/coletor/{codigo}")
    assert resp.status_code == 200
    assert resp.headers["content-type"] == "text/plain; charset=utf-8"
    assert f"/api/coletor/{codigo}/dados" in resp.text
    assert "Este código venceu" not in resp.text

    inventado = client.get("/api/coletor/" + "x" * 22)
    assert inventado.status_code == 200
    assert "Este código venceu" in inventado.text

    curto = client.get("/api/coletor/abc")
    assert curto.status_code == 200
    assert "Este código venceu" in curto.text


def test_cmd_attachment(client, client_tecnico):
    _, codigo = _criar(client_tecnico)

    resp = client.get(f"/api/coletor/{codigo}/coletor.cmd")
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("application/octet-stream")
    assert resp.headers["content-disposition"] == 'attachment; filename="coletor-patrimonio.cmd"'
    assert f"/api/coletor/{codigo}" in resp.text
    assert resp.text.startswith("@echo off")

    assert client.get(f"/api/coletor/{'y' * 22}/coletor.cmd").status_code == 410
    assert client.get("/api/coletor/curto/coletor.cmd").json() == {"detail": MSG_VENCIDO}


def test_envio_ok_e_andamento(client, client_tecnico):
    corpo, codigo = _criar(client_tecnico)

    antes = client_tecnico.get(f"/api/coletor/{corpo['id']}/andamento").json()
    assert antes["status"] == "aguardando"
    assert antes["expira_em"].endswith("Z")
    assert "campos" not in antes

    resp = _enviar(client, codigo)
    assert resp.status_code == 200
    assert resp.json() == {"ok": True}

    dados = client_tecnico.get(f"/api/coletor/{corpo['id']}/andamento").json()
    assert dados["status"] == "recebido"
    assert dados["campos"]["brand"] == "Dell"
    assert dados["campos"]["identificador"] == "ABC1234"
    assert dados["tipo_sugerido"] == "Notebook"
    assert dados["correspondencias"] == []


def test_envio_repetido_410_mantem_primeiro(client, client_tecnico):
    corpo, codigo = _criar(client_tecnico)
    assert _enviar(client, codigo).status_code == 200

    segundo = _enviar(client, codigo, {**PAYLOAD_DELL, "fabricante": "HP Inc."})
    assert segundo.status_code == 410
    assert segundo.json() == {"detail": MSG_VENCIDO}

    dados = client_tecnico.get(f"/api/coletor/{corpo['id']}/andamento").json()
    assert dados["status"] == "recebido"
    assert dados["campos"]["brand"] == "Dell"

    script = client.get(f"/api/coletor/{codigo}")
    assert script.status_code == 200
    assert "Este código venceu" in script.text


def test_envio_vencido_410(client, client_tecnico, executar_sql):
    corpo, codigo = _criar(client_tecnico)
    executar_sql(
        "UPDATE coletas_maquina SET expira_em = UTC_TIMESTAMP() - INTERVAL 1 MINUTE WHERE id = %s",
        (corpo["id"],),
    )

    assert _enviar(client, codigo).status_code == 410
    assert "Este código venceu" in client.get(f"/api/coletor/{codigo}").text
    assert client.get(f"/api/coletor/{codigo}/coletor.cmd").status_code == 410
    assert client_tecnico.get(f"/api/coletor/{corpo['id']}/andamento").json()["status"] == "vencido"


def test_envio_codigo_inexistente_410(client):
    resp = _enviar(client, "z" * 22)
    assert resp.status_code == 410
    assert resp.json() == {"detail": MSG_VENCIDO}
    assert _enviar(client, "curto").status_code == 410


def test_envio_grande_413(client, client_tecnico):
    corpo, codigo = _criar(client_tecnico)

    resp = client.post(
        f"/api/coletor/{codigo}/dados",
        content=b"{" + b" " * 70000 + b"}",
        headers={"content-type": "application/json"},
    )
    assert resp.status_code == 413
    assert resp.json() == {"detail": "Dados grandes demais."}
    linha = _consultar("SELECT dados, recebido_em FROM coletas_maquina WHERE id = %s", (corpo["id"],))[0]
    assert linha["dados"] is None and linha["recebido_em"] is None


def test_envio_invalido_422(client, client_tecnico):
    corpo, codigo = _criar(client_tecnico)

    resp = _enviar(client, codigo, {**PAYLOAD_DELL, "ram_bytes": "muito"})
    assert resp.status_code == 422
    assert resp.json() == {"detail": "Dados da coleta inválidos."}

    texto_enorme = client.post(
        f"/api/coletor/{codigo}/dados", json={**PAYLOAD_DELL, "modelo": "<b>" + "A" * 5000}
    )
    assert texto_enorme.status_code == 422

    nao_json = client.post(
        f"/api/coletor/{codigo}/dados", content=b"isto nao e json", headers={"content-type": "application/json"}
    )
    assert nao_json.status_code == 422

    linha = _consultar("SELECT dados, recebido_em FROM coletas_maquina WHERE id = %s", (corpo["id"],))[0]
    assert linha["dados"] is None and linha["recebido_em"] is None
    # o código continua utilizável depois de um envio inválido
    assert _enviar(client, codigo).status_code == 200


def test_andamento_de_outro_tecnico_404(client_tecnico, _fastapi_app, criar_usuario, senha_padrao_teste):
    from tests.conftest import _client_autenticado, _login_e_obter_access_token

    corpo, _ = _criar(client_tecnico)
    outro = criar_usuario("tecnico.dois", senha_padrao_teste, "Técnico")
    token = _login_e_obter_access_token(client_tecnico, outro["email"], senha_padrao_teste)
    client_outro = _client_autenticado(_fastapi_app, token)

    resp = client_outro.get(f"/api/coletor/{corpo['id']}/andamento")
    assert resp.status_code == 404
    assert resp.json() == {"detail": "Coleta não encontrada."}
    assert client_tecnico.get("/api/coletor/999999/andamento").status_code == 404


def test_gestor_ve_andamento_de_qualquer_um(client_tecnico, client_gestor):
    corpo, _ = _criar(client_tecnico)
    resp = client_gestor.get(f"/api/coletor/{corpo['id']}/andamento")
    assert resp.status_code == 200
    assert resp.json()["status"] == "aguardando"


def test_aprendiz_nao_ve_andamento(client_tecnico, client_aprendiz):
    corpo, _ = _criar(client_tecnico)
    assert client_aprendiz.get(f"/api/coletor/{corpo['id']}/andamento").status_code == 403


def test_andamento_com_correspondencia(client, client_tecnico, criar_item):
    item = criar_item(endereco_fisico="A4-BB-6D-11-22-33", identificador="OUTRO-SERIAL")
    corpo, codigo = _criar(client_tecnico)
    assert _enviar(client, codigo).status_code == 200

    dados = client_tecnico.get(f"/api/coletor/{corpo['id']}/andamento").json()
    assert len(dados["correspondencias"]) == 1
    achado = dados["correspondencias"][0]
    assert achado["id"] == item["id"]
    assert achado["motivos"] == ["mac"]
    assert achado["ativo"] is True


def test_coleta_nao_grava_em_items(client, client_tecnico):
    _, codigo = _criar(client_tecnico)
    antes = _consultar("SELECT COUNT(*) AS n FROM items")[0]["n"]
    assert _enviar(client, codigo).status_code == 200
    assert _consultar("SELECT COUNT(*) AS n FROM items")[0]["n"] == antes


@pytest.mark.unit
def test_job_limpar_coletas_registrado():
    from app import jobs
    from app.services import scheduler

    scheduler.limpar_jobs()
    try:
        jobs.registrar_jobs()
        assert "limpar_coletas" in scheduler.jobs_registrados()
        assert scheduler.jobs_registrados()["limpar_coletas"].diariamente_as == "03:30"
    finally:
        scheduler.limpar_jobs()
