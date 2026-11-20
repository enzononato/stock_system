"""Script PowerShell do coletor e comando do Win+R."""
import pytest

from app.schemas.coletor import ColetaBruta
from app.services import coletor_script
from app.services.coletor_script import (
    montar_cmd,
    montar_comando,
    montar_script,
    script_vencido,
)

CODIGO = "A" * 22


def test_comando_cabe_no_win_r():
    base = "https://patrimonio.revalle.com.br"
    comando = montar_comando(base, CODIGO)
    assert len(comando) <= 259
    assert comando == (
        'powershell -NoProfile -ExecutionPolicy Bypass -Command '
        '"[Net.ServicePointManager]::SecurityProtocol=\'Tls12\';'
        f"irm '{base}/api/coletor/{CODIGO}'|iex\""
    )


def test_script_tem_url_e_codigo():
    script = montar_script("https://x", CODIGO)
    assert f"https://x/api/coletor/{CODIGO}/dados" in script
    assert "__URL_ENVIO__" not in script
    assert "__VERSAO__" not in script


def test_script_recusa_codigo_estranho():
    with pytest.raises(ValueError):
        montar_script("https://x", "abc'; rm")
    with pytest.raises(ValueError):
        montar_comando("https://x", "abc'; rm")
    with pytest.raises(ValueError):
        montar_cmd("https://x", "abc'; rm")


def test_cmd_tem_crlf_e_comando():
    cmd = montar_cmd("https://x", CODIGO)
    assert cmd.startswith("@echo off\r\n")
    assert montar_comando("https://x", CODIGO) in cmd
    assert cmd.endswith("\r\n")
    assert "\n" not in cmd.replace("\r\n", "")


def test_template_cobre_todas_as_chaves():
    script = montar_script("https://x", CODIGO)
    for campo in ColetaBruta.model_fields:
        assert campo in script, campo
    for campo in ("tipo", "bytes", "barramento", "nome", "descricao", "mac", "ativo", "gateway"):
        assert campo in script, campo


def test_script_envia_a_versao_do_coletor():
    assert f"'{coletor_script.VERSAO_COLETOR}'" in montar_script("https://x", CODIGO)


def test_script_vencido_avisa_em_portugues():
    texto = script_vencido()
    assert "Este código venceu. Gere um novo na página do Patrimônio." in texto
    assert 'Read-Host "Aperte Enter para fechar"' in texto
