"""Script PowerShell do coletor e o comando que o técnico cola no Win+R.

O comando baixa este script do backend e o executa em memória; o script lê os dados da
máquina (somente leitura) e os envia por POST para a rota pública do código.
"""
import re
from pathlib import Path

VERSAO_COLETOR = "1"

_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "coletor" / "coletor.ps1"
_CODIGO_VALIDO = re.compile(r"^[A-Za-z0-9_-]{16,64}$")

_MSG_VENCIDO = "Este código venceu. Gere um novo na página do Patrimônio."


def _validar_codigo(codigo: str) -> None:
    if not _CODIGO_VALIDO.fullmatch(codigo or ""):
        raise ValueError("Código do coletor inválido.")


def montar_comando(base_url: str, codigo: str) -> str:
    """Linha única do Win+R (cabe nos 259 caracteres do diálogo Executar)."""
    _validar_codigo(codigo)
    return (
        "powershell -NoProfile -ExecutionPolicy Bypass -Command "
        "\"[Net.ServicePointManager]::SecurityProtocol='Tls12';"
        f"irm '{base_url}/api/coletor/{codigo}'|iex\""
    )


def montar_script(base_url: str, codigo: str) -> str:
    _validar_codigo(codigo)
    texto = _TEMPLATE.read_text(encoding="utf-8")
    return texto.replace("__URL_ENVIO__", f"{base_url}/api/coletor/{codigo}/dados").replace(
        "__VERSAO__", VERSAO_COLETOR
    )


def script_vencido() -> str:
    return f'Write-Host "{_MSG_VENCIDO}"\nRead-Host "Aperte Enter para fechar"\n'


def montar_cmd(base_url: str, codigo: str) -> str:
    """Arquivo .cmd equivalente ao comando (finais de linha CRLF)."""
    return "@echo off\r\n" + montar_comando(base_url, codigo) + "\r\n"
