"""Anexos de equipamento (nota fiscal, manutenção, comprovante de apagamento).

O tipo do arquivo é decidido pelo CONTEÚDO (assinatura dos primeiros bytes), nunca
pelo nome nem pelo Content-Type enviados pelo cliente. O nome do cliente é
descartado: a chave no storage é montada aqui.
"""
import uuid
from datetime import datetime
from typing import Optional

from fastapi import HTTPException, UploadFile
from fastapi.responses import Response

from app.core.storage import storage

MAX_ANEXO_BYTES = 10 * 1024 * 1024  # 10MB

CATEGORIA_NF = "notas_fiscais"
CATEGORIA_MANUTENCAO = "manutencoes"
CATEGORIA_APAGAMENTO = "apagamentos"

_MEDIA_TYPES = {"pdf": "application/pdf", "png": "image/png", "jpg": "image/jpeg"}


def tipo_do_anexo(conteudo: bytes) -> Optional[str]:
    """Extensão pelo conteúdo: pdf, png ou jpg; `None` para qualquer outra coisa."""
    if conteudo.startswith(b"%PDF"):
        return "pdf"
    if conteudo.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if conteudo.startswith(b"\xff\xd8\xff"):
        return "jpg"
    return None


def validar_anexo(conteudo: bytes) -> str:
    """Devolve a extensão do anexo ou levanta HTTPException (400 vazio/tipo, 413 tamanho)."""
    if not conteudo:
        raise HTTPException(status_code=400, detail="O arquivo enviado está vazio.")
    if len(conteudo) > MAX_ANEXO_BYTES:
        raise HTTPException(status_code=413, detail="Arquivo excede o tamanho máximo de 10MB.")
    extensao = tipo_do_anexo(conteudo)
    if extensao is None:
        raise HTTPException(status_code=400, detail="Envie um PDF ou uma imagem PNG/JPG.")
    return extensao


async def salvar_anexo(upload: UploadFile, categoria: str, prefixo: str) -> str:
    """Lê, valida e grava o anexo; devolve a chave do storage (`<categoria>/<arquivo>`)."""
    conteudo = await upload.read()
    extensao = validar_anexo(conteudo)
    nome = f"{prefixo}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}.{extensao}"
    return storage.save(categoria, nome, conteudo)


def resposta_anexo(key: str, nome_base: str) -> Response:
    """Download do anexo como attachment, com o media type pela extensão da chave."""
    if not storage.exists(key):
        raise HTTPException(status_code=404, detail="Arquivo não encontrado.")
    extensao = key.rsplit(".", 1)[-1].lower() if "." in key else ""
    return Response(
        content=storage.read(key),
        media_type=_MEDIA_TYPES.get(extensao, "application/octet-stream"),
        headers={"Content-Disposition": f'attachment; filename="{nome_base}.{extensao}"'},
    )
