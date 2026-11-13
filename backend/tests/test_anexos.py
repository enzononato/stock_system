"""Serviço de anexos (NF, manutenção, comprovante de apagamento)."""
import asyncio
import io
import os

import pytest
from fastapi import HTTPException, UploadFile

from app.core.storage import storage
from app.services.anexos import (
    CATEGORIA_NF,
    MAX_ANEXO_BYTES,
    resposta_anexo,
    salvar_anexo,
    tipo_do_anexo,
    validar_anexo,
)

pytestmark = pytest.mark.unit


def test_tipo_do_anexo_pelo_conteudo():
    assert tipo_do_anexo(b"%PDF-1.4") == "pdf"
    assert tipo_do_anexo(b"\x89PNG\r\n\x1a\nxxxx") == "png"
    assert tipo_do_anexo(b"\xff\xd8\xff\xe0abc") == "jpg"
    assert tipo_do_anexo(b"MZ\x90\x00") is None


def test_validar_anexo_vazio_400():
    with pytest.raises(HTTPException) as exc:
        validar_anexo(b"")
    assert exc.value.status_code == 400
    assert exc.value.detail == "O arquivo enviado está vazio."


def test_validar_anexo_acima_de_10mb_413():
    with pytest.raises(HTTPException) as exc:
        validar_anexo(b"%PDF" + b"0" * MAX_ANEXO_BYTES)
    assert exc.value.status_code == 413
    assert exc.value.detail == "Arquivo excede o tamanho máximo de 10MB."


def test_validar_anexo_no_limite_aceito():
    assert validar_anexo(b"%PDF" + b"0" * (MAX_ANEXO_BYTES - 4)) == "pdf"


def test_validar_anexo_executavel_400():
    with pytest.raises(HTTPException) as exc:
        validar_anexo(b"MZ\x90\x00\x03")
    assert exc.value.status_code == 400
    assert exc.value.detail == "Envie um PDF ou uma imagem PNG/JPG."


def test_salvar_anexo_ignora_nome_do_cliente(tmp_path, monkeypatch):
    monkeypatch.setattr(storage, "base_path", str(tmp_path))
    upload = UploadFile(filename="../../x.exe", file=io.BytesIO(b"%PDF-1.4"))
    key = asyncio.run(salvar_anexo(upload, CATEGORIA_NF, "nf_1"))
    assert key.startswith("notas_fiscais/nf_1_")
    assert key.endswith(".pdf")
    assert ".." not in key and "exe" not in key
    assert os.path.exists(os.path.join(storage.base_path, key))


def test_resposta_anexo_media_type(tmp_path, monkeypatch):
    monkeypatch.setattr(storage, "base_path", str(tmp_path))
    key = storage.save("manutencoes", "teste_resposta.png", b"\x89PNG\r\n\x1a\nabc")
    resp = resposta_anexo(key, "manutencao_9")
    assert resp.media_type == "image/png"
    assert resp.headers["content-disposition"] == 'attachment; filename="manutencao_9.png"'
    assert resp.body == b"\x89PNG\r\n\x1a\nabc"


def test_resposta_anexo_inexistente_404(tmp_path, monkeypatch):
    monkeypatch.setattr(storage, "base_path", str(tmp_path))
    with pytest.raises(HTTPException) as exc:
        resposta_anexo("notas_fiscais/nao_existe.pdf", "x")
    assert exc.value.status_code == 404
    assert exc.value.detail == "Arquivo não encontrado."
