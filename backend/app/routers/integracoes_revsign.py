"""Webhook do RevSign. Público (sem login): a autenticação é o segredo compartilhado."""
from fastapi import APIRouter, HTTPException, Request
from fastapi.concurrency import run_in_threadpool

from app.services import termos, termos_retorno
from app.services.revsign_client import WebhookInvalido, WebhookMalformado, get_revsign_client

router = APIRouter(prefix="/api/integrations/revsign", tags=["integracoes"])

MAX_CORPO_BYTES = 1024 * 1024  # 1 MB


@router.post("/webhook")
async def webhook_revsign(request: Request):
    if not termos.revsign_ligado():
        raise HTTPException(status_code=404, detail="Integração com o RevSign desligada.")

    muito_grande = HTTPException(status_code=413, detail="Corpo do webhook muito grande.")
    declarado = request.headers.get("content-length", "")
    if declarado.isdigit() and int(declarado) > MAX_CORPO_BYTES:
        raise muito_grande
    corpo = bytearray()
    async for pedaco in request.stream():
        corpo += pedaco
        if len(corpo) > MAX_CORPO_BYTES:
            raise muito_grande

    try:
        evento = get_revsign_client().validar_webhook(request.headers, bytes(corpo))
    except WebhookInvalido:
        raise HTTPException(status_code=401, detail="Webhook não autorizado.") from None
    except WebhookMalformado:
        raise HTTPException(status_code=400, detail="Corpo do webhook inválido.") from None

    resultado = await run_in_threadpool(termos_retorno.processar_evento, evento)
    return {"ok": True, "resultado": resultado}
