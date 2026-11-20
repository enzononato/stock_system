"""Coletor de dados da máquina.

O técnico gera um código de uso único (válido por poucos minutos) e cola o comando no
Win+R da máquina nova. As rotas públicas (script, .cmd e envio) são autorizadas só pelo
código: o banco guarda apenas o SHA-256 dele. O técnico acompanha o andamento por id.
"""
import secrets
from datetime import datetime, timedelta, timezone

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import PlainTextResponse, Response
from pydantic import ValidationError

from app.core.config import settings
from app.core.security import sha256_hex
from app.db.coleta_db import ColetaDBManager
from app.dependencies import CurrentUser, get_client_ip, gestor_or_tecnico, limiter
from app.schemas.coletor import ColetaBruta
from app.services import coletor_script
from app.services.coletor_normalizacao import normalizar

router = APIRouter(prefix="/api/coletor", tags=["coletor"])

MAX_CORPO_BYTES = 64 * 1024  # 64 KB
MSG_VENCIDO = "Este código venceu. Gere um novo na página do Patrimônio."


def _agora_utc() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def _iso_utc(momento: datetime) -> str:
    """O banco guarda UTC sem fuso; a API devolve ISO-8601 com `Z`."""
    return momento.strftime("%Y-%m-%dT%H:%M:%SZ")


def _base_url() -> str:
    return settings.APP_PUBLIC_URL.rstrip("/")


def _coleta_valida(codigo: str) -> bool:
    """Existe, ainda não recebeu dados e não venceu."""
    coleta = ColetaDBManager().por_hash(sha256_hex(codigo))
    return bool(coleta and coleta["recebido_em"] is None and coleta["expira_em"] > _agora_utc())


@router.post("", status_code=201)
def criar_coleta(user: CurrentUser = Depends(gestor_or_tecnico)):
    codigo = secrets.token_urlsafe(16)
    expira_em = _agora_utc() + timedelta(minutes=settings.COLETOR_VALIDADE_MINUTOS)
    coleta_id = ColetaDBManager().criar(sha256_hex(codigo), user.id, expira_em)
    return {
        "id": coleta_id,
        "comando": coletor_script.montar_comando(_base_url(), codigo),
        "url_arquivo": f"/api/coletor/{codigo}/coletor.cmd",
        "expira_em": _iso_utc(expira_em),
    }


@router.get("/{codigo}")
@limiter.limit(settings.COLETOR_RATE_LIMIT)
def script_do_coletor(request: Request, codigo: str):
    """Script do código, ou o aviso de vencido (inexistente, vencido, já usado ou malformado)."""
    texto = coletor_script.script_vencido()
    try:
        if _coleta_valida(codigo):
            texto = coletor_script.montar_script(_base_url(), codigo)
    except ValueError:
        pass
    return PlainTextResponse(texto, media_type="text/plain; charset=utf-8")


@router.get("/{codigo}/coletor.cmd")
@limiter.limit(settings.COLETOR_RATE_LIMIT)
def arquivo_cmd(request: Request, codigo: str):
    try:
        if _coleta_valida(codigo):
            conteudo = coletor_script.montar_cmd(_base_url(), codigo)
            return Response(
                content=conteudo.encode("utf-8"),
                media_type="application/octet-stream",
                headers={"Content-Disposition": 'attachment; filename="coletor-patrimonio.cmd"'},
            )
    except ValueError:
        pass
    raise HTTPException(status_code=410, detail=MSG_VENCIDO)


@router.post("/{codigo}/dados")
@limiter.limit(settings.COLETOR_RATE_LIMIT)
async def receber_dados(request: Request, codigo: str):
    muito_grande = HTTPException(status_code=413, detail="Dados grandes demais.")
    declarado = request.headers.get("content-length", "")
    if declarado.isdigit() and int(declarado) > MAX_CORPO_BYTES:
        raise muito_grande
    corpo = bytearray()
    async for pedaco in request.stream():
        corpo += pedaco
        if len(corpo) > MAX_CORPO_BYTES:
            raise muito_grande

    try:
        bruta = ColetaBruta.model_validate_json(bytes(corpo))
    except ValidationError:
        raise HTTPException(status_code=422, detail="Dados da coleta inválidos.") from None

    gravou = await run_in_threadpool(
        ColetaDBManager().registrar_dados,
        sha256_hex(codigo),
        bruta.model_dump(),
        get_client_ip(request),
    )
    if not gravou:
        raise HTTPException(status_code=410, detail=MSG_VENCIDO)
    return {"ok": True}


@router.get("/{coleta_id}/andamento")
def andamento(coleta_id: int, user: CurrentUser = Depends(gestor_or_tecnico)):
    db = ColetaDBManager()
    coleta = db.por_id(coleta_id)
    if coleta is None or (user.role != "Gestor" and coleta["criado_por_id"] != user.id):
        raise HTTPException(status_code=404, detail="Coleta não encontrada.")

    resposta: dict = {"expira_em": _iso_utc(coleta["expira_em"])}
    if coleta["recebido_em"] is None:
        resposta["status"] = "vencido" if coleta["expira_em"] <= _agora_utc() else "aguardando"
        return resposta

    normalizada = normalizar(ColetaBruta.model_validate(coleta["dados"]))
    campos = normalizada.campos
    resposta.update(
        status="recebido",
        campos=campos,
        tipo_sugerido=normalizada.tipo_sugerido,
        correspondencias=db.correspondencias(
            campos.get("identificador"), normalizada.macs, campos.get("anydesk")
        ),
    )
    return resposta
