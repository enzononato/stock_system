"""Versão da sessão por usuário.

Todo token (access, refresh, etapa de 2FA) leva o claim `sv` = `usuarios.sessao_versao`
do momento da emissão. Redefinir a senha ou zerar o 2FA incrementa o contador e
todos os tokens antigos deixam de valer. Token sem `sv` (emitido antes do deploy)
vale como `0`.

Este módulo não importa `app.dependencies` (que importa este).
"""
from fastapi import HTTPException, status

from app.db.user_manager_db import UserDBManager

MSG_SESSAO_ENCERRADA = "Sessão encerrada. Faça login novamente."


def dados_do_token(usuario: dict) -> dict:
    """Claims básicos de um token para o usuário (linha de `usuarios`)."""
    return {
        "sub": str(usuario["id"]),
        "username": usuario["username"],
        "role": usuario["role"],
        "sv": int(usuario.get("sessao_versao") or 0),
    }


def _sessao_encerrada() -> HTTPException:
    return HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail=MSG_SESSAO_ENCERRADA,
        headers={"WWW-Authenticate": "Bearer"},
    )


def validar_sessao_ativa(payload: dict) -> None:
    """401 se o usuário não existe mais ou se a versão do token difere da atual."""
    try:
        usuario_id = int(payload["sub"])
    except (KeyError, TypeError, ValueError):
        raise _sessao_encerrada()
    estado = UserDBManager().get_estado_sessao(usuario_id)
    if estado is None or int(payload.get("sv", 0)) != int(estado["sessao_versao"] or 0):
        raise _sessao_encerrada()
