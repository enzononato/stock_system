"""Versão da sessão por usuário.

Todo token (access, refresh, etapa de 2FA) leva o claim `sv` = `usuarios.sessao_versao`
do momento da emissão. Redefinir a senha ou zerar o 2FA incrementa o contador e
todos os tokens antigos deixam de valer. Token sem `sv` (emitido antes do deploy)
vale como `0`.

Este módulo não importa `app.dependencies` (que importa este).
"""
from fastapi import HTTPException, status

from app.db.user_manager_db import UserDBManager
from app.services.mfa import mfa_configurado

MSG_SESSAO_ENCERRADA = "Sessão encerrada. Faça login novamente."
MSG_CADASTRO_OBRIGATORIO = "Cadastre a verificação em duas etapas para continuar. Entre novamente."


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


def exige_cadastro_2fa(usuario: dict) -> bool:
    """O Gestor sem autenticador precisa cadastrar um antes de usar o sistema —
    exceto se o servidor está sem a chave de cifra (aí ninguém é obrigado)."""
    return usuario.get("role") == "Gestor" and not usuario.get("totp_ativo") and mfa_configurado()


def validar_sessao_ativa(payload: dict) -> None:
    """401 se o usuário não existe mais, se a versão do token difere da atual ou se
    é um Gestor que ainda não cadastrou o 2FA (vale para tokens anteriores ao cadastro)."""
    try:
        usuario_id = int(payload["sub"])
    except (KeyError, TypeError, ValueError):
        raise _sessao_encerrada()
    estado = UserDBManager().get_estado_sessao(usuario_id)
    if estado is None or int(payload.get("sv", 0)) != int(estado["sessao_versao"] or 0):
        raise _sessao_encerrada()
    if exige_cadastro_2fa(estado):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=MSG_CADASTRO_OBRIGATORIO,
            headers={"WWW-Authenticate": "Bearer"},
        )
