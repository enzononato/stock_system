"""Orquestra o 2FA: banco (`MfaDBManager`), cifra/TOTP (`mfa`) e auditoria.

Nunca grava segredo, hash nem códigos na auditoria.
"""
from app.db.mfa_db import MfaDBManager
from app.services import audit
from app.services.audit import AuditContext
from app.services.mfa import (
    CadastroNaoIniciado,
    CodigoInvalido,
    MfaErro,
    MfaIndisponivel,
    MfaJaAtivo,
    cifrar_segredo,
    decifrar_segredo,
    gerar_codigos_recuperacao,
    gerar_segredo,
    hash_codigo_recuperacao,
    mfa_configurado,
    parece_codigo_totp,
    uri_provisionamento,
    verificar_codigo_totp,
)


def iniciar_cadastro(db: MfaDBManager, usuario: dict) -> tuple[str, str]:
    """Começa o cadastro do autenticador: devolve (segredo, uri otpauth).

    O segredo fica cifrado e pendente (`totp_ativo=0`) até a confirmação; iniciar de
    novo sobrescreve um pendente anterior."""
    if usuario.get("totp_ativo"):
        raise MfaJaAtivo()
    if not mfa_configurado():
        raise MfaIndisponivel()
    segredo = gerar_segredo()
    db.salvar_segredo_pendente(usuario["id"], cifrar_segredo(segredo))
    return segredo, uri_provisionamento(segredo, usuario["username"])


def confirmar_cadastro(db: MfaDBManager, usuario: dict, codigo: str, ctx: AuditContext) -> list[str]:
    """Confirma o cadastro com um código do aplicativo, ativa o 2FA e devolve os 10
    códigos de recuperação em texto (única vez em que existem em texto)."""
    if usuario.get("totp_ativo"):
        raise MfaJaAtivo()
    if not mfa_configurado():
        raise MfaIndisponivel()
    estado = db.estado(usuario["id"])
    if not estado or not estado.get("totp_secret_cifrado"):
        raise CadastroNaoIniciado()
    segredo = decifrar_segredo(estado["totp_secret_cifrado"])
    passo = verificar_codigo_totp(segredo, codigo, None)
    if passo is None:
        raise CodigoInvalido()
    codigos = gerar_codigos_recuperacao()
    db.ativar(usuario["id"], [hash_codigo_recuperacao(c) for c in codigos], passo)
    audit.record(
        ctx, "usuario", usuario["id"], "ativar_2fa",
        antes={"totp_ativo": False}, depois={"totp_ativo": True},
    )
    return codigos


def verificar_segundo_fator(db: MfaDBManager, usuario: dict, codigo: str, ctx: AuditContext) -> int | None:
    """Confere o código da 2ª etapa do login.

    Código TOTP: devolve None. Código de recuperação: consome-o e devolve quantos
    restam. Código errado, repetido ou já usado: `CodigoInvalido`. Sem a chave de
    cifra, o TOTP levanta `MfaIndisponivel` (o de recuperação não precisa da chave).
    """
    usuario_id = usuario["id"]
    if parece_codigo_totp(codigo):
        estado = db.estado(usuario_id)
        if not estado or not estado.get("totp_secret_cifrado"):
            raise CodigoInvalido()
        segredo = decifrar_segredo(estado["totp_secret_cifrado"])
        passo = verificar_codigo_totp(segredo, codigo, estado["totp_ultimo_passo"])
        if passo is None or not db.registrar_passo(usuario_id, passo):
            raise CodigoInvalido()
        return None

    if not db.consumir_codigo(usuario_id, hash_codigo_recuperacao(codigo)):
        raise CodigoInvalido()
    restantes = db.contar_codigos_restantes(usuario_id)
    audit.record(
        ctx, "usuario", usuario_id, "usar_codigo_recuperacao",
        depois={"codigos_restantes": restantes},
    )
    return restantes


def regenerar_codigos(db: MfaDBManager, usuario: dict, ctx: AuditContext) -> list[str]:
    """Troca os códigos de recuperação por 10 novos (os antigos deixam de valer)."""
    if not usuario.get("totp_ativo"):
        raise MfaErro("A verificação em duas etapas não está ativa.")
    codigos = gerar_codigos_recuperacao()
    db.substituir_codigos(usuario["id"], [hash_codigo_recuperacao(c) for c in codigos])
    audit.record(
        ctx, "usuario", usuario["id"], "regenerar_codigos_2fa",
        depois={"codigos": len(codigos)},
    )
    return codigos


def desativar_2fa(db: MfaDBManager, usuario: dict, ctx: AuditContext) -> None:
    """Desativa o 2FA do próprio usuário (nunca para o Gestor, que é obrigado a tê-lo)."""
    if usuario.get("role") == "Gestor":
        raise MfaErro("Para o Gestor a verificação em duas etapas é obrigatória.")
    if not usuario.get("totp_ativo"):
        raise MfaErro("A verificação em duas etapas não está ativa.")
    db.desativar(usuario["id"])
    audit.record(
        ctx, "usuario", usuario["id"], "desativar_2fa",
        antes={"totp_ativo": True}, depois={"totp_ativo": False},
    )


def zerar_2fa(db: MfaDBManager, usuario: dict, ctx: AuditContext, origem: str) -> None:
    """Zera o 2FA do usuário (segredo, códigos) e encerra as sessões dele. Quem chama
    decide quem pode (Gestor na tela de usuários, ou o script no servidor)."""
    antes = {"totp_ativo": bool(usuario.get("totp_ativo"))}
    db.zerar(usuario["id"])
    audit.record(
        ctx, "usuario", usuario["id"], "zerar_2fa",
        antes=antes, depois={"totp_ativo": False, "origem": origem},
    )
