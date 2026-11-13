"""Orquestra o 2FA: banco (`MfaDBManager`), cifra/TOTP (`mfa`) e auditoria.

Nunca grava segredo, hash nem códigos na auditoria.
"""
from datetime import datetime, timedelta, timezone

from app.db.mfa_db import MfaDBManager
from app.services import audit
from app.services.audit import AuditContext
from app.services.mfa import (
    CadastroNaoIniciado,
    CodigoInvalido,
    MfaBloqueado,
    MfaErro,
    MfaIndisponivel,
    MfaJaAtivo,
    TokenEsgotado,
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
    # A gravação só vale com o 2FA inativo (condição na própria UPDATE): se alguém
    # ativou entre a leitura de `usuario` e aqui, o 2FA ativo não é derrubado.
    if not db.salvar_segredo_pendente(usuario["id"], cifrar_segredo(segredo)):
        raise MfaJaAtivo()
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
    # Ativa só o segredo cujo código acabou de ser verificado.
    if not db.ativar(
        usuario["id"], [hash_codigo_recuperacao(c) for c in codigos], passo, estado["totp_secret_cifrado"]
    ):
        raise CodigoInvalido()
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


MAX_FALHAS_POR_TOKEN = 5
MAX_FALHAS_POR_USUARIO = 10
JANELA_FALHAS = timedelta(minutes=15)
DURACAO_BLOQUEIO = timedelta(minutes=15)


def _agora_utc() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def exigir_verificacao_liberada(db: MfaDBManager, usuario: dict) -> None:
    """`MfaBloqueado` (429) enquanto o usuário está no bloqueio temporário da 2ª etapa."""
    ate = db.bloqueado_ate(usuario["id"])
    if ate is not None and ate > _agora_utc():
        raise MfaBloqueado()


def registrar_falha_segundo_fator(db: MfaDBManager, usuario: dict, jti: str, ctx: AuditContext) -> None:
    """Contabiliza um código errado da 2ª etapa e audita `falha_2fa` (sem código nem token).

    Levanta `MfaBloqueado` quando esta falha é a 10ª do usuário em 15 min (bloqueio de 15 min)
    e `TokenEsgotado` na 5ª falha do mesmo token; em ambos o chamador revoga o token."""
    agora = _agora_utc()
    do_token, do_usuario = db.registrar_falha(usuario["id"], jti, agora, agora - JANELA_FALHAS)
    audit.record(
        ctx, "usuario", usuario["id"], "falha_2fa",
        depois={"falhas_no_token": do_token, "falhas_em_15_min": do_usuario},
    )
    if do_usuario >= MAX_FALHAS_POR_USUARIO:
        db.bloquear(usuario["id"], agora + DURACAO_BLOQUEIO)
        raise MfaBloqueado()
    if do_token >= MAX_FALHAS_POR_TOKEN:
        raise TokenEsgotado()


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
