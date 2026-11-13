"""Orquestra o 2FA: banco (`MfaDBManager`), cifra/TOTP (`mfa`) e auditoria.

Nunca grava segredo, hash nem códigos na auditoria.
"""
from app.db.mfa_db import MfaDBManager
from app.services import audit
from app.services.audit import AuditContext
from app.services.mfa import (
    CodigoInvalido,
    decifrar_segredo,
    hash_codigo_recuperacao,
    parece_codigo_totp,
    verificar_codigo_totp,
)


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
