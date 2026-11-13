"""Serviço de 2FA (TOTP): cifra do segredo, verificação do código e códigos de recuperação.

Funções puras (sem banco). A chave de cifra é lida de `config.settings` no momento
da chamada, para os testes poderem trocá-la (o `test_auth` recarrega `app.core.config`).
"""
import hmac
import re
import secrets
from datetime import datetime, timezone

import pyotp
from cryptography.fernet import Fernet, InvalidToken

from app.core import config
from app.core.security import sha256_hex

ISSUER = "Revalle Patrimônio"
QTD_CODIGOS_RECUPERACAO = 10
_ALFABETO_RECUPERACAO = "abcdefghijklmnopqrstuvwxyz234567"

MSG_MFA_INDISPONIVEL = (
    "A verificação em duas etapas não está disponível: o servidor está sem a chave "
    "TOTP_ENCRYPTION_KEY. Avise o administrador."
)
MSG_SEGREDO_ILEGIVEL = (
    "Não foi possível validar o código do autenticador com a chave atual do servidor. "
    "Use um código de recuperação ou peça a um Gestor para zerar o seu 2FA."
)
MSG_CODIGO_INVALIDO = "Código inválido."
MSG_MFA_JA_ATIVO = "A verificação em duas etapas já está ativa."
MSG_CADASTRO_NAO_INICIADO = "Inicie o cadastro do autenticador antes de confirmar."


class MfaErro(Exception):
    """Erro de negócio do 2FA; `status_code` é o da resposta HTTP."""

    status_code = 400
    mensagem_padrao = ""

    def __init__(self, mensagem: str | None = None):
        super().__init__(mensagem or self.mensagem_padrao)


class MfaIndisponivel(MfaErro):
    status_code = 503
    mensagem_padrao = MSG_MFA_INDISPONIVEL


class CodigoInvalido(MfaErro):
    status_code = 400
    mensagem_padrao = MSG_CODIGO_INVALIDO


class MfaJaAtivo(MfaErro):
    status_code = 409
    mensagem_padrao = MSG_MFA_JA_ATIVO


class CadastroNaoIniciado(MfaErro):
    status_code = 400
    mensagem_padrao = MSG_CADASTRO_NAO_INICIADO


def _fernet() -> Fernet:
    chave = (config.settings.TOTP_ENCRYPTION_KEY or "").strip()
    if not chave:
        raise MfaIndisponivel()
    try:
        return Fernet(chave.encode("utf-8"))
    except (ValueError, TypeError):
        raise MfaIndisponivel()


def mfa_configurado() -> bool:
    """A chave de cifra existe e o Fernet a aceita."""
    try:
        _fernet()
    except MfaIndisponivel:
        return False
    return True


def cifrar_segredo(segredo: str) -> str:
    return _fernet().encrypt(segredo.encode("utf-8")).decode("utf-8")


def decifrar_segredo(cifrado: str) -> str:
    try:
        return _fernet().decrypt(cifrado.encode("utf-8")).decode("utf-8")
    except InvalidToken:
        raise MfaIndisponivel(MSG_SEGREDO_ILEGIVEL)


def gerar_segredo() -> str:
    return pyotp.random_base32()


def uri_provisionamento(segredo: str, username: str) -> str:
    return pyotp.TOTP(segredo).provisioning_uri(name=username, issuer_name=ISSUER)


def parece_codigo_totp(codigo: str) -> bool:
    """6 dígitos, ignorando espaços."""
    return re.fullmatch(r"[0-9]{6}", re.sub(r"\s", "", codigo or "")) is not None


def verificar_codigo_totp(
    segredo: str, codigo: str, ultimo_passo: int | None, agora: datetime | None = None
) -> int | None:
    """Passo (janela de 30 s) em que o código vale, ou None.

    Aceita ±1 passo, mas só um passo maior que o último já aceito (anti-replay)."""
    codigo = re.sub(r"\s", "", codigo or "")
    # [0-9] (não \d): dígitos Unicode quebrariam o compare_digest (só ASCII).
    if not re.fullmatch(r"[0-9]{6}", codigo):
        return None
    totp = pyotp.TOTP(segredo)
    atual = totp.timecode(agora or datetime.now(timezone.utc))
    for passo in (atual - 1, atual, atual + 1):
        if ultimo_passo is not None and passo <= ultimo_passo:
            continue
        if hmac.compare_digest(totp.generate_otp(passo), codigo):
            return passo
    return None


def gerar_codigos_recuperacao(quantidade: int = QTD_CODIGOS_RECUPERACAO) -> list[str]:
    """`quantidade` códigos `xxxx-xxxx-xxxx-xxxx` (16 caracteres, 80 bits), sem repetição."""
    codigos: list[str] = []
    while len(codigos) < quantidade:
        bruto = "".join(secrets.choice(_ALFABETO_RECUPERACAO) for _ in range(16))
        codigo = "-".join(bruto[i : i + 4] for i in range(0, 16, 4))
        if codigo not in codigos:
            codigos.append(codigo)
    return codigos


def normalizar_codigo_recuperacao(codigo: str) -> str:
    return re.sub(r"[-\s]", "", codigo or "").lower()


def hash_codigo_recuperacao(codigo: str) -> str:
    return sha256_hex(normalizar_codigo_recuperacao(codigo))
