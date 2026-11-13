"""F2 Task 4: serviço de 2FA (cifra, TOTP e códigos de recuperação). Sem banco."""
import re
from datetime import datetime, timedelta, timezone

import pyotp
import pytest
from cryptography.fernet import Fernet

from app.core import config
from app.core.security import sha256_hex
from app.services import mfa
from app.services.mfa import (
    MSG_MFA_INDISPONIVEL,
    MSG_SEGREDO_ILEGIVEL,
    MfaIndisponivel,
)

pytestmark = pytest.mark.unit


def _chave(monkeypatch, valor: str) -> None:
    monkeypatch.setattr(config.settings, "TOTP_ENCRYPTION_KEY", valor)


def test_mfa_configurado(monkeypatch):
    _chave(monkeypatch, "")
    assert mfa.mfa_configurado() is False
    _chave(monkeypatch, "nao-e-fernet")
    assert mfa.mfa_configurado() is False
    _chave(monkeypatch, Fernet.generate_key().decode())
    assert mfa.mfa_configurado() is True


def test_cifrar_e_decifrar(monkeypatch):
    _chave(monkeypatch, Fernet.generate_key().decode())
    cifrado = mfa.cifrar_segredo("JBSWY3DPEHPK3PXP")
    assert cifrado != "JBSWY3DPEHPK3PXP"
    assert mfa.decifrar_segredo(cifrado) == "JBSWY3DPEHPK3PXP"


def test_sem_chave_levanta_indisponivel(monkeypatch):
    _chave(monkeypatch, "")
    with pytest.raises(MfaIndisponivel) as exc:
        mfa.cifrar_segredo("JBSWY3DPEHPK3PXP")
    assert exc.value.status_code == 503
    assert str(exc.value) == MSG_MFA_INDISPONIVEL


def test_chave_trocada_levanta_segredo_ilegivel(monkeypatch):
    _chave(monkeypatch, Fernet.generate_key().decode())
    cifrado = mfa.cifrar_segredo("JBSWY3DPEHPK3PXP")
    _chave(monkeypatch, Fernet.generate_key().decode())
    with pytest.raises(MfaIndisponivel) as exc:
        mfa.decifrar_segredo(cifrado)
    assert exc.value.status_code == 503
    assert str(exc.value) == MSG_SEGREDO_ILEGIVEL


def test_gerar_segredo_base32():
    s = mfa.gerar_segredo()
    assert re.fullmatch(r"[A-Z2-7]{32}", s)


def test_uri_provisionamento():
    s = mfa.gerar_segredo()
    uri = mfa.uri_provisionamento(s, "gestor.teste")
    assert uri.startswith("otpauth://totp/")
    assert f"secret={s}" in uri
    assert "issuer=Revalle" in uri


AGORA = datetime(2026, 11, 12, 12, 0, 0, tzinfo=timezone.utc)


def test_verificar_codigo_totp_janela():
    s = mfa.gerar_segredo()
    totp = pyotp.TOTP(s)
    p = totp.timecode(AGORA)
    assert mfa.verificar_codigo_totp(s, totp.at(AGORA), None, AGORA) == p
    assert mfa.verificar_codigo_totp(s, totp.at(AGORA - timedelta(seconds=30)), None, AGORA) == p - 1
    assert mfa.verificar_codigo_totp(s, totp.at(AGORA + timedelta(seconds=30)), None, AGORA) == p + 1
    assert mfa.verificar_codigo_totp(s, totp.at(AGORA + timedelta(seconds=90)), None, AGORA) is None


def test_verificar_codigo_totp_anti_replay():
    s = mfa.gerar_segredo()
    totp = pyotp.TOTP(s)
    p = totp.timecode(AGORA)
    c = totp.at(AGORA)
    assert mfa.verificar_codigo_totp(s, c, p, AGORA) is None
    assert mfa.verificar_codigo_totp(s, c, p - 1, AGORA) == p


def test_verificar_codigo_totp_aceita_espacos():
    s = mfa.gerar_segredo()
    totp = pyotp.TOTP(s)
    p = totp.timecode(AGORA)
    c = totp.at(AGORA)
    assert mfa.verificar_codigo_totp(s, f"{c[:3]} {c[3:]}", None, AGORA) == p
    assert mfa.verificar_codigo_totp(s, "12345", None, AGORA) is None
    assert mfa.verificar_codigo_totp(s, "abcdef", None, AGORA) is None


def test_codigos_recuperacao():
    codigos = mfa.gerar_codigos_recuperacao()
    assert len(codigos) == 10
    assert len(set(codigos)) == 10
    assert all(re.fullmatch(r"[a-z2-7]{4}(-[a-z2-7]{4}){3}", c) for c in codigos)


def test_hash_codigo_normalizado():
    esperado = sha256_hex("abcdefghijklmnop")
    for variante in ("ABCD-EFGH-IJKL-MNOP", "abcdefghijklmnop", " abcd efgh ijkl mnop "):
        assert mfa.hash_codigo_recuperacao(variante) == esperado
    assert len(esperado) == 64


def test_parece_codigo_totp():
    assert mfa.parece_codigo_totp("123456") is True
    assert mfa.parece_codigo_totp("123 456") is True
    assert mfa.parece_codigo_totp("abcd-efgh-ijkl-mnop") is False
