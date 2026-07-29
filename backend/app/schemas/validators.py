"""Funções utilitárias de validação de domínio, reaproveitadas pelos
field_validators dos schemas de items, loans e peripherals.

Mantidas fora dos modelos Pydantic para evitar duplicar a mesma lógica entre
variantes de Create/Update de um mesmo recurso.
"""
import ipaddress
import re
from datetime import datetime
from typing import Optional

_MAC_RE = re.compile(
    r"^([0-9A-Fa-f]{2})[:\-]?([0-9A-Fa-f]{2})[:\-]?([0-9A-Fa-f]{2})"
    r"[:\-]?([0-9A-Fa-f]{2})[:\-]?([0-9A-Fa-f]{2})[:\-]?([0-9A-Fa-f]{2})$"
)


def blank_to_none(value: Optional[str]) -> Optional[str]:
    """Converte string vazia/apenas espaços em None; preserva o restante já
    sem espaços nas pontas. Usado antes de qualquer validação de campo
    opcional, já que formulários costumam enviar "" em vez de omitir o campo."""
    if value is None:
        return None
    value = value.strip()
    return value or None


def validar_cpf(cpf: str) -> str:
    """Valida um CPF (11 dígitos, dígitos verificadores corretos, sem
    sequência repetida) e retorna apenas os dígitos, para gravação
    normalizada no banco."""
    digits = re.sub(r"\D", "", cpf or "")
    if len(digits) != 11:
        raise ValueError("CPF deve conter 11 dígitos.")
    if digits == digits[0] * 11:
        raise ValueError("CPF inválido.")

    def _digito_verificador(base: str) -> str:
        peso = len(base) + 1
        soma = sum(int(d) * (peso - i) for i, d in enumerate(base))
        resto = soma % 11
        return "0" if resto < 2 else str(11 - resto)

    digito1 = _digito_verificador(digits[:9])
    digito2 = _digito_verificador(digits[:9] + digito1)
    if digits[9] != digito1 or digits[10] != digito2:
        raise ValueError("CPF inválido (dígitos verificadores não conferem).")
    return digits


def validar_nota_fiscal(nota_fiscal: Optional[str]) -> Optional[str]:
    """Nota fiscal é opcional, mas quando preenchida deve ter exatamente 9
    dígitos numéricos (o frontend já anuncia isso no placeholder)."""
    nota_fiscal = blank_to_none(nota_fiscal)
    if nota_fiscal is None:
        return None
    if not nota_fiscal.isdigit() or len(nota_fiscal) != 9:
        raise ValueError("Nota fiscal deve conter exatamente 9 dígitos numéricos.")
    return nota_fiscal


def normalizar_mac(mac: Optional[str]) -> Optional[str]:
    """Aceita AA:BB:CC:DD:EE:FF, AA-BB-CC-DD-EE-FF ou AABBCCDDEEFF e
    normaliza sempre para o formato AA:BB:CC:DD:EE:FF (maiúsculo)."""
    mac = blank_to_none(mac)
    if mac is None:
        return None
    match = _MAC_RE.match(mac)
    if not match:
        raise ValueError(
            "Endereço MAC inválido. Use o formato AA:BB:CC:DD:EE:FF, "
            "AA-BB-CC-DD-EE-FF ou AABBCCDDEEFF."
        )
    return ":".join(g.upper() for g in match.groups())


def validar_ip(ip: Optional[str]) -> Optional[str]:
    """Valida um endereço IPv4 ou IPv6 usando ipaddress.ip_address."""
    ip = blank_to_none(ip)
    if ip is None:
        return None
    try:
        ipaddress.ip_address(ip)
    except ValueError:
        raise ValueError("Endereço IP inválido.")
    return ip


def validar_data_br(
    data: Optional[str], campo: str = "data", obrigatorio: bool = False
) -> Optional[str]:
    """Valida uma data no formato dd/mm/aaaa, retornando a própria string
    (o parse para datetime continua a cargo do router/camada de dados)."""
    data = blank_to_none(data)
    if data is None:
        if obrigatorio:
            raise ValueError(f"Informe {campo} no formato dd/mm/aaaa.")
        return None
    try:
        datetime.strptime(data, "%d/%m/%Y")
    except ValueError:
        raise ValueError(f"Formato de {campo} inválido. Use dd/mm/aaaa.")
    return data
