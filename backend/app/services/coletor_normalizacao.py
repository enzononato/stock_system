"""Normalização da coleta bruta da máquina nos valores do formulário de cadastro.

Funções puras (sem banco). Tudo que sai de `normalizar` já está limpo: caracteres de controle
e `<`/`>` removidos e texto cortado no tamanho da coluna de destino.
"""
import re
from dataclasses import dataclass, field
from typing import Literal, Optional

from app.schemas.coletor import Adaptador, CamposColeta, ColetaBruta, Disco

# Tamanho máximo de cada campo = coluna de destino em `items`.
LIMITES: dict[str, int] = {
    "brand": 100,
    "model": 100,
    "identificador": 100,
    "host": 100,
    "cpu": 100,
    "sistema": 100,
    "licenca": 100,
    "ram": 50,
    "storage": 50,
    "anydesk": 50,
    "dominio": 50,
    "endereco_fisico": 150,
}

_GENERICOS_TEXTO = {
    "system manufacturer",
    "to be filled by o.e.m.",
    "default string",
    "system product name",
    "system version",
}
_GENERICOS_SERIAL = {
    "to be filled by o.e.m.",
    "to be filled by oem",
    "to be filled by o.e.m",
    "default string",
    "default",
    "system serial number",
    "chassis serial number",
    "base board serial number",
    "not specified",
    "not applicable",
    "123456789",
    "1234567890",
    "none",
    "n/a",
    "oem",
}

_MAPA_MARCAS = {
    "DELL": "Dell",
    "LENOVO": "Lenovo",
    "HP": "HP",
    "HEWLETT-PACKARD": "HP",
    "ASUSTEK": "ASUS",
    "ACER": "Acer",
    "SAMSUNG ELECTRONICS": "Samsung",
    "POSITIVO": "Positivo",
}

_SUFIXO_MARCA = re.compile(
    r"(?:^|[\s,]+)(?:Inc\.?|Ltda\.?|Tecnologia|Corporation|Corp\.?|Co\.?,?\s*Ltd\.?|Computer)[\s.,]*$",
    re.IGNORECASE,
)

_TAMANHOS_COMERCIAIS = (120, 128, 240, 250, 256, 480, 500, 512, 960, 1000, 1920, 2000, 3840, 4000)
_TOLERANCIA_TAMANHO = 0.12

# MACs de preenchimento (adaptador sem endereço real) não identificam máquina nenhuma.
_MACS_GENERICOS = {"000000000000", "FFFFFFFFFFFF"}

_CHASSI_NOTEBOOK = {8, 9, 10, 14, 30, 31, 32}
_CHASSI_DESKTOP = {3, 4, 5, 6, 7, 13, 15, 16, 35, 36}

TipoSugerido = Literal["Notebook", "Desktop"]


@dataclass
class ColetaNormalizada:
    campos: CamposColeta = field(default_factory=dict)
    tipo_sugerido: Optional[TipoSugerido] = None
    # MACs (normalizados AABBCCDDEEFF) dos adaptadores não USB, para a busca de cadastro.
    macs: list[str] = field(default_factory=list)


def _limpar(valor: Optional[str], limite: int) -> str:
    """Remove caracteres de controle e `<`/`>`, apara e corta no tamanho da coluna."""
    if not valor:
        return ""
    texto = "".join(c for c in valor if c.isprintable() and c not in "<>")
    return texto.strip()[:limite].rstrip()


def _generico(valor: str) -> bool:
    return " ".join(valor.lower().split()) in _GENERICOS_TEXTO


def marca(fabricante: Optional[str]) -> str:
    nome = " ".join((fabricante or "").split())
    if not nome or _generico(nome):
        return ""
    while True:
        sem_sufixo = _SUFIXO_MARCA.sub("", nome).strip(" ,.")
        if sem_sufixo == nome:
            break
        nome = sem_sufixo
    if not nome:
        return ""
    maiusculo = nome.upper()
    for prefixo, canonico in _MAPA_MARCAS.items():
        if maiusculo.startswith(prefixo):
            return canonico
    return nome.title()


def _modelo(bruta: ColetaBruta) -> str:
    fabricante = marca(bruta.fabricante)
    modelo = " ".join((bruta.modelo or "").split())
    if fabricante == "Lenovo":
        comercial = " ".join((bruta.versao_produto or "").split())
        if comercial and not _generico(comercial):
            modelo = comercial
    if _generico(modelo):
        return ""
    # "HP ProBook 440 G8" -> "ProBook 440 G8": a marca já vai no campo próprio.
    prefixo = f"{fabricante} ".lower()
    if fabricante and modelo.lower().startswith(prefixo) and modelo[len(prefixo):].strip():
        modelo = modelo[len(prefixo):].strip()
    return modelo


def cpu_curta(nome: Optional[str]) -> str:
    texto = nome or ""
    texto = texto.split("@", 1)[0]
    texto = re.sub(r"\((?:R|TM)\)", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\bwith\s+Radeon\b.*$", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\b\d+(?:st|nd|rd|th)\s+Gen\b", " ", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\b(?:CPU|Processor)\b", " ", texto, flags=re.IGNORECASE)
    texto = re.sub(r"^\s*(?:Intel|AMD)\b", " ", texto, flags=re.IGNORECASE)
    # `Core i5-1345U` -> `i5-1345U`; `Core Ultra 5 125U` -> `Ultra 5 125U`.
    texto = re.sub(r"\bCore\s+(?=i\d|Ultra\b)", "", texto, flags=re.IGNORECASE)
    return " ".join(texto.split())


def ram(total_bytes: Optional[int]) -> str:
    if not total_bytes or total_bytes < 0:
        return ""
    gb = round(total_bytes / 2**30)
    return f"{gb}GB" if gb > 0 else ""


def _tamanho_gb(total_bytes: int) -> int:
    """GB comercial mais próximo (dentro da tolerância); fora da lista, GB inteiro.

    O Windows informa o tamanho em bytes decimais (512110190592 = disco de "512 GB").
    """
    gb = total_bytes / 1e9
    mais_proximo = min(_TAMANHOS_COMERCIAIS, key=lambda c: abs(c - gb))
    if abs(mais_proximo - gb) / mais_proximo <= _TOLERANCIA_TAMANHO:
        return mais_proximo
    return round(gb)


def _tipo_disco(disco: Disco) -> str:
    tipo = (disco.tipo or "").strip().upper()
    if tipo in ("SSD", "HDD"):
        return tipo
    if tipo in ("", "UNSPECIFIED") and (disco.barramento or "").strip().upper() == "NVME":
        return "SSD"
    return ""


def armazenamento(discos: list[Disco]) -> str:
    partes: list[str] = []
    for disco in discos:
        if (disco.barramento or "").strip().upper() == "USB":
            continue
        if not disco.bytes:
            continue
        tamanho = _tamanho_gb(disco.bytes)
        if tamanho <= 0:
            continue
        tipo = _tipo_disco(disco)
        partes.append(f"{tamanho} {tipo}" if tipo else str(tamanho))
    return " + ".join(partes)


def sistema(caption: Optional[str]) -> str:
    texto = " ".join((caption or "").split())
    return re.sub(r"^Microsoft\s+", "", texto, flags=re.IGNORECASE)


def serial(valor: Optional[str]) -> str:
    texto = (valor or "").strip()
    # Vazio, só zeros/espaços/traços/pontos, ou só "x" (ex.: "xxxxxxxx").
    if not texto or re.fullmatch(r"[0\s.\-]*", texto) or re.fullmatch(r"[xX]+", texto):
        return ""
    if " ".join(texto.lower().split()) in _GENERICOS_SERIAL:
        return ""
    return texto


def _mac_hex(mac: Optional[str]) -> str:
    """`AABBCCDDEEFF` ou vazio quando o tamanho não é de um MAC ou ele é genérico."""
    hexa = re.sub(r"[^0-9A-Fa-f]", "", mac or "").upper()
    if len(hexa) != 12 or hexa in _MACS_GENERICOS:
        return ""
    return hexa


def _formatar_mac(hexa: str) -> str:
    return ":".join(hexa[i : i + 2] for i in range(0, 12, 2))


def mac_gravado(adaptadores: list[Adaptador]) -> str:
    validos = [(a, _mac_hex(a.mac)) for a in adaptadores]
    validos = [(a, m) for a, m in validos if m]
    regras = (
        lambda a: a.ativo and a.gateway,
        lambda a: a.ativo and a.tipo == "ethernet",
        lambda a: a.ativo and a.tipo == "wifi",
    )
    # Adaptadores internos primeiro: o MAC de uma dock/adaptador USB acompanha o acessório.
    # Só se nenhum interno servir é que um USB é gravado.
    for grupo in ([v for v in validos if not v[0].usb], [v for v in validos if v[0].usb]):
        for regra in regras:
            for adaptador, hexa in grupo:
                if regra(adaptador):
                    return _formatar_mac(hexa)
    return ""


def tipo_pelo_chassi(chassi: list[int]) -> Optional[TipoSugerido]:
    for codigo in chassi:
        if codigo in _CHASSI_NOTEBOOK:
            return "Notebook"
        if codigo in _CHASSI_DESKTOP:
            return "Desktop"
    return None


def _macs(adaptadores: list[Adaptador]) -> list[str]:
    vistos: list[str] = []
    for adaptador in adaptadores:
        if adaptador.usb:
            continue
        hexa = _mac_hex(adaptador.mac)
        if hexa and hexa not in vistos:
            vistos.append(hexa)
    return vistos


def normalizar(bruta: ColetaBruta) -> ColetaNormalizada:
    dominio = {True: "Sim", False: "Não"}.get(bruta.no_dominio, "")
    anydesk = re.sub(r"\D", "", bruta.anydesk or "")
    candidatos = {
        "brand": marca(bruta.fabricante),
        "model": _modelo(bruta),
        "identificador": serial(bruta.serial),
        "dominio": dominio,
        "host": (bruta.host or "").strip(),
        "endereco_fisico": mac_gravado(bruta.adaptadores),
        "cpu": cpu_curta(bruta.cpu),
        "ram": ram(bruta.ram_bytes),
        "storage": armazenamento(bruta.discos),
        "sistema": sistema(bruta.sistema),
        "licenca": (bruta.product_id or "").strip(),
        "anydesk": anydesk,
    }
    campos: CamposColeta = {}
    for chave, valor in candidatos.items():
        limpo = _limpar(valor, LIMITES[chave])
        if limpo:
            campos[chave] = limpo
    return ColetaNormalizada(
        campos=campos,
        tipo_sugerido=tipo_pelo_chassi(bruta.chassi),
        macs=_macs(bruta.adaptadores),
    )
