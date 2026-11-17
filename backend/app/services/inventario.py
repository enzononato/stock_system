"""Regras puras do inventário físico: leitura do código, chave de duplicidade e
conferência (unidade, local, usuário). Sem banco e sem HTTP.
"""
import re
import unicodedata
from dataclasses import dataclass
from typing import Literal, Optional

Resultado = Literal["ok", "local_errado", "usuario_diferente", "outra_unidade", "nao_cadastrado"]
Categoria = Literal[
    "ok", "local_errado", "usuario_diferente", "outra_unidade", "nao_cadastrado", "nao_encontrado"
]
Divergencia = Literal["baixado", "sem_unidade", "outra_unidade", "usuario", "local"]

ROTULOS_CATEGORIA = {
    "ok": "Conferido",
    "local_errado": "Local errado",
    "usuario_diferente": "Usuário diferente",
    "outra_unidade": "De outra unidade",
    "nao_cadastrado": "Não cadastrado",
    "nao_encontrado": "Não encontrado",
}
ROTULO_AINDA_NAO_LIDO = "Ainda não lido"
ORDEM_CATEGORIAS = (
    "nao_encontrado", "nao_cadastrado", "outra_unidade", "usuario_diferente", "local_errado", "ok",
)

# `;` cobre o leitor USB em layout americano digitando num Windows ABNT2, que
# transforma "https://host/e/12" em "httpsÇ;;host;e;12".
_URL_ETIQUETA = re.compile(r"[/;]e[/;](\d{1,10})[/;]?\s*$", re.I)
_NUMERO_DIGITADO = re.compile(r"(?:#|n[º°o.]?)?\s*(\d{1,10})", re.I)


class CodigoInvalido(ValueError):
    pass


@dataclass(frozen=True)
class Alvo:
    tipo: Literal["id", "identificador"]
    valor: "int | str"


def resolver_codigo(texto: str) -> Alvo:
    t = (texto or "").strip()
    if not t:
        raise CodigoInvalido("Informe o código lido.")
    m = _URL_ETIQUETA.search(t)
    if m:
        return Alvo("id", int(m.group(1)))
    m = _NUMERO_DIGITADO.fullmatch(t)
    if m:
        return Alvo("id", int(m.group(1)))
    return Alvo("identificador", " ".join(t.split()).upper())


def chave_da_leitura(alvo: Alvo, item_id: Optional[int]) -> str:
    if item_id is not None:
        return f"item:{item_id}"
    prefixo = "id" if alvo.tipo == "id" else "codigo"
    return f"{prefixo}:{alvo.valor}"[:120]


def normalizar_texto(v: Optional[str]) -> str:
    """Sem acento nem caixa e com espaços colapsados, como a collation do MySQL."""
    if v is None:
        return ""
    decomposto = unicodedata.normalize("NFD", str(v))
    sem_marcas = "".join(c for c in decomposto if not unicodedata.combining(c))
    return " ".join(sem_marcas.casefold().split())


@dataclass(frozen=True)
class Esperado:
    item_id: int
    ativo: bool
    revenda: Optional[str]
    local_id: Optional[int]
    colaborador_id: Optional[int]
    usuario_nome: Optional[str]


@dataclass(frozen=True)
class Encontrado:
    local_id: Optional[int]
    usuario_conferido: bool
    colaborador_id: Optional[int]
    colaborador_nome: Optional[str]


def _usuario_diverge(esperado: Esperado, encontrado: Encontrado) -> bool:
    achado_id = encontrado.colaborador_id
    achado_nome = normalizar_texto(encontrado.colaborador_nome)
    ninguem_achado = achado_id is None and not achado_nome
    if esperado.colaborador_id is not None:
        return achado_id != esperado.colaborador_id
    if normalizar_texto(esperado.usuario_nome):
        if ninguem_achado:
            return True
        return achado_nome != normalizar_texto(esperado.usuario_nome)
    return not ninguem_achado


def classificar(
    esperado: Optional[Esperado], unidade_nome: str, encontrado: Encontrado
) -> "tuple[Resultado, list[Divergencia]]":
    if esperado is None:
        return "nao_cadastrado", []
    if not esperado.ativo:
        return "nao_cadastrado", ["baixado"]

    divergencias: list[Divergencia] = []
    if not normalizar_texto(esperado.revenda):
        divergencias.append("sem_unidade")
    elif normalizar_texto(esperado.revenda) != normalizar_texto(unidade_nome):
        divergencias.append("outra_unidade")
    fora_da_unidade = bool(divergencias)
    if encontrado.usuario_conferido and _usuario_diverge(esperado, encontrado):
        divergencias.append("usuario")
    if (
        not fora_da_unidade
        and esperado.local_id is not None
        and encontrado.local_id is not None
        and esperado.local_id != encontrado.local_id
    ):
        divergencias.append("local")

    if fora_da_unidade:
        return "outra_unidade", divergencias
    if "usuario" in divergencias:
        return "usuario_diferente", divergencias
    if "local" in divergencias:
        return "local_errado", divergencias
    return "ok", divergencias


def montar_detalhe(
    divergencias: "list[Divergencia]",
    *,
    revenda: Optional[str],
    local_esperado_nome: Optional[str],
    local_encontrado_nome: Optional[str],
    usuario_esperado: Optional[str],
    usuario_encontrado: Optional[str],
) -> Optional[str]:
    frases = []
    for d in divergencias:
        if d == "baixado":
            frases.append("Equipamento baixado no sistema.")
        elif d == "sem_unidade":
            frases.append("Equipamento sem unidade cadastrada.")
        elif d == "outra_unidade":
            frases.append(f"Cadastrado na unidade {(revenda or '').upper()}.")
        elif d == "usuario":
            frases.append(
                f"Usuário cadastrado: {usuario_esperado or 'nenhum'} · "
                f"encontrado com: {usuario_encontrado or 'ninguém'}."
            )
        elif d == "local":
            frases.append(
                f"Local cadastrado: {local_esperado_nome} · encontrado em: {local_encontrado_nome}."
            )
    return " ".join(frases) or None
