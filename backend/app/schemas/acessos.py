"""Schemas do módulo de Acessos (F8).

O Patrimônio nunca pede nem guarda senha de sistema de terceiros: toda entrada herda
`EntradaSemSenha`, que recusa qualquer campo cujo nome lembre senha (qualquer grafia).
"""
import unicodedata
from datetime import datetime
from typing import Any, Literal, Optional

from pydantic import BaseModel, ConfigDict, field_validator, model_validator

from app.schemas.validators import blank_to_none

MSG_SEM_SENHA = "O Patrimônio não guarda senhas: remova o campo de senha e registre só o login."
PALAVRAS_SENHA = ("senha", "password", "passwd", "pwd")

TipoEvento = Literal["admissao", "desligamento", "mudanca_funcao"]
StatusEvento = Literal["aberto", "concluido", "cancelado"]
AcaoTarefa = Literal["criar", "desativar"]
StatusTarefa = Literal["pendente", "feito", "nao_se_aplica"]
StatusAcesso = Literal["ativo", "desativado"]

MSG_TEXTO_INVALIDO = "Valor inválido: informe um texto."


def normalizar_cargo(valor: Optional[str]) -> Optional[str]:
    """Aparado e com espaços internos repetidos reduzidos a um; vazio vira `None`."""
    if valor is None:
        return None
    texto = " ".join(valor.split())
    return texto or None


def chave_cargo(valor: Optional[str]) -> str:
    """Chave de comparação do cargo: sem acento, sem caixa e sem espaços repetidos."""
    texto = normalizar_cargo(valor) or ""
    sem_acento = "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    )
    return sem_acento.casefold()


class EntradaSemSenha(BaseModel):
    model_config = ConfigDict(extra="forbid")

    @model_validator(mode="before")
    @classmethod
    def _recusa_senha(cls, dados: Any):
        if isinstance(dados, dict):
            for chave in dados:
                minuscula = str(chave).lower()
                if any(palavra in minuscula for palavra in PALAVRAS_SENHA):
                    raise ValueError(MSG_SEM_SENHA)
        return dados


def _texto(valor: Any) -> Optional[str]:
    if valor is not None and not isinstance(valor, str):
        raise ValueError(MSG_TEXTO_INVALIDO)
    return blank_to_none(valor)


def _texto_max(valor: Any, limite: int, rotulo: str) -> Optional[str]:
    texto = _texto(valor)
    if texto is not None and len(texto) > limite:
        raise ValueError(f"{rotulo} muito longo (máximo {limite} caracteres).")
    return texto


def _cargo(valor: Any) -> Optional[str]:
    if valor is not None and not isinstance(valor, str):
        raise ValueError(MSG_TEXTO_INVALIDO)
    cargo = normalizar_cargo(valor)
    if cargo is not None and len(cargo) > 150:
        raise ValueError("Cargo muito longo (máximo 150 caracteres).")
    return cargo


def _nome_sistema(valor: Any) -> str:
    nome = _texto(valor)
    if nome is None:
        raise ValueError("Informe o nome do sistema.")
    if len(nome) > 100:
        raise ValueError("Nome do sistema muito longo (máximo 100 caracteres).")
    return nome


# ── Entrada ─────────────────────────────────────────────────────────────────

class SistemaCreate(EntradaSemSenha):
    nome: str
    descricao: Optional[str] = None

    @field_validator("nome", mode="before")
    @classmethod
    def _v_nome(cls, v):
        return _nome_sistema(v)

    @field_validator("descricao", mode="before")
    @classmethod
    def _v_descricao(cls, v):
        return _texto_max(v, 255, "Descrição")


class SistemaUpdate(EntradaSemSenha):
    """Campo ausente = sem mudança; `descricao: null` limpa."""

    nome: Optional[str] = None
    descricao: Optional[str] = None
    ativo: Optional[bool] = None

    @field_validator("nome", mode="before")
    @classmethod
    def _v_nome(cls, v):
        if v is None:
            return None  # o router trata como campo não enviado
        return _nome_sistema(v)

    @field_validator("descricao", mode="before")
    @classmethod
    def _v_descricao(cls, v):
        return _texto_max(v, 255, "Descrição")


class PacoteUpdate(EntradaSemSenha):
    cargo: str
    sistema_ids: list[int]

    @field_validator("cargo", mode="before")
    @classmethod
    def _v_cargo(cls, v):
        cargo = _cargo(v)
        if cargo is None:
            raise ValueError("Informe o cargo.")
        return cargo

    @field_validator("sistema_ids")
    @classmethod
    def _v_ids(cls, v):
        return list(dict.fromkeys(v))


class AcessoCreate(EntradaSemSenha):
    colaborador_id: int
    sistema_id: int
    login: str
    item_id: Optional[int] = None

    @field_validator("login", mode="before")
    @classmethod
    def _v_login(cls, v):
        login = _texto_max(v, 150, "Login")
        if login is None:
            raise ValueError("Informe o login.")
        return login


class AcessoUpdate(EntradaSemSenha):
    """Campo ausente = sem mudança; `item_id: null` limpa o equipamento."""

    login: Optional[str] = None
    item_id: Optional[int] = None

    @field_validator("login", mode="before")
    @classmethod
    def _v_login(cls, v):
        return _texto_max(v, 150, "Login")


class EventoCreate(EntradaSemSenha):
    colaborador_id: int
    tipo: TipoEvento
    cargo_anterior: Optional[str] = None
    cargo_novo: Optional[str] = None
    observacao: Optional[str] = None

    @field_validator("cargo_anterior", "cargo_novo", mode="before")
    @classmethod
    def _v_cargos(cls, v):
        return _cargo(v)

    @field_validator("observacao", mode="before")
    @classmethod
    def _v_observacao(cls, v):
        return _texto_max(v, 500, "Observação")


class TarefaResolver(EntradaSemSenha):
    status: Literal["feito", "nao_se_aplica", "manter"]
    login: Optional[str] = None
    item_id: Optional[int] = None
    observacao: Optional[str] = None

    @field_validator("login", mode="before")
    @classmethod
    def _v_login(cls, v):
        return _texto_max(v, 150, "Login")

    @field_validator("observacao", mode="before")
    @classmethod
    def _v_observacao(cls, v):
        return _texto_max(v, 500, "Observação")

    @model_validator(mode="after")
    def _manter_cabe_no_prefixo(self):
        # A tarefa grava "Acesso mantido: <motivo>" em VARCHAR(500).
        limite = 500 - len("Acesso mantido: ")
        if self.status == "manter" and self.observacao and len(self.observacao) > limite:
            raise ValueError(f"Motivo muito longo (máximo {limite} caracteres).")
        return self


class EventoCancelar(EntradaSemSenha):
    motivo_cancelamento: str

    @field_validator("motivo_cancelamento", mode="before")
    @classmethod
    def _v_motivo(cls, v):
        motivo = _texto(v)
        if motivo is None:
            raise ValueError("Informe o motivo do cancelamento.")
        if len(motivo) > 500:
            raise ValueError("Motivo muito longo (máximo 500 caracteres).")
        return motivo


# ── Respostas ───────────────────────────────────────────────────────────────

class SistemaResponse(BaseModel):
    id: int
    nome: str
    descricao: Optional[str] = None
    ativo: bool


class PacoteResponse(BaseModel):
    cargo: str
    sistemas: list[SistemaResponse]


class AcessoResponse(BaseModel):
    id: int
    colaborador_id: int
    colaborador_nome: str
    sistema_id: int
    sistema_nome: str
    login: str
    status: StatusAcesso
    item_id: Optional[int] = None
    item_descricao: Optional[str] = None
    criado_por: str
    criado_em: datetime
    desativado_por: Optional[str] = None
    desativado_em: Optional[datetime] = None


class EquipamentoResumo(BaseModel):
    id: int
    tipo: Optional[str] = None
    brand: Optional[str] = None
    model: Optional[str] = None
    identificador: Optional[str] = None
    status: str
    revenda: Optional[str] = None


class TarefaResponse(BaseModel):
    id: int
    evento_id: int
    sistema_id: int
    sistema_nome: str
    acao: AcaoTarefa
    status: StatusTarefa
    login: Optional[str] = None
    acesso_id: Optional[int] = None
    feito_por: Optional[str] = None
    feito_em: Optional[datetime] = None
    observacao: Optional[str] = None


class EventoResumo(BaseModel):
    id: int
    colaborador_id: int
    colaborador_nome: str
    tipo: TipoEvento
    cargo_anterior: Optional[str] = None
    cargo_novo: Optional[str] = None
    status: StatusEvento
    observacao: Optional[str] = None
    aberto_por: str
    aberto_em: datetime
    concluido_por: Optional[str] = None
    concluido_em: Optional[datetime] = None
    cancelado_por: Optional[str] = None
    cancelado_em: Optional[datetime] = None
    motivo_cancelamento: Optional[str] = None
    total_tarefas: int = 0
    tarefas_pendentes: int = 0


class EventoDetalhe(EventoResumo):
    colaborador_situacao: str
    tarefas: list[TarefaResponse]
    equipamentos: list[EquipamentoResumo]
    pode_concluir: bool
    pendencias: list[str]
