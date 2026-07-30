"""Schemas Pydantic da funcionalidade de Unidades.

As validações de domínio (CNPJ, CEP, UF) seguem o mesmo estilo de
`app/schemas/validators.py` (dígitos verificadores para CPF), mas ficam
definidas aqui — `validators.py` não é de posse deste módulo e não deve ser
alterado. `blank_to_none` é importado de lá por ser um utilitário puramente
genérico, sem lógica de domínio.
"""
import re
from typing import Optional

from pydantic import BaseModel, field_validator

from app.schemas.validators import blank_to_none

# Unidades federativas válidas (26 estados + Distrito Federal).
_UF_VALIDAS = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS",
    "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC",
    "SP", "SE", "TO",
}


def _digito_verificador_cnpj(base: str) -> str:
    """Calcula um dígito verificador de CNPJ a partir da base de dígitos.

    Os pesos usados pela Receita Federal começam em 2 (dígito mais à
    direita) e sobem até 9, reiniciando em 2 — o mesmo para os dois dígitos
    verificadores, apenas com bases de tamanhos diferentes (12 e 13 dígitos).
    Gerar os pesos de trás para frente evita hardcoded duas listas distintas.
    """
    pesos = []
    peso = 2
    for _ in range(len(base)):
        pesos.append(peso)
        peso = peso + 1 if peso < 9 else 2
    pesos.reverse()
    soma = sum(int(d) * p for d, p in zip(base, pesos))
    resto = soma % 11
    return "0" if resto < 2 else str(11 - resto)


def validar_cnpj(cnpj: Optional[str], *, obrigatorio: bool) -> Optional[str]:
    """Valida um CNPJ (14 dígitos, dígitos verificadores corretos) e retorna
    normalizado no formato `00.000.000/0000-00`.

    `obrigatorio=True` (UnidadeCreate) recusa valor vazio; `obrigatorio=False`
    (UnidadeUpdate) trata vazio/None como "não alterar este campo".
    """
    cnpj = blank_to_none(cnpj)
    if cnpj is None:
        if obrigatorio:
            raise ValueError("CNPJ é obrigatório.")
        return None

    digits = re.sub(r"\D", "", cnpj)
    if len(digits) != 14:
        raise ValueError("CNPJ deve conter 14 dígitos.")
    if digits == digits[0] * 14:
        raise ValueError("CNPJ inválido.")

    digito1 = _digito_verificador_cnpj(digits[:12])
    digito2 = _digito_verificador_cnpj(digits[:12] + digito1)
    if digits[12] != digito1 or digits[13] != digito2:
        raise ValueError("CNPJ inválido (dígitos verificadores não conferem).")

    return f"{digits[0:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:14]}"


def validar_cep(cep: Optional[str]) -> Optional[str]:
    """CEP é sempre opcional; quando preenchido, normaliza para `00000-000`."""
    cep = blank_to_none(cep)
    if cep is None:
        return None
    digits = re.sub(r"\D", "", cep)
    if len(digits) != 8:
        raise ValueError("CEP deve conter 8 dígitos (formato 00000-000).")
    return f"{digits[0:5]}-{digits[5:8]}"


def validar_uf(uf: Optional[str]) -> Optional[str]:
    """UF é sempre opcional; quando preenchida, deve ser uma sigla válida."""
    uf = blank_to_none(uf)
    if uf is None:
        return None
    uf = uf.upper()
    if uf not in _UF_VALIDAS:
        raise ValueError("UF inválida.")
    return uf


def validar_nome(nome: Optional[str], *, obrigatorio: bool) -> Optional[str]:
    nome = blank_to_none(nome)
    if nome is None:
        if obrigatorio:
            raise ValueError("Nome da unidade é obrigatório.")
        return None
    return nome


def validar_razao_social(razao_social: Optional[str], *, obrigatorio: bool) -> Optional[str]:
    razao_social = blank_to_none(razao_social)
    if razao_social is None:
        if obrigatorio:
            raise ValueError("Razão social é obrigatória.")
        return None
    return razao_social


class UnidadeCreate(BaseModel):
    nome: str
    razao_social: str
    cnpj: str
    endereco: Optional[str] = None
    cep: Optional[str] = None
    cidade: Optional[str] = None
    uf: Optional[str] = None

    @field_validator("nome")
    @classmethod
    def _v_nome(cls, v):
        return validar_nome(v, obrigatorio=True)

    @field_validator("razao_social")
    @classmethod
    def _v_razao_social(cls, v):
        return validar_razao_social(v, obrigatorio=True)

    @field_validator("cnpj")
    @classmethod
    def _v_cnpj(cls, v):
        return validar_cnpj(v, obrigatorio=True)

    @field_validator("cep")
    @classmethod
    def _v_cep(cls, v):
        return validar_cep(v)

    @field_validator("uf")
    @classmethod
    def _v_uf(cls, v):
        return validar_uf(v)

    @field_validator("endereco", "cidade")
    @classmethod
    def _v_opcionais(cls, v):
        return blank_to_none(v)


class UnidadeUpdate(BaseModel):
    nome: Optional[str] = None
    razao_social: Optional[str] = None
    cnpj: Optional[str] = None
    endereco: Optional[str] = None
    cep: Optional[str] = None
    cidade: Optional[str] = None
    uf: Optional[str] = None

    @field_validator("nome")
    @classmethod
    def _v_nome(cls, v):
        return validar_nome(v, obrigatorio=False)

    @field_validator("razao_social")
    @classmethod
    def _v_razao_social(cls, v):
        return validar_razao_social(v, obrigatorio=False)

    @field_validator("cnpj")
    @classmethod
    def _v_cnpj(cls, v):
        return validar_cnpj(v, obrigatorio=False)

    @field_validator("cep")
    @classmethod
    def _v_cep(cls, v):
        return validar_cep(v)

    @field_validator("uf")
    @classmethod
    def _v_uf(cls, v):
        return validar_uf(v)

    @field_validator("endereco", "cidade")
    @classmethod
    def _v_opcionais(cls, v):
        return blank_to_none(v)


class UnidadeResponse(BaseModel):
    id: int
    nome: str
    razao_social: str
    cnpj: str
    endereco: Optional[str] = None
    cep: Optional[str] = None
    cidade: Optional[str] = None
    uf: Optional[str] = None
    is_active: bool

    class Config:
        from_attributes = True


class IndicadoresResponse(BaseModel):
    termos_emitidos: int
    termos_confirmados: int
    devolucoes_concluidas: int
    itens_total: int
    itens_por_status: dict[str, int]
    emprestimos_ativos: int
    perifericos_vinculados: int
