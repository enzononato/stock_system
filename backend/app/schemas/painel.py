"""Schemas do painel (F7): resumo de indicadores e alertas."""
from datetime import date, datetime
from decimal import Decimal
from typing import Optional

from pydantic import BaseModel


class EquipamentosPorUnidade(BaseModel):
    unidade: Optional[str] = None
    total: int
    por_status: dict[str, int]


class ResumoEquipamentos(BaseModel):
    total: int
    por_status: dict[str, int]
    por_unidade: list[EquipamentosPorUnidade]


class ResumoManutencao(BaseModel):
    equipamentos_em_manutencao: int
    ordens_em_andamento: int


class ResumoTermos(BaseModel):
    pendentes: int
    por_status: dict[str, int]
    sem_solicitacao: int


class LinhasPorUnidade(BaseModel):
    unidade_nome: str
    quantidade: int
    custo_total: Decimal


class ResumoLinhas(BaseModel):
    ativas: int
    custo_mensal: Decimal
    por_unidade: list[LinhasPorUnidade]


class ResumoPainel(BaseModel):
    gerado_em: datetime
    equipamentos: ResumoEquipamentos
    manutencao: ResumoManutencao
    termos: ResumoTermos
    linhas: ResumoLinhas
