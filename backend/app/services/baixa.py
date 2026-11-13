"""Regras da baixa de equipamento (F4): apagamento de dados (LGPD) e bloqueios.

Para os tipos com armazenamento de dados pessoais a baixa exige registrar se os
dados foram apagados, o método, quem fez e quando (comprovante opcional). "Não
se aplica" só vale para equipamento fora de posse (roubo ou perda); para os
demais tipos o registro é sempre `nao_aplica`.
"""
from dataclasses import dataclass
from datetime import date, datetime
from typing import Optional

from app.schemas.validators import blank_to_none

TIPOS_EXIGEM_APAGAMENTO = ("Notebook", "Desktop", "Celular", "Tablet", "HD")
METODOS_APAGAMENTO = {
    "formatacao_segura": "Formatação segura",
    "reset_fabrica": "Reset de fábrica",
    "destruicao_fisica": "Destruição física",
}
# Motivos de baixa (chaves de REMOVAL_REASONS) em que o equipamento não está mais com a empresa.
MOTIVOS_SEM_POSSE = ("Roubo", "Perda")


@dataclass(frozen=True)
class Apagamento:
    dados_apagados: str  # "sim" | "nao_aplica"
    metodo: str  # chave de METODOS_APAGAMENTO ou "nao_aplica"
    apagado_por: Optional[str]
    apagado_em: Optional[date]


NAO_APLICA = Apagamento("nao_aplica", "nao_aplica", None, None)


def exige_apagamento(tipo: Optional[str]) -> bool:
    return tipo in TIPOS_EXIGEM_APAGAMENTO


def validar_apagamento(
    *,
    tipo: Optional[str],
    motivo: str,
    dados_apagados: Optional[str],
    metodo: Optional[str],
    apagado_por: Optional[str],
    apagado_em: Optional[str],
    hoje: Optional[date] = None,
) -> Apagamento:
    """Valida o registro de apagamento da baixa; levanta `ValueError(mensagem)`."""
    if not exige_apagamento(tipo):
        return NAO_APLICA
    dados_apagados = blank_to_none(dados_apagados)
    if dados_apagados is None:
        raise ValueError("Informe se os dados do equipamento foram apagados.")
    if dados_apagados not in ("sim", "nao_aplica"):
        raise ValueError("Valor inválido para dados apagados.")
    if dados_apagados == "nao_aplica":
        if motivo in MOTIVOS_SEM_POSSE:
            return NAO_APLICA
        raise ValueError(
            f'Para {tipo}, a baixa exige o apagamento dos dados (LGPD). '
            '"Não se aplica" só vale para Roubo ou Perda.'
        )

    metodo = blank_to_none(metodo)
    if metodo not in METODOS_APAGAMENTO:
        raise ValueError("Informe o método de apagamento dos dados.")
    apagado_por = blank_to_none(apagado_por)
    if apagado_por is None:
        raise ValueError("Informe quem apagou os dados.")
    if len(apagado_por) > 100:
        raise ValueError("Quem apagou: máximo de 100 caracteres.")
    texto_data = blank_to_none(apagado_em)
    try:
        data = datetime.strptime(texto_data or "", "%d/%m/%Y").date()
    except ValueError:
        raise ValueError("Informe a data do apagamento no formato dd/mm/aaaa.")
    if data > (hoje or date.today()):
        raise ValueError("A data do apagamento não pode ser no futuro.")
    return Apagamento("sim", metodo, apagado_por, data)


def motivo_bloqueio_baixa(item: dict, tem_manutencao_em_andamento: bool) -> Optional[str]:
    """Por que o item não pode ser baixado agora, ou `None` se pode."""
    status = item.get("status")
    if status == "Disponível":
        return None
    if status == "Em Manutenção":
        if tem_manutencao_em_andamento:
            return "Conclua ou cancele a manutenção em andamento antes da baixa."
        return None  # sem conserto: a baixa é o desfecho esperado
    return "Não é possível remover produto emprestado."
