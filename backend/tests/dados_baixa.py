"""Dados de apagamento de dados (LGPD) reaproveitados pelos testes de baixa."""
from datetime import datetime

APAGAMENTO_OK = {
    "dados_apagados": "sim",
    "metodo_apagamento": "formatacao_segura",
    "apagado_por": "Técnico TI",
    "apagado_em": datetime.now().strftime("%d/%m/%Y"),
}
