import os
import sys

# --- Lógica para encontrar recursos como o logo.ico para o pyinstaller ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
TERMS_DIR = 'terms'
REMOVAL_NOTES_DIR = 'notas_remocao' # Diretório para notas de remoção
SIGNED_TERMS_DIR = 'termos_assinados' # Diretório para termos assinados
RETURN_TERMS_DIR = 'termos_devolucao' # Diretório para os termos de devolução gerados
SIGNED_RETURN_TERMS_DIR = 'termos_devolucao_assinados' # Diretório para os termos de devolução assinados


if not os.path.exists(TERMS_DIR):
    os.makedirs(TERMS_DIR)
    
if not os.path.exists(REMOVAL_NOTES_DIR):
    os.makedirs(REMOVAL_NOTES_DIR)

if not os.path.exists(SIGNED_TERMS_DIR):
    os.makedirs(SIGNED_TERMS_DIR)
    
if not os.path.exists(RETURN_TERMS_DIR):
    os.makedirs(RETURN_TERMS_DIR)

if not os.path.exists(SIGNED_RETURN_TERMS_DIR):
    os.makedirs(SIGNED_RETURN_TERMS_DIR)

# Usuário "Mãe" e senha
ADMIN_USER = "mãe"
ADMIN_PASS = "Revalle@123"

# Opções fixas para Centro de Custo e Revendas
CENTER_COST_OPTIONS = [
    "101 - Puxada",
    "202 - Armazém",
    "301 - Administrativo",
    "401 - Vendas",
    "501 - Entrega",
    "601 - CSC"
]

REVENDAS_OPTIONS = [
    "Revalle Juazeiro",
    "Revalle Bonfim",
    "Revalle Petrolina",
    "Revalle Ribeira",
    "Revalle Paulo Afonso",
    "Revalle Alagoinhas",
    "Revalle Serrinha"
]

TERMO_MODELOS = {
    "Revalle Juazeiro": resource_path("modelos/termo_juazeiro.docx"),
    "Revalle Bonfim": resource_path("modelos/termo_bonfim.docx"),
    "Revalle Petrolina": resource_path("modelos/termo_petrolina.docx"),
    "Revalle Ribeira": resource_path("modelos/termo_ribeira.docx"),
    "Revalle Paulo Afonso": resource_path("modelos/termo_pauloafonso.docx"),
    "Revalle Alagoinhas": resource_path("modelos/termo_alagoinhas.docx"),
    "Revalle Serrinha": resource_path("modelos/termo_serrinha.docx"),
}

TERMO_DEVOLUCAO_MODELOS = {
    "Revalle Juazeiro": resource_path("modelos/devolucao/termo_devolucao_juazeiro.docx"),
    "Revalle Bonfim": resource_path("modelos/devolucao/termo_devolucao_bonfim.docx"),
    "Revalle Petrolina": resource_path("modelos/devolucao/termo_devolucao_petrolina.docx"),
    "Revalle Ribeira": resource_path("modelos/devolucao/termo_devolucao_ribeira.docx"),
    "Revalle Paulo Afonso": resource_path("modelos/devolucao/termo_devolucao_pauloafonso.docx"),
    "Revalle Alagoinhas": resource_path("modelos/devolucao/termo_devolucao_alagoinhas.docx"),
    "Revalle Serrinha": resource_path("modelos/devolucao/termo_devolucao_serrinha.docx"),
}


SETORES_OPTIONS = [
    "Contabilidade",
    "Financeiro",
    "Departamento pessoal",
    "Cultura",
    "Gente",
    "Armazém",
    "Distribuição",
    "Segurança",
    "Compras",
    "TI",
    "CME",
    "Vendas",
    "Puxada",
    "Faturamento",
    "Portaria",
    "Auditório",
    "Caixa",
    "Diretoria"
]

# Dicionário com os motivos de remoção/substituição e se exigem anexo
REMOVAL_REASONS = {
    "Roubo": True,          # True = anexo obrigatório
    "Perda": False,         # False = anexo não obrigatório
    "Obsolescência": False,
    "Doação": True,
    "Venda": True
}