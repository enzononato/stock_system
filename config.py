import os

TERMS_DIR = 'terms'

if not os.path.exists(TERMS_DIR):
    os.makedirs(TERMS_DIR)

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
    "Revalle Juazeiro": "modelos/termo_juazeiro.docx",
    "Revalle Bonfim": "modelos/termo_bonfim.docx",
    "Revalle Petrolina": "modelos/termo_petrolina.docx",
    "Revalle Ribeira": "modelos/termo_ribeira.docx",
    "Revalle Paulo Afonso": "modelos/termo_pauloafonso.docx",
    "Revalle Alagoinhas": "modelos/termo_alagoinhas.docx",
    "Revalle Serrinha": "modelos/termo_serrinha.docx",
}
