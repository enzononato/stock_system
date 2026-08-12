import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from app.db.unidade_db import get_unidade_manager
from app.schemas.unidades import UnidadeCreate

units = [
    {
        "nome": "Revalle Petrolina",
        "razao_social": "BEIRA RIO REVENDA DE BEBIDAS LTDA",
        "cnpj": "07.717.961/0001-60",
        "endereco": "AV AVENIDA MARIO RODRIGUES COELHO, n°99 , DISTRITO INDUSTRIAL",
        "cidade": "Petrolina",
        "uf": "PE",
        "cep": None
    },
    {
        "nome": "Revalle Juazeiro",
        "razao_social": "REVENDA VALLE DA INTEGRAÇÃO LTDA",
        "cnpj": "04.690.106/0001-15",
        "endereco": "Centro Industrial São Francisco, Quadra QID, Lotes 5/6 s/n, Av. João Paulo II",
        "cidade": "Juazeiro",
        "uf": "BA",
        "cep": "48905-630"
    },
    {
        "nome": "Revalle Ribeira",
        "razao_social": "REVENDA REVALLE DO NORDESTE DA BAHIA LTDA",
        "cnpj": "28.098.474/0001-37",
        "endereco": "Avenida Edval Calazans de Macedo, Galpão 01, No nº 363, zona norte",
        "cidade": "Ribeira do Pombal",
        "uf": "BA",
        "cep": "48400-000"
    },
    {
        "nome": "Revalle Paulo Afonso",
        "razao_social": "REVENDA REVALLE DO NORDESTE DA BAHIA LTDA",
        "cnpj": "28.098.474/0002-18",
        "endereco": "RODOVIA BA 210, Nº4, QUADRA 201 LOTE4 , BAIRRO TANCREDO NEVES II",
        "cidade": "Paulo Afonso",
        "uf": "BA",
        "cep": "48601-901"
    },
    {
        "nome": "Revalle Bonfim",
        "razao_social": "REVENDA VALLE DA INTEGRAÇÃO LTDA",
        "cnpj": "04.690.106/0003-87",
        "endereco": "Rodovia Lomanto Junior KM 104, BR 407, S/N, KM 104",
        "cidade": "Senhor do Bonfim",
        "uf": "BA",
        "cep": "48970-000"
    },
    {
        "nome": "Revalle Serrinha",
        "razao_social": "REVENDA REVALLE AGRESTE DA BAHIA LTDA",
        "cnpj": "54.677.520/0002-43",
        "endereco": "Avenida Lomanto Junior Margens da Br 116 QUADRA 537 LOTE 290",
        "cidade": "Serrinha",
        "uf": "BA",
        "cep": "48700-000"
    },
    {
        "nome": "Revalle Alagoinhas",
        "razao_social": "REVENDA REVALLE AGRESTE DA BAHIA LTDA",
        "cnpj": "54.677.520/0001-62",
        "endereco": "Rodovia Governador Mario Covas QUADRA 294 LOTE 0410",
        "cidade": "Alagoinhas",
        "uf": "BA",
        "cep": "48013-756"
    }
]

def main():
    manager = get_unidade_manager()
    for u in units:
        validated = UnidadeCreate(**u).model_dump(exclude_none=True)
        existing = manager.get_unidade_por_nome(validated["nome"])
        if existing:
            ok, msg = manager.update_unidade(existing["id"], validated)
            status = "ATUALIZADO"
        else:
            ok, msg = manager.add_unidade(validated)
            status = "CRIADO"
        print(f"[{status}] {validated['nome']}: {msg}")

if __name__ == "__main__":
    main()
