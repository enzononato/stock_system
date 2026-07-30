"""
Script IDEMPOTENTE de seed das 7 Unidades (dados jurídicos de cada revenda).

Os dados abaixo foram extraídos dos 7 templates de termo de responsabilidade
existentes e validados pelos dígitos verificadores do CNPJ (ver
`app/schemas/unidades.py::validar_cnpj`) antes de serem fixados aqui. Já vêm
no formato normalizado que a camada de dados grava (`00.000.000/0000-00`
para CNPJ, `00000-000` para CEP), então este script insere diretamente via
`UnidadeDBManager.add_unidade` — sem passar pelos schemas Pydantic, que
existem para validar entrada de usuário via API, não uma lista fixa já
conferida manualmente.

Petrolina não tem CEP no documento de origem (é a única unidade em PE, as
demais são na Bahia) — fica vazio, o usuário preenche depois pela tela de
Unidades.

Idempotente: não sobrescreve unidade já existente (checa por nome antes de
inserir), então rodar quantas vezes for preciso é seguro.

Uso:
    python backend/scripts/seed_unidades.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from app.db.unidade_db import UnidadeDBManager  # noqa: E402

UNIDADES_SEED = [
    {
        "nome": "Revalle Juazeiro",
        "razao_social": "REVENDA VALLE DA INTEGRAÇÃO LTDA",
        "cnpj": "04.690.106/0001-15",
        "endereco": "Centro Industrial São Francisco, Quadra QID, Lotes 5/6 s/n, Av. João Paulo II",
        "cep": "48905-630",
        "cidade": "Juazeiro",
        "uf": "BA",
    },
    {
        "nome": "Revalle Bonfim",
        "razao_social": "REVENDA VALLE DA INTEGRAÇÃO LTDA",
        "cnpj": "04.690.106/0003-87",
        "endereco": "Rodovia Lomanto Junior KM 104, BR 407, S/N",
        "cep": "48970-000",
        "cidade": "Senhor do Bonfim",
        "uf": "BA",
    },
    {
        "nome": "Revalle Petrolina",
        "razao_social": "BEIRA RIO REVENDA DE BEBIDAS LTDA",
        "cnpj": "07.717.961/0001-60",
        "endereco": "Avenida Mario Rodrigues Coelho, nº 99, Distrito Industrial",
        "cep": None,  # ausente no documento de origem; preenchido depois pela tela
        "cidade": "Petrolina",
        "uf": "PE",
    },
    {
        "nome": "Revalle Ribeira",
        "razao_social": "REVENDA REVALLE DO NORDESTE DA BAHIA LTDA",
        "cnpj": "28.098.474/0001-37",
        "endereco": "Avenida Edval Calazans de Macedo, Galpão 01, nº 363, Zona Norte",
        "cep": "48400-000",
        "cidade": "Ribeira do Pombal",
        "uf": "BA",
    },
    {
        "nome": "Revalle Paulo Afonso",
        "razao_social": "REVENDA REVALLE DO NORDESTE DA BAHIA LTDA",
        "cnpj": "28.098.474/0002-18",
        "endereco": "Rodovia BA 210, nº 4, Quadra 201, Lote 4, Bairro Tancredo Neves II",
        "cep": "48601-901",
        "cidade": "Paulo Afonso",
        "uf": "BA",
    },
    {
        "nome": "Revalle Alagoinhas",
        "razao_social": "REVENDA REVALLE AGRESTE DA BAHIA LTDA",
        "cnpj": "54.677.520/0001-62",
        "endereco": "Rodovia Governador Mario Covas, Quadra 294, Lote 0410",
        "cep": "48013-756",
        "cidade": "Alagoinhas",
        "uf": "BA",
    },
    {
        "nome": "Revalle Serrinha",
        "razao_social": "REVENDA REVALLE AGRESTE DA BAHIA LTDA",
        "cnpj": "54.677.520/0002-43",
        "endereco": "Avenida Lomanto Junior, Margens da BR 116, Quadra 537, Lote 290",
        "cep": "48700-000",
        "cidade": "Serrinha",
        "uf": "BA",
    },
]


def seed_unidades() -> None:
    db = UnidadeDBManager()
    for dados in UNIDADES_SEED:
        existente = db.get_unidade_por_nome(dados["nome"])
        if existente:
            print(f"Unidade '{dados['nome']}' já existe (id={existente['id']}) — nada a fazer, não sobrescrevo.")
            continue
        ok, msg = db.add_unidade(dados)
        if ok:
            print(msg)
        else:
            print(f"Falha ao inserir '{dados['nome']}': {msg}")


if __name__ == "__main__":
    seed_unidades()
