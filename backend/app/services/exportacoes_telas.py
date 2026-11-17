"""Colunas e linhas de cada tela exportável em Excel/PDF (F7).

Cada `planilha_*` consulta a camada de dados com os MESMOS filtros e a MESMA ordem da tela e
devolve uma `Planilha`; `gerar` entrega o arquivo por `app.services.exports`. Unidade/revenda sai
em maiúsculas (`Coluna(maiusculas=True)`) e o CPF nunca sai por inteiro.
"""
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from typing import Literal, Optional

from fastapi import Response

from app.db.acesso_db import descrever_item
from app.db.inventory_manager_db import InventoryDBManager
from app.db.linha_db import LinhaDBManager
from app.db.manutencao_db import ROTULOS_STATUS, ManutencaoDBManager
from app.schemas.linhas import formatar_numero_linha
from app.services.exports import Coluna, gerar_excel, gerar_pdf, resposta_arquivo

Formato = Literal["xlsx", "pdf"]

LIMITE_LINHAS = {"xlsx": 50_000, "pdf": 5_000}

ROTULO_ORDENACAO = {
    "id": "Patrimônio",
    "tipo": "Tipo",
    "equipamento": "Equipamento",
    "status": "Status",
    "colaborador": "Colaborador",
    "setor": "Setor",
    "cargo": "Cargo",
    "unidade": "Unidade",
    "cadastro": "Cadastro",
}

ROTULO_STATUS_LINHA = {"ativa": "Ativa", "suspensa": "Suspensa", "cancelada": "Cancelada"}


class ExportacaoGrande(Exception):
    """A consulta tem mais linhas do que o formato aceita; `detail` é a mensagem em português."""

    def __init__(self, detail: str):
        super().__init__(detail)
        self.detail = detail


def _milhar(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def checar_limite(total: int, formato: str) -> None:
    limite = LIMITE_LINHAS[formato]
    if total <= limite:
        return
    nome = "PDF" if formato == "pdf" else "Excel"
    sugestao = " ou exporte em Excel." if formato == "pdf" else "."
    raise ExportacaoGrande(
        f"A exportação em {nome} está limitada a {_milhar(limite)} linhas e esta consulta tem "
        f"{_milhar(total)}. Refine os filtros{sugestao}"
    )


@dataclass
class Planilha:
    titulo: str
    colunas_excel: list[Coluna]
    colunas_pdf: list[Coluna]
    linhas: list[dict]
    filtros: dict[str, str] = field(default_factory=dict)
    nome_base: str = "exportacao"
    paisagem: bool = True


def gerar(planilha: Planilha, formato: Formato) -> Response:
    if formato == "pdf":
        conteudo = gerar_pdf(
            planilha.titulo, planilha.colunas_pdf, planilha.linhas, planilha.filtros,
            paisagem=planilha.paisagem,
        )
    else:
        conteudo = gerar_excel(planilha.titulo, planilha.colunas_excel, planilha.linhas, planilha.filtros)
    return resposta_arquivo(conteudo, planilha.nome_base, formato)


def data_hora_br(v: Optional[datetime]) -> Optional[str]:
    return v.strftime("%d/%m/%Y %H:%M") if v else None


def _data_br(v: Optional[date]) -> Optional[str]:
    return v.strftime("%d/%m/%Y") if v else None


def _preenchidos(**filtros: Optional[str]) -> dict[str, str]:
    return {nome: valor for nome, valor in filtros.items() if valor}


# ── Estoque ─────────────────────────────────────────────────────────────────

def planilha_estoque(
    inv: InventoryDBManager,
    *,
    tipo: Optional[str],
    status: Optional[str],
    revenda: Optional[str],
    search: Optional[str],
    colaborador_id: Optional[int],
    sort: str,
    direcao: str,
    formato: Formato,
    colaborador_nome: Optional[str] = None,
) -> Planilha:
    itens, total = inv.list_items(
        tipo=tipo, status=status, revenda=revenda, search=search, colaborador_id=colaborador_id,
        sort=sort, direcao=direcao, limit=LIMITE_LINHAS[formato] + 1, offset=0,
    )
    checar_limite(total, formato)
    linhas = []
    for i in itens:
        linhas.append({
            **i,
            "colaborador": i.get("colaborador_nome") or i.get("assigned_to"),
            "equipamento": " ".join(p for p in (i.get("brand"), i.get("model")) if p),
        })
    filtros = _preenchidos(
        Tipo=tipo,
        Status=status,
        Unidade=(revenda or "").upper(),
        Colaborador=colaborador_nome or (f"#{colaborador_id}" if colaborador_id else None),
        Busca=(search or "").strip(),
    )
    filtros["Ordenação"] = (
        f"{ROTULO_ORDENACAO[sort]} ({'crescente' if direcao == 'asc' else 'decrescente'})"
    )
    return Planilha(
        titulo="Estoque de equipamentos",
        colunas_excel=[
            Coluna("id", "Patrimônio", 12, "numero"),
            Coluna("tipo", "Tipo", 16),
            Coluna("brand", "Marca", 16),
            Coluna("model", "Modelo", 20),
            Coluna("identificador", "Identificador", 20),
            Coluna("status", "Status", 18),
            Coluna("colaborador", "Colaborador", 26),
            Coluna("colaborador_setor", "Setor", 18),
            Coluna("colaborador_cargo", "Cargo", 20),
            Coluna("superior_nome", "Superior", 24),
            Coluna("revenda", "Unidade", 20, maiusculas=True),
            Coluna("local_nome", "Local", 20),
            Coluna("date_registered", "Cadastro", 12, "data"),
            Coluna("garantia_ate", "Garantia até", 13, "data"),
            Coluna("valor_compra", "Valor de compra", 16, "moeda"),
        ],
        colunas_pdf=[
            Coluna("id", "Patrimônio", 10, "numero"),
            Coluna("tipo", "Tipo", 12),
            Coluna("equipamento", "Equipamento", 24),
            Coluna("identificador", "Identificador", 16),
            Coluna("status", "Status", 14),
            Coluna("colaborador", "Colaborador", 22),
            Coluna("colaborador_setor", "Setor", 14),
            Coluna("revenda", "Unidade", 16, maiusculas=True),
            Coluna("local_nome", "Local", 14),
        ],
        linhas=linhas,
        filtros=filtros,
        nome_base="estoque",
    )


# ── Histórico ───────────────────────────────────────────────────────────────

def planilha_historico(
    inv: InventoryDBManager, *, search: Optional[str], item_id: Optional[int], formato: Formato
) -> Planilha:
    registros, total = inv.list_history(
        search=search, item_id=item_id, limit=LIMITE_LINHAS[formato] + 1, offset=0
    )
    checar_limite(total, formato)
    linhas = [{**r, "data": data_hora_br(r.get("data_operacao"))} for r in registros]
    filtros = _preenchidos(
        Equipamento=f"#{item_id}" if item_id is not None else None,
        Busca=(search or "").strip(),
    )
    return Planilha(
        titulo="Histórico de operações",
        colunas_excel=[
            Coluna("id", "ID", 8, "numero"),
            Coluna("data", "Data", 17),
            Coluna("operation", "Operação", 22),
            Coluna("operador", "Operador", 16),
            Coluna("tipo", "Tipo", 14),
            Coluna("marca", "Marca", 14),
            Coluna("modelo", "Modelo", 18),
            Coluna("identificador", "Identificador", 18),
            Coluna("usuario", "Usuário", 22),
            Coluna("setor", "Setor", 16),
            Coluna("revenda", "Unidade", 20, maiusculas=True),
            Coluna("details", "Detalhes", 40),
        ],
        colunas_pdf=[
            Coluna("id", "ID", 6, "numero"),
            Coluna("data", "Data", 14),
            Coluna("operation", "Operação", 18),
            Coluna("operador", "Operador", 12),
            Coluna("tipo", "Tipo", 11),
            Coluna("marca", "Marca", 11),
            Coluna("modelo", "Modelo", 14),
            Coluna("identificador", "Identificador", 14),
            Coluna("usuario", "Usuário", 16),
            Coluna("setor", "Setor", 12),
            Coluna("revenda", "Unidade", 16, maiusculas=True),
        ],
        linhas=linhas,
        filtros=filtros,
        nome_base="historico",
    )


# ── Manutenções ─────────────────────────────────────────────────────────────

def planilha_manutencoes(
    db: ManutencaoDBManager,
    *,
    status: Optional[str],
    item_id: Optional[int],
    revenda: Optional[str],
    tipo: Optional[str],
    de: Optional[date],
    ate: Optional[date],
    busca: Optional[str],
    formato: Formato,
) -> Planilha:
    registros, total, custo_total = db.listar(
        status=status, item_id=item_id, revenda=revenda, tipo=tipo, de=de, ate=ate, busca=busca,
        limit=LIMITE_LINHAS[formato] + 1, offset=0,
    )
    checar_limite(total, formato)
    linhas: list[dict] = []
    for m in registros:
        equipamento = " ".join(
            p for p in (m.get("item_tipo"), m.get("item_brand"), m.get("item_model")) if p
        )
        linhas.append({
            **m,
            "aberta": data_hora_br(m.get("aberta_em")),
            "equipamento": f"#{m['item_id']} {equipamento}".strip(),
            "status_rotulo": ROTULOS_STATUS.get(m["status"], m["status"]),
            "retorno": data_hora_br(m.get("retornada_em")),
        })
    linhas.append({"id": "TOTAL", "custo": custo_total})
    periodo = " a ".join(p for p in (_data_br(de), _data_br(ate)) if p)
    filtros = _preenchidos(
        Status=("Em andamento" if status == "em_andamento" else ROTULOS_STATUS.get(status or "", status)),
        Equipamento=f"#{item_id}" if item_id is not None else None,
        Unidade=(revenda or "").upper(),
        Tipo=tipo,
        Período=periodo,
        Busca=(busca or "").strip(),
    )
    return Planilha(
        titulo="Manutenções",
        colunas_excel=[
            Coluna("id", "Nº", 8, "numero"),
            Coluna("aberta", "Aberta em", 17),
            Coluna("equipamento", "Equipamento", 34),
            Coluna("item_identificador", "Identificador", 18),
            Coluna("item_revenda", "Unidade", 20, maiusculas=True),
            Coluna("defeito", "Defeito", 34),
            Coluna("assistencia", "Assistência", 22),
            Coluna("status_rotulo", "Status", 14),
            Coluna("previsao_retorno", "Previsão", 12, "data"),
            Coluna("retorno", "Retorno", 17),
            Coluna("custo", "Custo", 14, "moeda"),
            Coluna("custo_acumulado_item", "Custo acumulado do equipamento", 18, "moeda"),
        ],
        colunas_pdf=[
            Coluna("id", "Nº", 6, "numero"),
            Coluna("aberta", "Aberta em", 14),
            Coluna("equipamento", "Equipamento", 26),
            Coluna("item_identificador", "Identificador", 14),
            Coluna("item_revenda", "Unidade", 16, maiusculas=True),
            Coluna("defeito", "Defeito", 26),
            Coluna("assistencia", "Assistência", 16),
            Coluna("status_rotulo", "Status", 12),
            Coluna("previsao_retorno", "Previsão", 11, "data"),
            Coluna("retorno", "Retorno", 14),
            Coluna("custo", "Custo", 12, "moeda"),
        ],
        linhas=linhas,
        filtros=filtros,
        nome_base="manutencoes",
    )


# ── Linhas telefônicas ──────────────────────────────────────────────────────

def planilha_linhas(
    db: LinhaDBManager,
    *,
    busca: Optional[str],
    status: Optional[str],
    unidade_id: Optional[int],
    unidade_nome: Optional[str],
    operadora: Optional[str],
    formato: Formato,
) -> Planilha:
    registros, total = db.listar(
        busca=busca, status=status, unidade_id=unidade_id, operadora=operadora,
        limit=LIMITE_LINHAS[formato] + 1, offset=0,
    )
    checar_limite(total, formato)
    linhas: list[dict] = []
    total_ativas = Decimal("0.00")
    for r in registros:
        aparelho = descrever_item({
            "tipo": r.get("item_tipo"), "brand": r.get("item_brand"),
            "model": r.get("item_model"), "identificador": r.get("item_identificador"),
        })
        if aparelho and r.get("item_ativo") is not None and not r["item_ativo"]:
            aparelho += " (baixado)"
        colaborador = r.get("colaborador_nome")
        if colaborador and r.get("colaborador_situacao") == "desligado":
            colaborador += " (desligado)"
        if r["status"] == "ativa":
            total_ativas += Decimal(r["custo_mensal"])
        linhas.append({
            **r,
            "numero_formatado": formatar_numero_linha(r["numero"]),
            "status_rotulo": ROTULO_STATUS_LINHA.get(r["status"], r["status"]),
            "aparelho": aparelho,
            "colaborador": colaborador,
        })
    linhas.append({"numero_formatado": "Total mensal (linhas ativas)", "custo_mensal": total_ativas})
    filtros = _preenchidos(
        Busca=(busca or "").strip(),
        Status=ROTULO_STATUS_LINHA.get(status or "", status),
        Unidade=(unidade_nome or "").upper(),
        Operadora=operadora,
    )
    return Planilha(
        titulo="Linhas telefônicas",
        colunas_excel=[
            Coluna("numero_formatado", "Número", 18),
            Coluna("operadora", "Operadora", 14),
            Coluna("plano", "Plano", 22),
            Coluna("custo_mensal", "Custo mensal", 14, "moeda"),
            Coluna("status_rotulo", "Status", 12),
            Coluna("unidade_nome", "Unidade", 20, maiusculas=True),
            Coluna("aparelho", "Aparelho", 36),
            Coluna("colaborador", "Colaborador", 26),
            Coluna("iccid", "ICCID", 24),
        ],
        colunas_pdf=[
            Coluna("numero_formatado", "Número", 18),
            Coluna("operadora", "Operadora", 12),
            Coluna("plano", "Plano", 20),
            Coluna("custo_mensal", "Custo mensal", 13, "moeda"),
            Coluna("status_rotulo", "Status", 11),
            Coluna("unidade_nome", "Unidade", 18, maiusculas=True),
            Coluna("aparelho", "Aparelho", 30),
            Coluna("colaborador", "Colaborador", 22),
        ],
        linhas=linhas,
        filtros=filtros,
        nome_base="linhas",
    )


# ── Termos pendentes de assinatura ──────────────────────────────────────────

ROTULO_TIPO_TERMO = {"entrega": "Entrega", "devolucao": "Devolução", "uso": "Uso"}


def planilha_termos_pendentes(db, *, busca: Optional[str], formato: Formato) -> Planilha:
    from app.core.tempo import agora_local
    from app.services.painel import ROTULO_STATUS_TERMO

    registros, total = db.termos_pendentes(busca=busca, limit=LIMITE_LINHAS[formato] + 1)
    checar_limite(total, formato)
    agora, agora_termos = datetime.now(), agora_local()
    linhas: list[dict] = []
    for t in registros:
        referencia = t.get("referencia")
        relogio = agora if t["legado"] else agora_termos
        equipamento = " ".join(
            p for p in (t.get("item_tipo"), t.get("item_brand"), t.get("item_model")) if p
        )
        linhas.append({
            "solicitacao": t["termo_id"] if t["termo_id"] is not None else "—",
            "equipamento": f"#{t['item_id']} {equipamento}".strip(),
            "identificador": t.get("item_identificador"),
            "revenda": t.get("item_revenda"),
            "colaborador": t.get("colaborador_nome"),
            "termo": ROTULO_TIPO_TERMO.get(t["tipo"], t["tipo"]),
            "status": ROTULO_STATUS_TERMO.get(t["status"], t["status"]),
            "solicitado_por": t.get("solicitado_por"),
            "solicitado_em": data_hora_br(t.get("solicitado_em")),
            "enviado_em": data_hora_br(t.get("enviado_em")),
            "agendado_para": data_hora_br(t.get("agendado_para")),
            "dias": (relogio - referencia).days if referencia else None,
            "motivo": t.get("motivo"),
        })
    colunas = [
        Coluna("solicitacao", "Solicitação", 12, "numero"),
        Coluna("equipamento", "Equipamento", 34),
        Coluna("identificador", "Identificador", 18),
        Coluna("revenda", "Unidade", 20, maiusculas=True),
        Coluna("colaborador", "Colaborador", 26),
        Coluna("termo", "Termo", 11),
        Coluna("status", "Status", 17),
        Coluna("solicitado_por", "Solicitado por", 16),
        Coluna("solicitado_em", "Solicitado em", 17),
        Coluna("enviado_em", "Enviado em", 17),
        Coluna("agendado_para", "Agendado para", 17),
        Coluna("dias", "Dias pendente", 13, "numero"),
    ]
    return Planilha(
        titulo="Termos pendentes de assinatura",
        colunas_excel=[*colunas, Coluna("motivo", "Motivo", 40)],
        colunas_pdf=colunas,
        linhas=linhas,
        filtros=_preenchidos(Busca=(busca or "").strip()),
        nome_base="termos_pendentes",
    )


# ── Relatório mensal ────────────────────────────────────────────────────────

# Mesmos cabeçalhos do CSV de `routers/reports.py` (copiados: aquele arquivo não muda).
HEADERS_PT = {
    "history_id": "ID Histórico",
    "item_id": "ID Item",
    "peripheral_id": "ID Periférico",
    "operador": "Operador",
    "usuario": "Usuário",
    "cpf": "CPF",
    "cargo": "Cargo",
    "center_cost": "Centro de Custo",
    "setor": "Setor",
    "fornecedor": "Fornecedor",
    "revenda": "Revenda",
    "details": "Detalhes",
    "data_emprestimo": "Data Empréstimo",
    "data_confirmacao": "Data Confirmação",
    "data_devolucao": "Data Devolução",
    "operation_type": "Operação",
    "tipo": "Tipo",
    "brand": "Marca",
    "model": "Modelo",
    "identificador": "Identificador",
    "nota_fiscal": "Nota Fiscal",
}
_CHAVES_COM_DATA = ("data_emprestimo", "data_confirmacao", "data_devolucao")
_CHAVES_NUMERICAS = ("history_id", "item_id", "peripheral_id")
_CHAVES_PDF = (
    "operation_type", "data_emprestimo", "data_devolucao", "usuario", "cargo", "setor",
    "center_cost", "revenda", "tipo", "brand", "model", "identificador", "operador",
)


def mascarar_cpf(valor) -> Optional[str]:
    """`***.456.789-**`; formato inesperado vira `***`; vazio vira `None`."""
    if valor is None or str(valor).strip() == "":
        return None
    digitos = "".join(ch for ch in str(valor) if ch.isdigit())
    if len(digitos) != 11:
        return "***"
    return f"***.{digitos[3:6]}.{digitos[6:9]}-**"


def planilha_relatorio_mensal(
    inv: InventoryDBManager, *, ano: int, mes: int, formato: Formato
) -> Planilha:
    registros = inv.generate_monthly_report(ano, mes)
    checar_limite(len(registros), formato)
    linhas = []
    for r in registros:
        linha = dict(r)
        linha["cpf"] = mascarar_cpf(r.get("cpf"))
        for chave in _CHAVES_COM_DATA:
            linha[chave] = data_hora_br(r.get(chave))
        linhas.append(linha)

    def coluna(chave: str) -> Coluna:
        return Coluna(
            chave, HEADERS_PT[chave], 12 if chave in _CHAVES_NUMERICAS else 18,
            "numero" if chave in _CHAVES_NUMERICAS else "texto", maiusculas=(chave == "revenda"),
        )

    return Planilha(
        titulo=f"Relatório mensal {mes:02d}/{ano}",
        colunas_excel=[coluna(c) for c in HEADERS_PT],
        colunas_pdf=[coluna(c) for c in _CHAVES_PDF],
        linhas=linhas,
        filtros={"Período": f"{mes:02d}/{ano}"},
        nome_base=f"relatorio_mensal_{ano}_{mes:02d}",
    )
