"""Fontes da linha do tempo implementadas pela F4: histórico, manutenções,
transferências e baixa. Cada função é registrada com `@registrar_fonte`."""
from decimal import Decimal
from typing import Optional

from app.db.database_mysql import get_cursor
from app.services.baixa import METODOS_APAGAMENTO
from app.services.linha_do_tempo import Evento, registrar_fonte


def moeda_br(valor: Decimal) -> str:
    """`Decimal("1234.5")` -> `"R$ 1.234,50"`."""
    texto = f"{Decimal(valor):,.2f}"
    return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _data_br(valor) -> str:
    return valor.strftime("%d/%m/%Y")


def _consultar(sql: str, params: tuple) -> list[dict]:
    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


def _evento_de_historico(linha: dict) -> Evento:
    operacao = linha["operation"]
    usuario = linha.get("usuario")
    mapa = {
        "Cadastro": ("cadastro", "Cadastro do equipamento"),
        "Edição": ("edicao", "Edição dos dados"),
        "Empréstimo": ("emprestimo", f"Empréstimo para {usuario}" if usuario else "Empréstimo"),
        "Confirmação Empréstimo": ("emprestimo_confirmado", "Empréstimo confirmado"),
        "Devolução": ("devolucao", "Devolução iniciada"),
        "Confirmação Devolução": ("devolucao_confirmada", "Devolução confirmada"),
        "Estorno": ("estorno", "Estorno de operação"),
        "Vínculo Periférico": ("periferico", "Periférico vinculado"),
        "Desvínculo Periférico": ("periferico", "Periférico desvinculado"),
        "Substituição Periférico": ("periferico", "Periférico substituído"),
    }
    tipo, titulo = mapa.get(operacao, ("outro", operacao))
    return Evento(
        ocorrido_em=linha["data_operacao"],
        origem="historico",
        tipo=tipo,
        titulo=titulo,
        descricao=linha.get("details"),
        usuario=linha.get("operador"),
        referencia_id=linha["id"],
    )


@registrar_fonte
def fonte_historico(item_id: int) -> list[Evento]:
    """`history` do item, sem estornados, sem a baixa (fonte própria) e sem as
    edições que são transferências (fonte própria)."""
    linhas = _consultar(
        "SELECT * FROM history WHERE item_id=%s AND is_reversed=0 AND data_operacao IS NOT NULL "
        "AND operation <> 'Exclusão' "
        "AND NOT (operation = 'Edição' AND COALESCE(details, '') LIKE %s) "
        "ORDER BY data_operacao, id",
        (item_id, "Transferência %"),
    )
    return [_evento_de_historico(linha) for linha in linhas]


@registrar_fonte
def fonte_manutencoes(item_id: int) -> list[Evento]:
    eventos: list[Evento] = []
    for m in _consultar("SELECT * FROM manutencoes WHERE item_id=%s ORDER BY id", (item_id,)):
        mid = m["id"]

        def evento(quando, tipo, titulo, descricao=None, usuario=None):
            eventos.append(
                Evento(
                    ocorrido_em=quando, origem="manutencao", tipo=tipo, titulo=titulo,
                    descricao=descricao, usuario=usuario, referencia_id=mid,
                )
            )

        evento(m["aberta_em"], "manutencao_aberta", "Manutenção aberta", m["defeito"], m["aberta_por"])
        if m["enviada_em"]:
            previsao = f"Previsão de retorno: {_data_br(m['previsao_retorno'])}" if m["previsao_retorno"] else None
            evento(m["enviada_em"], "manutencao_enviada", f"Enviado à assistência {m['assistencia']}", previsao)
        if m["em_reparo_em"]:
            evento(m["em_reparo_em"], "manutencao_em_reparo", "Em reparo na assistência")
        if m["retornada_em"]:
            custo = f"Custo: {moeda_br(m['custo'])}" if m["custo"] is not None else None
            if m["status"] == "sem_conserto":
                descricao = "Baixa sugerida." + (f" {custo}" if custo else "")
                evento(
                    m["retornada_em"], "manutencao_sem_conserto", "Manutenção encerrada sem conserto",
                    descricao, m["encerrada_por"],
                )
            else:
                evento(
                    m["retornada_em"], "manutencao_retornada", "Retorno da manutenção",
                    custo, m["encerrada_por"],
                )
        if m["cancelada_em"]:
            evento(
                m["cancelada_em"], "manutencao_cancelada", "Manutenção cancelada",
                usuario=m["encerrada_por"],
            )
    return eventos


@registrar_fonte
def fonte_transferencias(item_id: int) -> list[Evento]:
    linhas = _consultar(
        "SELECT t.*, ld.nome AS de_local_nome, lp.nome AS para_local_nome FROM transferencias t "
        "LEFT JOIN locais ld ON ld.id = t.de_local_id LEFT JOIN locais lp ON lp.id = t.para_local_id "
        "WHERE t.item_id=%s ORDER BY t.transferido_em, t.id",
        (item_id,),
    )
    eventos = []
    for t in linhas:
        descricao: Optional[str] = t["motivo"]
        if t["de_local_nome"] or t["para_local_nome"]:
            descricao += f" · Local: {t['de_local_nome'] or '—'} → {t['para_local_nome'] or '—'}"
        eventos.append(
            Evento(
                ocorrido_em=t["transferido_em"],
                origem="transferencia",
                tipo="transferencia",
                titulo=f"Transferência {t['de_unidade'] or '—'} → {t['para_unidade']}",
                descricao=descricao,
                usuario=t["transferido_por"],
                referencia_id=t["id"],
            )
        )
    return eventos


@registrar_fonte
def fonte_baixa(item_id: int) -> list[Evento]:
    linhas = _consultar(
        "SELECT h.id, h.operador, h.data_operacao, h.details, b.dados_apagados, b.metodo, "
        "b.apagado_por, b.apagado_em FROM history h "
        "LEFT JOIN baixas_apagamento b ON b.history_id = h.id "
        "WHERE h.item_id=%s AND h.operation='Exclusão' AND h.is_reversed=0 AND h.data_operacao IS NOT NULL "
        "ORDER BY h.data_operacao, h.id",
        (item_id,),
    )
    eventos = []
    for h in linhas:
        if h["dados_apagados"] == "sim":
            descricao = (
                f"Dados apagados ({METODOS_APAGAMENTO.get(h['metodo'], h['metodo'])}) "
                f"por {h['apagado_por']} em {_data_br(h['apagado_em'])}"
            )
        elif h["dados_apagados"] == "nao_aplica":
            descricao = "Apagamento de dados: não se aplica"
        else:
            descricao = None  # baixa antiga, sem registro de apagamento
        eventos.append(
            Evento(
                ocorrido_em=h["data_operacao"],
                origem="baixa",
                tipo="baixa",
                titulo=f"Baixa do equipamento — {h['details']}",
                descricao=descricao,
                usuario=h["operador"],
                referencia_id=h["id"],
            )
        )
    return eventos
