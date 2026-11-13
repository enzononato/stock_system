"""Regras e textos dos alertas do painel (F7). Tudo calculado na hora; nada é gravado.

Os limiares vêm de `app.core.config.settings` no momento da chamada (o `tests/test_auth.py`
recarrega o módulo de configuração, então o objeto `settings` pode ser trocado).
"""
from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from typing import Literal, Optional

from app.core import config
from app.core.tempo import agora_local
from app.db.acesso_db import descrever_item
from app.db.manutencao_db import ROTULOS_STATUS
from app.db.painel_db import PainelDBManager
from app.schemas.linhas import formatar_numero_linha
from app.services.termos import ROTULO_TIPO

TipoAlerta = Literal[
    "colaborador_desligado",
    "devolucao_pendente",
    "termo_pendente",
    "manutencao_parada",
    "linha_aparelho_baixado",
    "garantia_vencendo",
]

TIPOS_ALERTA: tuple[str, ...] = (
    "colaborador_desligado",
    "devolucao_pendente",
    "termo_pendente",
    "manutencao_parada",
    "linha_aparelho_baixado",
    "garantia_vencendo",
)

SEVERIDADE = {
    "colaborador_desligado": "alta",
    "devolucao_pendente": "alta",
    "termo_pendente": "media",
    "manutencao_parada": "media",
    "linha_aparelho_baixado": "media",
    "garantia_vencendo": "baixa",
}
_RANKING_SEVERIDADE = {"alta": 0, "media": 1, "baixa": 2}

ROTULO_STATUS_TERMO = {
    "agendado": "Agendado",
    "enviado": "Enviado",
    "manual": "Assinatura manual",
    "erro": "Erro",
    "sem_solicitacao": "Sem solicitação",
}


@dataclass(frozen=True)
class Limites:
    garantia_dias: int
    termo_pendente_dias: int
    devolucao_pendente_dias: int
    manutencao_dias: int


def limites_atuais() -> Limites:
    s = config.settings
    return Limites(
        garantia_dias=int(s.ALERTA_GARANTIA_DIAS),
        termo_pendente_dias=int(s.ALERTA_TERMO_PENDENTE_DIAS),
        devolucao_pendente_dias=int(s.ALERTA_DEVOLUCAO_PENDENTE_DIAS),
        manutencao_dias=int(s.ALERTA_MANUTENCAO_DIAS),
    )


def _equip(item_id: int, linha: dict) -> str:
    return f"#{item_id} {descrever_item(linha) or 'Equipamento'}"


def _alerta(
    tipo: str, id_chave, titulo: str, descricao: str, link: str, *, item_id=None,
    colaborador_id=None, desde=None, dias: Optional[int] = None,
) -> dict:
    return {
        "chave": f"{tipo}:{id_chave}",
        "tipo": tipo,
        "severidade": SEVERIDADE[tipo],
        "titulo": titulo,
        "descricao": descricao,
        "link": link,
        "item_id": item_id,
        "colaborador_id": colaborador_id,
        "desde": desde,
        "dias": dias,
    }


def _passou(agora: datetime, desde: datetime, dias: int) -> bool:
    return agora - desde > timedelta(days=dias)


def _garantias(db: PainelDBManager, hoje: date, limites: Limites) -> list[dict]:
    saida = []
    for i in db.garantias_entre(hoje, hoje + timedelta(days=limites.garantia_dias)):
        dias = (i["garantia_ate"] - hoje).days
        titulo = "Garantia vence hoje" if dias == 0 else f"Garantia vence em {dias} dia(s)"
        unidade = (i["revenda"] or "").strip().upper() or "SEM UNIDADE"
        saida.append(
            _alerta(
                "garantia_vencendo", i["id"], titulo, f"{_equip(i['id'], i)} · {unidade}",
                f"/items/{i['id']}", item_id=i["id"], desde=i["garantia_ate"], dias=dias,
            )
        )
    return saida


def _termos(db: PainelDBManager, agora: datetime, agora_termos: datetime, limites: Limites) -> list[dict]:
    linhas, _ = db.termos_pendentes()
    saida = []
    for t in linhas:
        desde = t["referencia"]
        if desde is None:
            continue
        referencia_agora = agora if t["legado"] else agora_termos
        if not _passou(referencia_agora, desde, limites.termo_pendente_dias):
            continue
        dias = (referencia_agora - desde).days
        if t["legado"]:
            titulo = f"Termo de entrega sem assinatura há {dias} dias"
            chave = f"item{t['item_id']}"
        else:
            titulo = f"Termo de {ROTULO_TIPO.get(t['tipo'], t['tipo'])} pendente há {dias} dias"
            chave = t["termo_id"]
        linha_item = {
            "tipo": t["item_tipo"], "brand": t["item_brand"], "model": t["item_model"],
            "identificador": t["item_identificador"],
        }
        descricao = (
            f"{_equip(t['item_id'], linha_item)} · {t['colaborador_nome'] or 'sem colaborador'}"
            f" · {ROTULO_STATUS_TERMO.get(t['status'], t['status'])}"
        )
        saida.append(
            _alerta(
                "termo_pendente", chave, titulo, descricao, f"/items/{t['item_id']}?aba=termos",
                item_id=t["item_id"], colaborador_id=t["colaborador_id"], desde=desde, dias=dias,
            )
        )
    return saida


def _devolucoes(db: PainelDBManager, agora: datetime, limites: Limites) -> list[dict]:
    saida = []
    for i in db.devolucoes_pendentes():
        desde = i["desde"]
        if desde is None:
            titulo, dias = "Devolução pendente sem data registrada", None
        elif _passou(agora, desde, limites.devolucao_pendente_dias):
            dias = (agora - desde).days
            titulo = f"Devolução pendente há {dias} dias"
        else:
            continue
        quem = i["colaborador_nome"] or i["assigned_to"] or "sem colaborador"
        saida.append(
            _alerta(
                "devolucao_pendente", i["id"], titulo, f"{_equip(i['id'], i)} · {quem}",
                f"/items/{i['id']}", item_id=i["id"], colaborador_id=i["colaborador_id"],
                desde=desde, dias=dias,
            )
        )
    return saida


def _manutencoes(db: PainelDBManager, agora: datetime, limites: Limites) -> list[dict]:
    saida = []
    for m in db.manutencoes_em_andamento():
        desde = m["ultimo_movimento"]
        if desde is None or not _passou(agora, desde, limites.manutencao_dias):
            continue
        dias = (agora - desde).days
        descricao = (
            f"{_equip(m['item_id'], m)} · {ROTULOS_STATUS[m['status']]} · "
            f"{m['assistencia'] or 'sem assistência'}"
        )
        saida.append(
            _alerta(
                "manutencao_parada", m["id"], f"Manutenção parada há {dias} dias", descricao,
                f"/items/{m['item_id']}?aba=manutencoes", item_id=m["item_id"], desde=desde, dias=dias,
            )
        )
    return saida


def _desligados(db: PainelDBManager) -> list[dict]:
    saida = []
    for i in db.desligados_com_equipamento():
        saida.append(
            _alerta(
                "colaborador_desligado", i["item_id"],
                f"{i['colaborador_nome']} (desligado) ainda está com equipamento",
                f"{_equip(i['item_id'], i)} · {i['status']}", f"/items/{i['item_id']}",
                item_id=i["item_id"], colaborador_id=i["colaborador_id"],
            )
        )
    return saida


def _linhas(db: PainelDBManager) -> list[dict]:
    saida = []
    for linha in db.linhas_em_aparelho_baixado():
        numero = linha["numero"]
        saida.append(
            _alerta(
                "linha_aparelho_baixado", linha["linha_id"],
                f"Linha {formatar_numero_linha(numero)} ativa em aparelho baixado",
                f"{_equip(linha['item_id'], linha)} · {(linha['unidade_nome'] or '').upper()}",
                f"/linhas?busca={numero}", item_id=linha["item_id"],
            )
        )
    return saida


def _chave_de_ordem(a: dict):
    # Dentro da severidade: sem data primeiro, depois a mais antiga (dias decrescente); na garantia,
    # a que vence primeiro (dias crescente), porque "dias" ali é o tempo que falta.
    tem_data = a["desde"] is not None
    dias = a["dias"] or 0
    return (
        _RANKING_SEVERIDADE[a["severidade"]],
        tem_data,
        dias if a["tipo"] == "garantia_vencendo" else -dias,
        a["chave"],
    )


def calcular_alertas(
    db: PainelDBManager,
    *,
    agora: Optional[datetime] = None,
    agora_termos: Optional[datetime] = None,
    limites: Optional[Limites] = None,
) -> dict:
    agora = agora or datetime.now()
    agora_termos = agora_termos or agora_local()
    limites = limites or limites_atuais()

    alertas = [
        *_desligados(db),
        *_devolucoes(db, agora, limites),
        *_termos(db, agora, agora_termos, limites),
        *_manutencoes(db, agora, limites),
        *_linhas(db),
        *_garantias(db, agora.date(), limites),
    ]
    alertas.sort(key=_chave_de_ordem)
    por_tipo = {tipo: 0 for tipo in TIPOS_ALERTA}
    for a in alertas:
        por_tipo[a["tipo"]] += 1
    return {
        "gerado_em": agora,
        "limites": asdict(limites),
        "total": len(alertas),
        "por_tipo": por_tipo,
        "alertas": alertas,
    }
