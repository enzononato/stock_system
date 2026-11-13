"""Fontes da linha do tempo ligadas pela F7: termos, linha telefônica e inventário.

Cada função é registrada com `@registrar_fonte` (ordem: termos, linha, inventário) e o módulo é
carregado por `MODULOS_DE_FONTES`. O e-mail do destinatário do termo nunca entra nos textos.
"""
from app.db.database_mysql import get_cursor
from app.db.inventario_db import InventarioDBManager
from app.db.linha_db import LinhaDBManager
from app.schemas.linhas import formatar_numero_linha
from app.services.inventario import ROTULOS_CATEGORIA
from app.services.linha_do_tempo import Evento, registrar_fonte
from app.services.termos import ROTULO_TIPO

ROTULO_EVENTO_LINHA = {
    "ativacao": "Ativação",
    "troca_aparelho": "Troca de aparelho",
    "troca_usuario": "Troca de usuário",
    "portabilidade": "Portabilidade",
    "suspensao": "Suspensão",
    "reativacao": "Reativação",
    "cancelamento": "Cancelamento",
}

_ESTADOS_FINAIS = {
    "recusado": ("termo_recusado", "Termo de {r} recusado"),
    "expirado": ("termo_expirado", "Termo de {r} expirado"),
    "erro": ("termo_erro", "Assinatura recebida, mas a confirmação falhou"),
}


def _data_hora_br(valor) -> str:
    return valor.strftime("%d/%m/%Y %H:%M")


@registrar_fonte
def fonte_termos(item_id: int) -> list[Evento]:
    with get_cursor(commit=False) as cur:
        cur.execute(
            "SELECT t.*, c.nome AS colaborador_nome FROM termo_solicitacoes t "
            "LEFT JOIN colaboradores c ON c.id = t.colaborador_id "
            "WHERE t.item_id = %s ORDER BY t.id",
            (item_id,),
        )
        termos = list(cur.fetchall())

    eventos: list[Evento] = []
    for t in termos:
        r = ROTULO_TIPO.get(t["tipo"], t["tipo"])
        nome = t.get("colaborador_nome") or t.get("destinatario_nome")

        def evento(quando, tipo, titulo, descricao=None, usuario=None, _t=t):
            eventos.append(
                Evento(
                    ocorrido_em=quando, origem="termo", tipo=tipo, titulo=titulo,
                    descricao=descricao, usuario=usuario, referencia_id=_t["id"],
                )
            )

        if t["agendado_para"]:
            descricao = f"Agendado para {_data_hora_br(t['agendado_para'])}"
        elif t["status"] == "manual":
            descricao = t["motivo"]
        else:
            descricao = None
        evento(
            t["solicitado_em"], "termo_solicitado", f"Termo de {r} solicitado",
            descricao, t["solicitado_por"],
        )
        if t["enviado_em"]:
            evento(
                t["enviado_em"], "termo_enviado", f"Termo de {r} enviado para assinatura",
                f"Destinatário: {nome}" if nome else None,
            )
        if t["assinado_em"]:
            evento(t["assinado_em"], "termo_assinado", f"Termo de {r} assinado", nome)
        if t["cancelado_em"]:
            evento(t["cancelado_em"], "termo_cancelado", "Solicitação de termo cancelada", t["motivo"])
        if t["status"] in _ESTADOS_FINAIS:
            tipo, titulo = _ESTADOS_FINAIS[t["status"]]
            evento(t["atualizado_em"], tipo, titulo.format(r=r), t["motivo"])
    return eventos


@registrar_fonte
def fonte_linha_telefonica(item_id: int) -> list[Evento]:
    eventos: list[Evento] = []
    for e in LinhaDBManager().eventos_do_aparelho(item_id):
        if e["tipo"] == "edicao":
            continue  # dados da linha, não do aparelho
        num = formatar_numero_linha(e["linha_numero"])
        if e["tipo"] == "troca_aparelho" and e.get("item_id_anterior") == item_id:
            titulo = f"Linha {num} saiu deste aparelho"
        elif e["tipo"] == "troca_aparelho" and e.get("item_id") == item_id:
            titulo = f"Linha {num} passou para este aparelho"
        else:
            titulo = f"Linha {num}: {ROTULO_EVENTO_LINHA.get(e['tipo'], e['tipo'])}"
        if e.get("de_valor") or e.get("para_valor"):
            descricao = f"{e.get('de_valor') or '—'} → {e.get('para_valor') or '—'}"
        else:
            descricao = e.get("observacao")
        eventos.append(
            Evento(
                ocorrido_em=e["ocorrido_em"], origem="linha", tipo=f"linha_{e['tipo']}",
                titulo=titulo, descricao=descricao, usuario=e.get("por_usuario"),
                referencia_id=e["id"],
            )
        )
    return eventos


@registrar_fonte
def fonte_inventario(item_id: int) -> list[Evento]:
    eventos: list[Evento] = []
    for e in InventarioDBManager().historico_do_item(item_id):
        if e["data"] is None:
            continue
        base = f"Inventário #{e['inventario_id']} ({(e['unidade_nome'] or '').upper()})"
        if e["evento"] == "leitura":
            local = e.get("local_encontrado_nome")
            eventos.append(
                Evento(
                    ocorrido_em=e["data"], origem="inventario", tipo="inventario_leitura",
                    titulo=f"{base}: {ROTULOS_CATEGORIA.get(e['resultado'], e['resultado'])}",
                    descricao=f"Local onde foi encontrado: {local}" if local else None,
                    usuario=e.get("por"), referencia_id=e["inventario_id"],
                )
            )
        else:
            eventos.append(
                Evento(
                    ocorrido_em=e["data"], origem="inventario", tipo="inventario_nao_encontrado",
                    titulo=f"{base}: Não encontrado", usuario=e.get("por"),
                    referencia_id=e["inventario_id"],
                )
            )
    return eventos
