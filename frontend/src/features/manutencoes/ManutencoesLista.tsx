import { Link } from 'react-router-dom'
import type { Manutencao } from '@/api/manutencoes'
import { StatusPill } from '@/components/ui/StatusPill'
import { isoParaBR } from '@/features/ciclo-vida/format'
import { formatMoeda, upper } from '@/lib/format'
import { formatDateTime } from '@/lib/utils'
import { ManutencaoAcoes } from './ManutencaoAcoes'
import { ROTULO_STATUS_MANUTENCAO, TOM_STATUS_MANUTENCAO } from './status'

interface Props {
  manutencoes: Manutencao[]
  /** Mostra equipamento, unidade e custo acumulado (telas que misturam equipamentos). */
  mostrarEquipamento: boolean
}

// Em telas estreitas cada linha vira um cartão: as células empilham, com o rótulo vindo de `data-label`.
const celula =
  'flex items-start justify-between gap-3 px-3 py-1.5 before:shrink-0 before:text-[10px] before:font-semibold before:uppercase before:tracking-wider before:text-muted-foreground before:content-[attr(data-label)] md:table-cell md:py-3 md:before:hidden'

/** Tabela de manutenções (cartões em telas estreitas). */
export function ManutencoesLista({ manutencoes, mostrarEquipamento }: Props) {
  return (
    <div className="rounded-lg border border-border">
      <table className="block w-full text-sm md:table">
        <thead className="hidden md:table-header-group">
          <tr className="border-b border-border text-left text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
            <th className="px-3 py-2">Abertura</th>
            {mostrarEquipamento && <th className="px-3 py-2">Equipamento</th>}
            <th className="px-3 py-2">Defeito</th>
            <th className="px-3 py-2">Assistência</th>
            <th className="px-3 py-2">Previsão</th>
            <th className="px-3 py-2">Status</th>
            <th className="px-3 py-2">Custo</th>
            {mostrarEquipamento && <th className="px-3 py-2">Custo acumulado</th>}
            <th className="px-3 py-2">Ações</th>
          </tr>
        </thead>
        <tbody className="block divide-y divide-border md:table-row-group">
          {manutencoes.map((m) => (
            <tr key={m.id} className="block py-2 md:table-row md:py-0">
              <td data-label="Abertura" className={celula}>
                {formatDateTime(m.aberta_em)}
              </td>
              {mostrarEquipamento && (
                <td data-label="Equipamento" className={celula}>
                  <div className="min-w-0 text-right md:text-left">
                    <Link
                      to={`/items/${m.item_id}?aba=manutencoes`}
                      className="font-medium text-brand-text hover:underline"
                    >
                      {[m.item_tipo, m.item_brand, m.item_model].filter(Boolean).join(' ') || `#${m.item_id}`}
                    </Link>
                    {m.item_identificador && (
                      <p className="font-mono text-xs text-muted-foreground">{m.item_identificador}</p>
                    )}
                    {m.item_revenda && <p className="text-xs text-muted-foreground">{upper(m.item_revenda)}</p>}
                  </div>
                </td>
              )}
              <td data-label="Defeito" className={celula}>
                {m.defeito}
              </td>
              <td data-label="Assistência" className={celula}>
                {m.assistencia ?? '—'}
              </td>
              <td data-label="Previsão" className={celula}>
                {isoParaBR(m.previsao_retorno) || '—'}
              </td>
              <td data-label="Status" className={celula}>
                <StatusPill status={ROTULO_STATUS_MANUTENCAO[m.status]} tone={TOM_STATUS_MANUTENCAO[m.status]} />
              </td>
              <td data-label="Custo" className={celula}>
                {formatMoeda(m.custo) || '—'}
              </td>
              {mostrarEquipamento && (
                <td data-label="Custo acumulado" className={celula}>
                  {formatMoeda(m.custo_acumulado_item) || '—'}
                </td>
              )}
              <td data-label="Ações" className={celula}>
                <ManutencaoAcoes manutencao={m} />
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
