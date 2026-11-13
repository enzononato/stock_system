import { Link } from 'react-router-dom'
import { TIPO_TERMO_ROTULO, type TermoSolicitacao } from '@/api/termos'
import { formatDateTime } from '@/lib/utils'
import { TermoStatusPill } from './TermoStatusPill'

interface TermosListaProps {
  termos: TermoSolicitacao[]
  /** Mostra a coluna do equipamento (usada na aba do colaborador). */
  mostrarItem?: boolean
}

export function TermosLista({ termos, mostrarItem = false }: TermosListaProps) {
  return (
    <div className="overflow-x-auto rounded-lg border border-border">
      <table className="w-full text-sm">
        <thead className="bg-surface-alt text-left text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
          <tr>
            {mostrarItem && <th className="px-3 py-2">Equipamento</th>}
            <th className="px-3 py-2">Tipo</th>
            <th className="px-3 py-2">Status</th>
            <th className="px-3 py-2">Solicitado por / em</th>
            <th className="px-3 py-2">Enviado em</th>
            <th className="px-3 py-2">Assinado em</th>
            <th className="px-3 py-2">Motivo</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-border">
          {termos.map((t, i) => (
            <tr key={t.id ?? `legado-${i}`}>
              {mostrarItem && (
                <td className="px-3 py-2">
                  <Link
                    to={`/items/${t.item_id}?aba=termos`}
                    className="rounded-sm font-semibold text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring"
                  >
                    {t.item_descricao || `Equipamento ${t.item_id}`}
                  </Link>
                </td>
              )}
              <td className="whitespace-nowrap px-3 py-2">{TIPO_TERMO_ROTULO[t.tipo]}</td>
              <td className="px-3 py-2">
                <TermoStatusPill status={t.status} />
              </td>
              <td className="px-3 py-2">
                <span className="block">{t.solicitado_por}</span>
                <span className="block text-xs text-muted-foreground">{formatDateTime(t.solicitado_em)}</span>
              </td>
              <td className="whitespace-nowrap px-3 py-2">{t.enviado_em ? formatDateTime(t.enviado_em) : '—'}</td>
              <td className="whitespace-nowrap px-3 py-2">{t.assinado_em ? formatDateTime(t.assinado_em) : '—'}</td>
              <td className="px-3 py-2 text-muted-foreground">{t.motivo || '—'}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
