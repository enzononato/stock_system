import type { EquipamentosPorUnidade as Unidade } from '@/api/painel'
import { StatusPill } from '@/components/ui/StatusPill'
import { upper } from '@/lib/format'

const COLUNAS_STATUS = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução', 'Em Manutenção']

/** Equipamentos ativos por unidade e status. A tabela rola na horizontal em telas estreitas. */
export function EquipamentosPorUnidade({ unidades }: { unidades: Unidade[] }) {
  if (unidades.length === 0) {
    return <p className="py-6 text-center text-body-sm text-muted-foreground">Nenhum equipamento cadastrado.</p>
  }
  return (
    <div className="overflow-x-auto">
      <table className="w-full min-w-[34rem] text-left text-sm">
        <thead>
          <tr className="border-b border-border">
            <th scope="col" className="py-2 pr-3 text-caption font-semibold text-muted-foreground">
              Unidade
            </th>
            <th scope="col" className="px-2 py-2 text-right text-caption font-semibold text-muted-foreground">
              Total
            </th>
            {COLUNAS_STATUS.map((status) => (
              <th key={status} scope="col" className="px-2 py-2 text-right">
                <StatusPill status={status} />
              </th>
            ))}
          </tr>
        </thead>
        <tbody className="divide-y divide-border">
          {unidades.map((u) => (
            <tr key={u.unidade ?? '__sem_unidade__'}>
              <th scope="row" className="py-2 pr-3 font-medium text-foreground">
                {u.unidade ? upper(u.unidade) : 'SEM UNIDADE'}
              </th>
              <td className="px-2 py-2 text-right font-mono font-semibold text-foreground">{u.total}</td>
              {COLUNAS_STATUS.map((status) => {
                const n = u.por_status[status] ?? 0
                return (
                  <td
                    key={status}
                    className={`px-2 py-2 text-right font-mono ${n === 0 ? 'text-muted-foreground' : 'text-foreground'}`}
                  >
                    {n}
                  </td>
                )
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
