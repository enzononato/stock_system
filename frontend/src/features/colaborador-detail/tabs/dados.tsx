import type { ReactNode } from 'react'
import { Link } from 'react-router-dom'
import type { Colaborador } from '@/api/colaboradores'
import { formatTelefone, upper } from '@/lib/format'
import { formatDateTime } from '@/lib/utils'
import type { ColaboradorDetailTab } from '../types'

const VAZIO = '—'

function DadosTab({ colaborador: c }: { colaborador: Colaborador }) {
  let superior: ReactNode = 'Não informado'
  if (c.superior_nome) {
    superior =
      c.superior_id != null ? (
        <Link
          to={`/colaboradores/${c.superior_id}`}
          className="rounded-sm text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
        >
          {upper(c.superior_nome)}
        </Link>
      ) : (
        upper(c.superior_nome)
      )
  }

  const linhas: { label: string; value: ReactNode }[] = [
    { label: 'Matrícula', value: c.matricula || VAZIO },
    { label: 'E-mail', value: c.email || VAZIO },
    { label: 'Telefone', value: formatTelefone(c.telefone) || VAZIO },
    { label: 'Superior imediato', value: superior },
    { label: 'Origem', value: c.origem === 'rh' ? 'RH' : 'Manual' },
    { label: 'Última sincronização', value: c.sincronizado_em ? formatDateTime(c.sincronizado_em) : VAZIO },
  ]

  return (
    <dl className="grid grid-cols-1 gap-x-6 gap-y-3 sm:grid-cols-2 lg:grid-cols-3">
      {linhas.map((l) => (
        <div key={l.label} className="min-w-0">
          <dt className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">{l.label}</dt>
          <dd className="mt-0.5 break-words text-sm font-medium text-foreground">{l.value}</dd>
        </div>
      ))}
    </dl>
  )
}

const tab: ColaboradorDetailTab = { id: 'dados', label: 'Dados', order: 10, Component: DadosTab }
export default tab
