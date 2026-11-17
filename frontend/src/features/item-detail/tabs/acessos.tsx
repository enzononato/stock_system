import { Link } from 'react-router-dom'
import { useQuery } from '@tanstack/react-query'
import { listAcessos } from '@/api/acessos'
import type { Item } from '@/api/items'
import { STATUS_ACESSO } from '@/components/acessos/rotulos'
import { StatusPill } from '@/components/ui/StatusPill'
import { upper } from '@/lib/format'
import type { ItemDetailTab } from '../types'

function AcessosTab({ item }: { item: Item }) {
  const { data, isLoading, isError } = useQuery({
    queryKey: ['acessos', 'ficha', { item_id: item.id }],
    queryFn: () => listAcessos({ item_id: item.id }),
  })

  if (isLoading) return <p className="text-sm text-muted-foreground">Carregando acessos...</p>
  if (isError) return <p className="text-sm text-muted-foreground">Não foi possível carregar os acessos.</p>
  const acessos = data ?? []
  if (acessos.length === 0) {
    return <p className="text-sm text-muted-foreground">Nenhum acesso vinculado a este equipamento.</p>
  }

  return (
    <ul className="divide-y divide-border rounded-lg border border-border">
      {acessos.map((a) => {
        const st = STATUS_ACESSO[a.status]
        return (
          <li key={a.id} className="flex flex-wrap items-center justify-between gap-2 px-4 py-3">
            <div className="min-w-0">
              <p className="text-sm">
                <span className="font-semibold">{a.sistema_nome}</span>
                <span className="text-muted-foreground"> · </span>
                <span>{a.login}</span>
              </p>
              <Link
                to={`/colaboradores/${a.colaborador_id}`}
                className="rounded-sm text-xs font-semibold text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
              >
                {upper(a.colaborador_nome)}
              </Link>
            </div>
            <StatusPill status={st.label} tone={st.tone} />
          </li>
        )
      })}
    </ul>
  )
}

const tab: ItemDetailTab = {
  id: 'acessos',
  label: 'Acessos',
  order: 60,
  roles: ['Gestor', 'Técnico'],
  Component: AcessosTab,
}
export default tab
