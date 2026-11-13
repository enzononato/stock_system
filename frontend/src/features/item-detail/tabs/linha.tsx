import { Link } from 'react-router-dom'
import { useQuery } from '@tanstack/react-query'
import type { Item } from '@/api/items'
import { listEventosDoAparelho, listLinhas } from '@/api/linhas'
import { LinhaEventosList } from '@/components/linhas/LinhaEventosList'
import { StatusPill } from '@/components/ui/StatusPill'
import { statusLinhaInfo } from '@/features/linhas/status'
import { formatMoeda, formatTelefone, upper } from '@/lib/format'
import type { ItemDetailTab } from '../types'

const TIPOS_COM_LINHA = ['Celular', 'Tablet']

function LinhaTab({ item }: { item: Item }) {
  if (!item.tipo || !TIPOS_COM_LINHA.includes(item.tipo)) {
    return <p className="text-sm text-muted-foreground">Linhas telefônicas só se aplicam a celulares e tablets.</p>
  }
  return <LinhaDoAparelho item={item} />
}

function LinhaDoAparelho({ item }: { item: Item }) {
  const linhas = useQuery({
    queryKey: ['linhas', 'aparelho', item.id],
    queryFn: () => listLinhas({ item_id: item.id, limit: 25 }),
  })
  const eventos = useQuery({
    queryKey: ['linhas', 'eventos-aparelho', item.id],
    queryFn: () => listEventosDoAparelho(item.id),
  })

  return (
    <div className="space-y-6">
      <section className="space-y-3">
        {linhas.isLoading ? (
          <p className="text-sm text-muted-foreground">Carregando linha...</p>
        ) : linhas.isError ? (
          <p className="text-sm text-muted-foreground">Não foi possível carregar a linha.</p>
        ) : (linhas.data?.items.length ?? 0) === 0 ? (
          <p className="text-sm text-muted-foreground">Nenhuma linha vinculada a este aparelho.</p>
        ) : (
          linhas.data!.items.map((l) => {
            const info = statusLinhaInfo(l.status)
            return (
              <div key={l.id} className="space-y-2 rounded-lg border border-border p-4">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <h4 className="text-heading-sm tabular-nums text-foreground">{formatTelefone(l.numero)}</h4>
                  <StatusPill status={info.label} tone={info.tone} />
                </div>
                <p className="text-sm text-foreground">
                  <span>{[l.operadora, l.plano].filter(Boolean).join(' · ')}</span>
                  <span> · {formatMoeda(l.custo_mensal)}</span>
                </p>
                <p className="flex flex-wrap gap-x-2 text-xs text-muted-foreground">
                  {l.colaborador_nome && <span>{upper(l.colaborador_nome)}</span>}
                  <span>{upper(l.unidade_nome)}</span>
                </p>
                <Link
                  to={`/linhas?busca=${l.numero}`}
                  className="inline-block rounded-sm text-sm font-semibold text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
                >
                  Abrir em Linhas
                </Link>
              </div>
            )
          })
        )}
      </section>

      <section className="space-y-3">
        <h3 className="text-heading-sm text-foreground">Histórico de linhas neste aparelho</h3>
        {eventos.isLoading ? (
          <p className="text-sm text-muted-foreground">Carregando histórico...</p>
        ) : eventos.isError ? (
          <p className="text-sm text-muted-foreground">Não foi possível carregar o histórico de linhas.</p>
        ) : (
          <LinhaEventosList eventos={eventos.data ?? []} mostrarNumero />
        )}
      </section>
    </div>
  )
}

const tab: ItemDetailTab = { id: 'linha', label: 'Linha', order: 50, Component: LinhaTab }
export default tab
