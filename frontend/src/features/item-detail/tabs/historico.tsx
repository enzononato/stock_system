import { useEffect, useState } from 'react'
import { useQuery, keepPreviousData } from '@tanstack/react-query'
import type { Item } from '@/api/items'
import { listHistoryPaginated } from '@/api/history'
import { Pagination } from '@/components/ui/Pagination'
import { usePageSize } from '@/hooks/usePageSize'
import { formatDateTime } from '@/lib/utils'
import type { ItemDetailTab } from '../types'

function HistoricoTab({ item }: { item: Item }) {
  const [pageSize, setPageSize] = usePageSize('item-historico')
  const [page, setPage] = useState(1)

  // Volta para a primeira página ao trocar de equipamento ou de tamanho de página.
  useEffect(() => {
    setPage(1)
  }, [item.id, pageSize])

  const { data, isLoading, isError, isFetching } = useQuery({
    queryKey: ['item-history', item.id, page, pageSize],
    queryFn: () => listHistoryPaginated({ item_id: item.id, limit: pageSize, offset: (page - 1) * pageSize }),
    placeholderData: keepPreviousData,
  })

  const entries = data?.items ?? []
  const total = data?.total ?? 0
  const lastPage = Math.max(1, Math.ceil(total / pageSize))
  const current = Math.min(page, lastPage)

  if (isLoading) return <p className="text-sm text-muted-foreground">Carregando histórico...</p>
  if (isError) return <p className="text-sm text-muted-foreground">Não foi possível carregar o histórico.</p>
  if (total === 0) {
    return <p className="text-sm text-muted-foreground">Nenhuma movimentação registrada para este equipamento.</p>
  }

  return (
    <div className="space-y-3">
      <ul className="divide-y divide-border rounded-lg border border-border">
        {entries.map((e) => (
          <li key={e.id} className="flex flex-wrap items-start justify-between gap-2 px-4 py-3">
            <div className="min-w-0">
              <p className="text-sm font-semibold text-foreground">{e.operation ?? '—'}</p>
              <p className="text-xs text-muted-foreground">
                {[e.usuario && `Usuário: ${e.usuario}`, e.operador && `Operador: ${e.operador}`]
                  .filter(Boolean)
                  .join(' · ') || '—'}
              </p>
              {e.details && <p className="mt-1 text-xs text-muted-foreground">{e.details}</p>}
            </div>
            <time className="shrink-0 text-xs text-muted-foreground">{formatDateTime(e.data_operacao)}</time>
          </li>
        ))}
      </ul>
      <Pagination
        page={current}
        pageSize={pageSize}
        total={total}
        onPageChange={setPage}
        onPageSizeChange={setPageSize}
        isFetching={isFetching}
      />
    </div>
  )
}

const tab: ItemDetailTab = { id: 'historico', label: 'Histórico', order: 90, Component: HistoricoTab }
export default tab
