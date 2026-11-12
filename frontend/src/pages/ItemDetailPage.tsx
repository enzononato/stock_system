import { useMemo } from 'react'
import { Link, useParams, useSearchParams } from 'react-router-dom'
import { useQuery } from '@tanstack/react-query'
import { isAxiosError } from 'axios'
import { ArrowLeft } from 'lucide-react'
import { getItem } from '@/api/items'
import { DetailTabs } from '@/components/ui/DetailTabs'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { useAuth } from '@/contexts/AuthContext'
import { getItemDetailTabs } from '@/features/item-detail/registry'
import { upper } from '@/lib/format'

export default function ItemDetailPage() {
  const { id } = useParams<{ id: string }>()
  const itemId = Number(id)
  const { user } = useAuth()
  const [searchParams, setSearchParams] = useSearchParams()

  const tabs = useMemo(() => getItemDetailTabs(user?.role), [user?.role])
  const activeTab = tabs.find((t) => t.id === searchParams.get('aba')) ?? tabs[0]

  const { data: item, isLoading, error } = useQuery({
    queryKey: ['item', itemId],
    queryFn: () => getItem(itemId),
    enabled: Number.isFinite(itemId),
    retry: (count, err) => !(isAxiosError(err) && err.response?.status === 404) && count < 1,
  })

  const voltar = (
    <Link
      to="/stock"
      className="inline-flex items-center gap-1.5 rounded-sm text-sm font-medium text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
    >
      <ArrowLeft className="size-4" aria-hidden />
      Voltar ao estoque
    </Link>
  )

  if (isLoading) return <LoadingState label="Carregando equipamento..." />

  const notFound = !Number.isFinite(itemId) || (isAxiosError(error) && error.response?.status === 404)
  if (notFound) {
    return (
      <div className="mx-auto max-w-5xl space-y-4">
        {voltar}
        <EmptyState
          title="Equipamento não encontrado"
          description="O equipamento pode ter sido removido ou o endereço está incorreto."
        />
      </div>
    )
  }
  if (error || !item) return <ErrorState error={error} />

  function selectTab(tabId: string) {
    const next = new URLSearchParams(searchParams)
    next.set('aba', tabId)
    setSearchParams(next, { replace: true })
  }

  const Active = activeTab?.Component

  return (
    <div className="mx-auto max-w-5xl space-y-6">
      {voltar}

      <header className="space-y-1 border-b border-border pb-4">
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
            {item.tipo}
          </span>
          <StatusPill status={item.status} />
          {item.revenda && (
            <span className="text-caption font-medium text-muted-foreground">{upper(item.revenda)}</span>
          )}
        </div>
        <h1 className="text-heading-lg font-semibold tracking-tight text-foreground">
          {[item.brand, item.model].filter(Boolean).join(' ')}
        </h1>
        {item.identificador && (
          <p className="font-mono text-sm text-muted-foreground">{item.identificador}</p>
        )}
      </header>

      <DetailTabs tabs={tabs} activeId={activeTab?.id} onSelect={selectTab} label="Seções do equipamento">
        {Active && <Active item={item} />}
      </DetailTabs>
    </div>
  )
}
