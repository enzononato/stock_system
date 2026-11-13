import { useEffect, useState } from 'react'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import { Plus, X } from 'lucide-react'
import { useSearchParams } from 'react-router-dom'
import { listManutencoes, type StatusManutencao } from '@/api/manutencoes'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Pagination } from '@/components/ui/Pagination'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { AbrirManutencaoDialog } from '@/features/manutencoes/AbrirManutencaoDialog'
import { ManutencoesLista } from '@/features/manutencoes/ManutencoesLista'
import { ROTULO_STATUS_MANUTENCAO } from '@/features/manutencoes/status'
import { useConstants } from '@/hooks/useConstants'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { usePageSize } from '@/hooks/usePageSize'
import { formatMoeda, upper } from '@/lib/format'

const filtroClass =
  'h-9 rounded border border-input bg-surface-alt px-3 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring'

type FiltroStatus = '' | StatusManutencao | 'em_andamento'

export default function ManutencoesPage() {
  const { revendas, equipmentTypes } = useConstants()
  const [searchParams, setSearchParams] = useSearchParams()
  const itemIdParam = searchParams.get('item_id')
  const itemId = itemIdParam && /^\d+$/.test(itemIdParam) ? Number(itemIdParam) : undefined

  const [pageSize, setPageSize] = usePageSize('manutencoes')
  const [busca, setBusca] = useState('')
  const [status, setStatus] = useState<FiltroStatus>('')
  const [revenda, setRevenda] = useState('')
  const [tipo, setTipo] = useState('')
  const [de, setDe] = useState('')
  const [ate, setAte] = useState('')
  const [novaAberta, setNovaAberta] = useState(false)
  const buscaAplicada = useDebouncedValue(busca.trim(), 300)

  // A página vale só para o conjunto de filtros em que foi escolhida: ao mudar qualquer
  // filtro (ou o tamanho da página) volta para a 1 na mesma renderização, sem uma
  // requisição com filtro novo e página antiga.
  const filtroKey = JSON.stringify([buscaAplicada, status, revenda, tipo, de, ate, itemId ?? null, pageSize])
  const [paginaDoFiltro, setPaginaDoFiltro] = useState({ page: 1, key: filtroKey })
  if (paginaDoFiltro.key !== filtroKey) setPaginaDoFiltro({ page: 1, key: filtroKey })
  const page = paginaDoFiltro.key === filtroKey ? paginaDoFiltro.page : 1
  const setPage = (p: number) => setPaginaDoFiltro({ page: p, key: filtroKey })

  const periodoInvalido = Boolean(de && ate && de > ate)

  const { data, isLoading, isError, error, isFetching, refetch } = useQuery({
    queryKey: [
      'manutencoes',
      { busca: buscaAplicada, status, revenda, tipo, de, ate, itemId, page, pageSize },
    ],
    queryFn: () =>
      listManutencoes({
        status: status || undefined,
        item_id: itemId,
        revenda: revenda || undefined,
        tipo: tipo || undefined,
        de: de || undefined,
        ate: ate || undefined,
        busca: buscaAplicada || undefined,
        limit: pageSize,
        offset: (page - 1) * pageSize,
      }),
    placeholderData: keepPreviousData,
    enabled: !periodoInvalido,
  })

  const total = data?.total ?? 0
  const items = data?.items ?? []

  // Se a lista encolheu e a página atual deixou de existir, volta para a última.
  useEffect(() => {
    if (!data) return
    const lastPage = Math.max(1, Math.ceil(data.total / pageSize))
    if (page > lastPage) setPage(lastPage)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, page, pageSize])

  function limparEquipamento() {
    const next = new URLSearchParams(searchParams)
    next.delete('item_id')
    setSearchParams(next, { replace: true })
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Operações"
        eyebrowDetail="Ciclo de vida"
        title="Manutenções"
        description="Ordens de manutenção dos equipamentos, com custo por ordem e por equipamento."
        actions={
          <Button type="button" size="sm" onClick={() => setNovaAberta(true)}>
            <Plus className="mr-1.5 size-3.5" />
            Nova manutenção
          </Button>
        }
      />

      <div className="surface-panel flex flex-wrap items-end gap-3 p-3">
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Buscar
          <Input
            aria-label="Buscar manutenções"
            placeholder="Marca, modelo, identificador, assistência ou defeito"
            value={busca}
            onChange={(e) => setBusca(e.target.value)}
            className="w-72"
          />
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Status
          <select
            aria-label="Status"
            className={filtroClass}
            value={status}
            onChange={(e) => setStatus(e.target.value as FiltroStatus)}
          >
            <option value="">Todas</option>
            <option value="em_andamento">Em andamento</option>
            {(Object.keys(ROTULO_STATUS_MANUTENCAO) as StatusManutencao[]).map((s) => (
              <option key={s} value={s}>
                {ROTULO_STATUS_MANUTENCAO[s]}
              </option>
            ))}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Unidade
          <select
            aria-label="Unidade"
            className={filtroClass}
            value={revenda}
            onChange={(e) => setRevenda(e.target.value)}
          >
            <option value="">Todas</option>
            {revendas.map((r) => (
              <option key={r} value={r}>
                {upper(r)}
              </option>
            ))}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Tipo
          <select aria-label="Tipo" className={filtroClass} value={tipo} onChange={(e) => setTipo(e.target.value)}>
            <option value="">Todos</option>
            {equipmentTypes.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Aberta de
          <input
            type="date"
            aria-label="Aberta de"
            className={filtroClass}
            value={de}
            max={ate || undefined}
            onChange={(e) => setDe(e.target.value)}
          />
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          até
          <input
            type="date"
            aria-label="até"
            className={filtroClass}
            value={ate}
            min={de || undefined}
            onChange={(e) => setAte(e.target.value)}
          />
        </label>
        {itemId !== undefined && (
          <span className="inline-flex h-9 items-center gap-1.5 rounded-full border border-border bg-surface px-3 text-sm text-foreground">
            {`Equipamento #${itemId}`}
            <button
              type="button"
              aria-label="Limpar filtro de equipamento"
              onClick={limparEquipamento}
              className="rounded-full p-0.5 text-muted-foreground hover:text-foreground"
            >
              <X className="size-3.5" aria-hidden />
            </button>
          </span>
        )}
      </div>

      {data && (
        <p className="text-body-sm font-medium text-foreground">
          {`${total} ${total === 1 ? 'manutenção' : 'manutenções'} · Custo total: ${formatMoeda(data.custo_total)}`}
        </p>
      )}

      {periodoInvalido ? (
        <p role="alert" className="text-body-sm text-destructive">
          Período inválido: a data inicial é maior que a final.
        </p>
      ) : isLoading ? (
        <LoadingState label="Carregando manutenções..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : items.length === 0 ? (
        <EmptyState
          title="Nenhuma manutenção encontrada"
          description="Ajuste os filtros ou abra uma nova manutenção."
        />
      ) : (
        <div className="space-y-3">
          <ManutencoesLista manutencoes={items} mostrarEquipamento />
          <Pagination
            page={page}
            pageSize={pageSize}
            total={total}
            isFetching={isFetching}
            onPageChange={setPage}
            onPageSizeChange={setPageSize}
          />
        </div>
      )}

      <AbrirManutencaoDialog open={novaAberta} onOpenChange={setNovaAberta} />
    </div>
  )
}
