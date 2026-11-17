import { useEffect, useState } from 'react'
import { Link } from 'react-router-dom'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import { Plus } from 'lucide-react'
import { listInventarios, type StatusInventario } from '@/api/inventarios'
import { listUnidades } from '@/api/unidades'
import { NovoInventarioDialog } from '@/components/inventario/NovoInventarioDialog'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { Pagination } from '@/components/ui/Pagination'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { ROTULO_STATUS_INVENTARIO } from '@/features/inventario/resultados'
import { usePageSize } from '@/hooks/usePageSize'
import { upper } from '@/lib/format'
import { formatDateTime } from '@/lib/utils'

const filtroClass =
  'h-9 rounded border border-input bg-surface-alt px-3 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring'

const linkClass =
  'rounded-sm text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2'

/** Inventários físicos por unidade: abrir, acompanhar e consultar relatórios. */
export default function InventariosPage() {
  const [pageSize, setPageSize] = usePageSize('inventarios')
  const [unidadeId, setUnidadeId] = useState('')
  const [status, setStatus] = useState('')
  const [dialogAberto, setDialogAberto] = useState(false)

  const chaveFiltros = JSON.stringify([unidadeId, status])
  const [paginaDosFiltros, setPaginaDosFiltros] = useState({ page: 1, chave: chaveFiltros })
  if (paginaDosFiltros.chave !== chaveFiltros) setPaginaDosFiltros({ page: 1, chave: chaveFiltros })
  const page = paginaDosFiltros.chave === chaveFiltros ? paginaDosFiltros.page : 1
  const setPage = (p: number) => setPaginaDosFiltros({ page: p, chave: chaveFiltros })

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades', { includeInactive: false }],
    queryFn: () => listUnidades(),
  })

  const { data, isLoading, isError, error, isFetching, refetch } = useQuery({
    queryKey: ['inventarios', 'lista', { unidadeId, status, page, pageSize }],
    queryFn: () =>
      listInventarios({
        unidade_id: unidadeId ? Number(unidadeId) : undefined,
        status: (status || undefined) as StatusInventario | undefined,
        limit: pageSize,
        offset: (page - 1) * pageSize,
      }),
    placeholderData: keepPreviousData,
  })
  const items = data?.items ?? []
  const total = data?.total ?? 0

  useEffect(() => {
    if (!data) return
    const ultima = Math.max(1, Math.ceil(data.total / pageSize))
    if (page > ultima) setPage(ultima)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, page, pageSize])

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Inventário"
        eyebrowDetail="Contagem física por unidade"
        title="Inventários"
        description="Confira os equipamentos de cada unidade lendo o QR das etiquetas, e acompanhe o que falta encontrar."
        actions={
          <Button onClick={() => setDialogAberto(true)}>
            <Plus aria-hidden="true" />
            Novo inventário
          </Button>
        }
      />

      <div className="surface-panel flex flex-wrap items-end gap-3 p-3">
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Unidade
          <select
            aria-label="Unidade"
            className={`${filtroClass} min-w-[200px]`}
            value={unidadeId}
            onChange={(e) => setUnidadeId(e.target.value)}
          >
            <option value="">Todas</option>
            {unidades.map((u) => (
              <option key={u.id} value={u.id}>
                {upper(u.nome)}
              </option>
            ))}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Status
          <select aria-label="Status" className={filtroClass} value={status} onChange={(e) => setStatus(e.target.value)}>
            <option value="">Todos</option>
            {(Object.keys(ROTULO_STATUS_INVENTARIO) as StatusInventario[]).map((s) => (
              <option key={s} value={s}>
                {ROTULO_STATUS_INVENTARIO[s]}
              </option>
            ))}
          </select>
        </label>
      </div>

      {isLoading ? (
        <LoadingState label="Carregando inventários..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : items.length === 0 ? (
        <EmptyState
          title="Nenhum inventário encontrado"
          description="Abra um novo inventário para começar a contagem de uma unidade."
        />
      ) : (
        <div className="surface-panel overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="border-b border-border bg-surface-alt">
                <tr>
                  {['Nº', 'Unidade', 'Status', 'Aberto por', 'Aberto em', 'Encerrado em', 'Leituras', ''].map((h, i) => (
                    <th key={h || i} className="whitespace-nowrap px-4 py-2.5 text-left text-caption text-muted-foreground">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {items.map((inv) => (
                  <tr
                    key={inv.id}
                    className="border-b border-border transition-colors duration-micro last:border-0 hover:bg-surface-alt"
                  >
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm font-mono tabular-nums">#{inv.id}</td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{upper(inv.unidade_nome)}</td>
                    <td className="whitespace-nowrap px-4 py-2.5">
                      <StatusPill status={ROTULO_STATUS_INVENTARIO[inv.status]} />
                    </td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{inv.iniciado_por}</td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{formatDateTime(inv.iniciado_em)}</td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{formatDateTime(inv.encerrado_em)}</td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm tabular-nums">{inv.total_leituras}</td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">
                      <div className="flex items-center gap-4">
                        {inv.status === 'aberto' && (
                          <Link to={`/inventarios/${inv.id}/leitura`} className={linkClass}>
                            Ler
                          </Link>
                        )}
                        <Link to={`/inventarios/${inv.id}`} className={linkClass}>
                          Relatório
                        </Link>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="border-t border-border px-4 py-3">
            <Pagination
              page={page}
              pageSize={pageSize}
              total={total}
              isFetching={isFetching}
              onPageChange={setPage}
              onPageSizeChange={(n) => {
                setPageSize(n)
                setPage(1)
              }}
            />
          </div>
        </div>
      )}

      <NovoInventarioDialog open={dialogAberto} onOpenChange={setDialogAberto} />
    </div>
  )
}
