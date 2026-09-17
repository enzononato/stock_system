import { useEffect, useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import { listItemsPaginated, type Item } from '@/api/items'
import { listUnidades } from '@/api/unidades'
import { ItemDetailsModal } from '@/components/equipment/ItemDetailsModal'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import { useAuth } from '@/contexts/AuthContext'
import { useConstants } from '@/hooks/useConstants'
import { cn, formatDate } from '@/lib/utils'
import { Eye, Pencil, Plus, RefreshCw, Search, X } from 'lucide-react'

const STATUS_OPTIONS = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução']
const PAGE_SIZE = 7

function MonoBadge({ status }: { status?: string }) {
  if (!status) return <span className="text-muted-foreground text-xs">—</span>
  let symbol = '●'
  if (status === 'Disponível') symbol = '○'
  else if (status.startsWith('Pendente')) symbol = '!'
  else if (status === 'Baixado' || status === 'Descartado') symbol = '×'
  return (
    <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-sm border border-border text-foreground font-mono text-xs font-medium">
      <span className="font-bold text-[11px]">{symbol}</span>
      <span className="uppercase tracking-wider text-[11px]">{status}</span>
    </span>
  )
}

export default function StockPage() {
  const { hasRole } = useAuth()
  const navigate = useNavigate()
  const { equipmentTypes } = useConstants()

  const [search, setSearch] = useState('')
  const [filterTipo, setFilterTipo] = useState('all')
  const [filterStatus, setFilterStatus] = useState('all')
  const [filterRevenda, setFilterRevenda] = useState('all')
  const [page, setPage] = useState(0)
  const [selectedItem, setSelectedItem] = useState<Item | null>(null)

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades-filter'],
    queryFn: () => listUnidades(),
  })

  const { data, isLoading, isFetching, error, refetch } = useQuery({
    queryKey: ['items', filterTipo, filterStatus, filterRevenda, page, search],
    queryFn: () =>
      listItemsPaginated({
        tipo: filterTipo !== 'all' ? filterTipo : undefined,
        status: filterStatus !== 'all' ? filterStatus : undefined,
        revenda: filterRevenda !== 'all' ? filterRevenda : undefined,
        search: search.trim() || undefined,
        limit: PAGE_SIZE,
        offset: page * PAGE_SIZE,
      }),
  })

  const items = data?.items ?? []
  const total = data?.total ?? 0
  const hasFilters =
    filterTipo !== 'all' || filterStatus !== 'all' || filterRevenda !== 'all' || search !== ''
  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE))

  useEffect(() => {
    if (data && page > 0 && items.length === 0 && total > 0) {
      setPage(Math.max(0, Math.ceil(total / PAGE_SIZE) - 1))
    }
  }, [data, items.length, page, total])

  function resetFilters() {
    setSearch('')
    setFilterTipo('all')
    setFilterStatus('all')
    setFilterRevenda('all')
    setPage(0)
  }

  return (
    <div className="space-y-6">
      {/* Cabeçalho Operacional */}
      <div className="border-b border-border pb-5">
        <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-3">
          <div>
            <span className="text-caption text-muted-foreground font-mono">
              INVENTÁRIO · DATA GRID OPERACIONAL
            </span>
            <div className="flex items-baseline gap-3 mt-0.5">
              <h1 className="text-heading-lg font-semibold tracking-tight text-foreground">
                Estoque de Equipamentos
              </h1>
              <span className="text-caption font-mono num text-muted-foreground">
                ({total} ativos registrados)
              </span>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => void refetch()}
              disabled={isFetching}
              className="h-8 text-xs"
            >
              <RefreshCw className={cn('mr-1.5 size-3.5', isFetching && 'animate-spin')} />
              {isFetching ? 'Atualizando…' : 'Atualizar'}
            </Button>
            {hasRole('Gestor', 'Técnico') && (
              <Button
                size="sm"
                onClick={() => navigate('/register')}
                className="h-8 text-xs"
              >
                <Plus className="mr-1.5 size-3.5" />
                Novo Ativo
              </Button>
            )}
          </div>
        </div>
      </div>

      {/* Toolbar de Filtros */}
      <div className="bg-surface border border-border rounded-md p-4 space-y-3">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
          {/* Busca */}
          <div className="relative flex-1 max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
            <Input
              value={search}
              onChange={(e) => {
                setSearch(e.target.value)
                setPage(0)
              }}
              placeholder="Buscar por patrimônio, marca, modelo, usuário ou serial..."
              className="h-9 pl-9 text-xs"
            />
            {search && (
              <button
                type="button"
                onClick={() => setSearch('')}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
              >
                <X className="size-3.5" />
              </button>
            )}
          </div>

          {/* Filtros em linha */}
          <div className="flex flex-wrap items-center gap-2">
            <Select value={filterTipo} onValueChange={(v) => { setFilterTipo(v); setPage(0) }}>
              <SelectTrigger className="h-9 w-36 text-xs">
                <SelectValue placeholder="Tipo" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todos os tipos</SelectItem>
                {equipmentTypes.map((t) => (
                  <SelectItem key={t} value={t}>{t}</SelectItem>
                ))}
              </SelectContent>
            </Select>

            <Select value={filterStatus} onValueChange={(v) => { setFilterStatus(v); setPage(0) }}>
              <SelectTrigger className="h-9 w-36 text-xs">
                <SelectValue placeholder="Status" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todos status</SelectItem>
                {STATUS_OPTIONS.map((s) => (
                  <SelectItem key={s} value={s}>{s}</SelectItem>
                ))}
              </SelectContent>
            </Select>

            <Select value={filterRevenda} onValueChange={(v) => { setFilterRevenda(v); setPage(0) }}>
              <SelectTrigger className="h-9 w-40 text-xs">
                <SelectValue placeholder="Unidade" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todas as unidades</SelectItem>
                {unidades.map((u) => (
                  <SelectItem key={u.id} value={u.nome}>{u.nome}</SelectItem>
                ))}
              </SelectContent>
            </Select>

            {hasFilters && (
              <Button
                variant="ghost"
                size="sm"
                onClick={resetFilters}
                className="h-9 px-2 text-xs text-muted-foreground hover:text-foreground"
              >
                <X className="mr-1 size-3" /> Limpar
              </Button>
            )}
          </div>
        </div>
      </div>

      {/* Tabela de Alta Densidade */}
      <div className="bg-surface border border-border rounded-md overflow-hidden">
        <table className="w-full text-left">
          <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground bg-surface-alt">
            <tr>
              <th className="py-2.5 px-4 font-semibold">Patrimônio</th>
              <th className="py-2.5 px-4 font-semibold">Tipo</th>
              <th className="py-2.5 px-4 font-semibold">Equipamento</th>
              <th className="py-2.5 px-4 font-semibold">Status</th>
              <th className="py-2.5 px-4 font-semibold hidden md:table-cell">Alocado Para</th>
              <th className="py-2.5 px-4 font-semibold hidden lg:table-cell">Unidade</th>
              <th className="py-2.5 px-4 font-semibold hidden xl:table-cell">Cadastro</th>
              <th className="py-2.5 px-4" />
            </tr>
          </thead>
          <tbody className="divide-y divide-border">
            {isLoading ? (
              <tr>
                <td colSpan={8} className="py-16 text-center text-caption text-muted-foreground">
                  Carregando equipamentos…
                </td>
              </tr>
            ) : error ? (
              <tr>
                <td colSpan={8} className="py-16 text-center text-caption text-destructive">
                  Erro ao carregar. <button onClick={() => void refetch()} className="underline">Tentar novamente</button>
                </td>
              </tr>
            ) : items.length === 0 ? (
              <tr>
                <td colSpan={8} className="py-16 text-center text-caption text-muted-foreground">
                  {hasFilters
                    ? 'Nenhum patrimônio corresponde aos filtros aplicados.'
                    : 'Não há equipamentos cadastrados no estoque.'}
                </td>
              </tr>
            ) : (
              items.map((row) => (
                <tr
                  key={row.id}
                  className="hover:bg-surface-alt cursor-pointer transition-colors group"
                  onClick={() => setSelectedItem(row)}
                >
                  <td className="py-2.5 px-4">
                    <button
                      type="button"
                      className="num font-mono text-xs font-semibold text-foreground hover:underline flex items-center gap-1"
                      onClick={(e) => { e.stopPropagation(); setSelectedItem(row) }}
                    >
                      <span>#{row.id}</span>
                      {row.identificador && (
                        <span className="text-muted-foreground font-normal">({row.identificador})</span>
                      )}
                    </button>
                  </td>
                  <td className="py-2.5 px-4">
                    <span className="font-medium text-foreground truncate max-w-[140px] block text-body-sm">
                      {row.tipo || '—'}
                    </span>
                  </td>
                  <td className="py-2.5 px-4">
                    <span className="text-foreground truncate max-w-[200px] block text-body-sm" title={`${row.brand} ${row.model}`}>
                      {row.brand} {row.model || '—'}
                    </span>
                  </td>
                  <td className="py-2.5 px-4">
                    <MonoBadge status={row.status} />
                  </td>
                  <td className="py-2.5 px-4 hidden md:table-cell">
                    <span className="text-muted-foreground truncate max-w-[160px] block text-body-sm">
                      {row.assigned_to || '—'}
                    </span>
                  </td>
                  <td className="py-2.5 px-4 hidden lg:table-cell">
                    <span className="text-muted-foreground text-xs truncate max-w-[120px] block">
                      {row.revenda || '—'}
                    </span>
                  </td>
                  <td className="py-2.5 px-4 hidden xl:table-cell">
                    <span className="num font-mono text-xs text-muted-foreground">
                      {formatDate(row.date_registered)}
                    </span>
                  </td>
                  <td className="py-2.5 px-4">
                    <div className="flex items-center justify-end gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                      <Button
                        variant="ghost"
                        size="sm"
                        className="size-8 p-0 text-muted-foreground hover:text-foreground"
                        onClick={(e) => { e.stopPropagation(); setSelectedItem(row) }}
                        title="Ver ficha técnica"
                      >
                        <Eye className="size-3.5" />
                      </Button>
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>

        {/* Paginação */}
        {totalPages > 1 && (
          <div className="flex items-center justify-between border-t border-border px-4 py-3 text-xs text-muted-foreground bg-surface-alt">
            <span>
              Página {page + 1} de {totalPages} · {total} registros
            </span>
            <div className="flex items-center gap-2">
              <Button
                size="sm"
                variant="outline"
                disabled={page === 0}
                onClick={() => setPage((p) => Math.max(0, p - 1))}
                className="h-7 px-2.5 text-xs"
              >
                ← Anterior
              </Button>
              <Button
                size="sm"
                variant="outline"
                disabled={page >= totalPages - 1}
                onClick={() => setPage((p) => p + 1)}
                className="h-7 px-2.5 text-xs"
              >
                Próxima →
              </Button>
            </div>
          </div>
        )}
      </div>

      {/* Modal de detalhe */}
      <ItemDetailsModal item={selectedItem} onClose={() => setSelectedItem(null)} />
    </div>
  )
}
