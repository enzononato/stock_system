import { useEffect, useState } from 'react'
import { useQuery, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { Link, useNavigate } from 'react-router-dom'
import { type Item } from '@/api/items'
import { listUnidades } from '@/api/unidades'
import { ColaboradorSelect } from '@/components/colaboradores/ColaboradorSelect'
import { ItemDetailsModal } from '@/components/equipment/ItemDetailsModal'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Pagination } from '@/components/ui/Pagination'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import { StatusPill } from '@/components/ui/StatusPill'
import { useAuth } from '@/contexts/AuthContext'
import {
  listEstoque,
  proximaOrdenacao,
  type EstoqueSort,
  type SortDir,
} from '@/features/estoque/api'
import { SortableHeader } from '@/features/estoque/SortableHeader'
import { useConstants } from '@/hooks/useConstants'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { usePageSize } from '@/hooks/usePageSize'
import { upper } from '@/lib/format'
import { formatDate } from '@/lib/utils'
import { Eye, Pencil, Plus, RefreshCw, Search, X } from 'lucide-react'

// "Em Manutenção" é criado pela F4; até lá o filtro só volta vazio.
const STATUS_OPTIONS = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução', 'Em Manutenção']

const COLUNAS: { label: string; coluna: EstoqueSort }[] = [
  { label: 'Patrimônio', coluna: 'id' },
  { label: 'Tipo', coluna: 'tipo' },
  { label: 'Equipamento', coluna: 'equipamento' },
  { label: 'Status', coluna: 'status' },
  { label: 'Colaborador', coluna: 'colaborador' },
  { label: 'Setor', coluna: 'setor' },
  { label: 'Cargo', coluna: 'cargo' },
  { label: 'Unidade', coluna: 'unidade' },
  { label: 'Cadastro', coluna: 'cadastro' },
]
const TOTAL_COLUNAS = COLUNAS.length + 1

export default function StockPage() {
  const { hasRole } = useAuth()
  const navigate = useNavigate()
  const queryClient = useQueryClient()
  const { equipmentTypes, isLoading: constantsLoading } = useConstants()

  const [search, setSearch] = useState('')
  const busca = useDebouncedValue(search.trim(), 400)
  const [filterTipo, setFilterTipo] = useState('all')
  const [filterStatus, setFilterStatus] = useState('all')
  // Nome original da unidade (o texto em CAIXA ALTA é só a exibição).
  const [filterRevenda, setFilterRevenda] = useState('all')
  const [colaboradorId, setColaboradorId] = useState<number | null>(null)
  const [ordem, setOrdem] = useState<{ sort: EstoqueSort; dir: SortDir }>({ sort: 'id', dir: 'desc' })
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = usePageSize('estoque')
  const [selectedItem, setSelectedItem] = useState<Item | null>(null)

  // Nova busca (já com atraso) volta para a primeira página.
  useEffect(() => {
    setPage(1)
  }, [busca])

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades-filter'],
    queryFn: () => listUnidades(),
  })

  // Lista paginada de verdade: filtros, busca e ordenação resolvidos no SQL.
  const { data, isLoading, isFetching, error, refetch } = useQuery({
    queryKey: [
      'items',
      'estoque',
      { tipo: filterTipo, status: filterStatus, revenda: filterRevenda, colaboradorId, busca, sort: ordem.sort, dir: ordem.dir, page, pageSize },
    ],
    queryFn: () =>
      listEstoque({
        tipo: filterTipo !== 'all' ? filterTipo : undefined,
        status: filterStatus !== 'all' ? filterStatus : undefined,
        revenda: filterRevenda !== 'all' ? filterRevenda : undefined,
        colaborador_id: colaboradorId ?? undefined,
        search: busca || undefined,
        sort: ordem.sort,
        dir: ordem.dir,
        limit: pageSize,
        offset: (page - 1) * pageSize,
      }),
    placeholderData: keepPreviousData,
  })

  const items = data?.items ?? []
  const total = data?.total ?? 0
  const lastPage = Math.max(1, Math.ceil(total / pageSize))
  const hasFilters =
    filterTipo !== 'all' ||
    filterStatus !== 'all' ||
    filterRevenda !== 'all' ||
    colaboradorId !== null ||
    search !== ''

  // Total encolheu (filtro, baixa feita por outra pessoa): volta para a última página válida.
  useEffect(() => {
    if (total > 0 && page > lastPage) setPage(lastPage)
  }, [total, page, lastPage])

  function mudarFiltro<T>(setter: (v: T) => void) {
    return (v: T) => {
      setter(v)
      setPage(1)
    }
  }

  function resetFilters() {
    setSearch('')
    setFilterTipo('all')
    setFilterStatus('all')
    setFilterRevenda('all')
    setColaboradorId(null)
    setPage(1)
  }

  function handleSort(coluna: EstoqueSort) {
    setOrdem((atual) => proximaOrdenacao(atual, coluna))
    setPage(1)
  }

  function handleRefresh() {
    // Invalida tudo que usa o prefixo 'items' (lista paginada e afins).
    void queryClient.invalidateQueries({ queryKey: ['items'] })
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Inventário"
        title="Estoque de Equipamentos"
        description={`${total} ativos registrados`}
        actions={
          <>
            <Button variant="outline" size="sm" onClick={handleRefresh} loading={isFetching}>
              {!isFetching && <RefreshCw />}
              Atualizar
            </Button>
            {hasRole('Gestor', 'Técnico') && (
              <Button size="sm" onClick={() => navigate('/register')}>
                <Plus />
                Novo Ativo
              </Button>
            )}
          </>
        }
      />

      {/* Filtros */}
      <div className="rounded-md border border-border bg-surface p-5">
        <div className="flex flex-col gap-3 lg:flex-row lg:flex-wrap lg:items-center">
          <div className="relative w-full lg:max-w-md lg:flex-1">
            <Search className="absolute left-3 top-1/2 size-3.5 -translate-y-1/2 text-muted-foreground" />
            <Input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              aria-label="Buscar equipamentos"
              placeholder="Buscar por patrimônio, marca, modelo, usuário ou serial..."
              className="h-9 pl-9 text-xs"
            />
            {search && (
              <button
                type="button"
                onClick={() => setSearch('')}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
                title="Limpar busca"
                aria-label="Limpar busca"
              >
                <X className="size-3.5" />
              </button>
            )}
          </div>

          <div className="grid grid-cols-1 gap-2 sm:flex sm:flex-wrap sm:items-center">
            <Select value={filterTipo} onValueChange={mudarFiltro(setFilterTipo)} disabled={constantsLoading}>
              <SelectTrigger aria-label="Tipo" className="h-9 w-full text-xs sm:w-40">
                <SelectValue placeholder="Tipo" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todos os tipos</SelectItem>
                {equipmentTypes.map((t) => (
                  <SelectItem key={t} value={t}>
                    {t}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>

            <Select value={filterStatus} onValueChange={mudarFiltro(setFilterStatus)}>
              <SelectTrigger aria-label="Status" className="h-9 w-full text-xs sm:w-40">
                <SelectValue placeholder="Status" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todos status</SelectItem>
                {STATUS_OPTIONS.map((s) => (
                  <SelectItem key={s} value={s}>
                    {s}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>

            <Select value={filterRevenda} onValueChange={mudarFiltro(setFilterRevenda)}>
              <SelectTrigger aria-label="Unidade" className="h-9 w-full text-xs sm:w-44">
                <SelectValue placeholder="Unidade" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todas as unidades</SelectItem>
                {unidades.map((u) => (
                  <SelectItem key={u.id} value={u.nome}>
                    {upper(u.nome)}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>

            <div className="w-full sm:w-72">
              <ColaboradorSelect
                value={colaboradorId}
                onChange={(c) => mudarFiltro(setColaboradorId)(c?.id ?? null)}
                somenteAtivos={false}
                placeholder="Todos os colaboradores"
              />
            </div>

            {hasFilters && (
              <Button variant="ghost" size="sm" onClick={resetFilters} className="text-muted-foreground hover:text-foreground">
                <X /> Limpar
              </Button>
            )}
          </div>
        </div>
      </div>

      {/* Tabela */}
      <div className="rounded-md border border-border bg-surface">
        <div className="overflow-x-auto">
          <table className="w-full min-w-[960px] text-left">
            <thead className="bg-brand-soft text-brand-soft-foreground uppercase tracking-wider">
              <tr>
                {COLUNAS.map((c) => (
                  <SortableHeader
                    key={c.coluna}
                    label={c.label}
                    coluna={c.coluna}
                    sort={ordem.sort}
                    dir={ordem.dir}
                    onSort={handleSort}
                  />
                ))}
                <th scope="col" className="px-4 py-3">
                  <span className="sr-only">Ações</span>
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-border">
              {isLoading ? (
                <tr>
                  <td colSpan={TOTAL_COLUNAS} className="py-16 text-center text-caption text-muted-foreground">
                    Carregando equipamentos…
                  </td>
                </tr>
              ) : error ? (
                <tr>
                  <td colSpan={TOTAL_COLUNAS} className="py-16 text-center text-caption text-destructive">
                    Erro ao carregar.{' '}
                    <button type="button" onClick={() => void refetch()} className="underline">
                      Tentar novamente
                    </button>
                  </td>
                </tr>
              ) : items.length === 0 ? (
                <tr>
                  <td colSpan={TOTAL_COLUNAS} className="py-16 text-center text-caption text-muted-foreground">
                    {hasFilters
                      ? 'Nenhum patrimônio corresponde aos filtros aplicados.'
                      : 'Não há equipamentos cadastrados no estoque.'}
                  </td>
                </tr>
              ) : (
                items.map((row) => (
                  <tr
                    key={row.id}
                    className="group cursor-pointer transition-colors hover:bg-surface-alt"
                    onClick={() => setSelectedItem(row)}
                  >
                    <td className="px-4 py-2.5">
                      <Link
                        to={`/items/${row.id}`}
                        onClick={(e) => e.stopPropagation()}
                        className="num flex items-center gap-1 font-mono text-xs font-semibold text-brand-text hover:underline"
                      >
                        <span>#{row.id}</span>
                        {row.identificador && (
                          <span className="font-normal text-muted-foreground">({row.identificador})</span>
                        )}
                      </Link>
                    </td>
                    <td className="px-4 py-2.5">
                      <span className="block max-w-[140px] truncate text-body-sm font-medium text-foreground">
                        {row.tipo || '—'}
                      </span>
                    </td>
                    <td className="px-4 py-2.5">
                      <span
                        className="block max-w-[200px] truncate text-body-sm text-foreground"
                        title={[row.brand, row.model].filter(Boolean).join(' ')}
                      >
                        {row.brand} {row.model || '—'}
                      </span>
                    </td>
                    <td className="px-4 py-2.5">
                      <StatusPill status={row.status} />
                    </td>
                    <td className="px-4 py-2.5">
                      <span className="block max-w-[160px] truncate text-body-sm text-muted-foreground">
                        {row.colaborador_nome || row.assigned_to || '—'}
                      </span>
                    </td>
                    <td className="px-4 py-2.5">
                      <span className="block max-w-[140px] truncate text-body-sm text-muted-foreground">
                        {row.colaborador_setor || '—'}
                      </span>
                    </td>
                    <td className="px-4 py-2.5">
                      <span className="block max-w-[140px] truncate text-body-sm text-muted-foreground">
                        {row.colaborador_cargo || '—'}
                      </span>
                    </td>
                    <td className="px-4 py-2.5">
                      <span className="block max-w-[140px] truncate text-xs text-muted-foreground">
                        {upper(row.revenda) || '—'}
                      </span>
                    </td>
                    <td className="px-4 py-2.5">
                      <span className="num font-mono text-xs text-muted-foreground">
                        {formatDate(row.date_registered)}
                      </span>
                    </td>
                    <td className="px-4 py-2.5">
                      {/* Visíveis no toque; só somem até o hover em telas largas. */}
                      <div className="flex items-center justify-end gap-1 opacity-100 transition-opacity lg:opacity-0 lg:focus-within:opacity-100 lg:group-hover:opacity-100">
                        <Button
                          variant="ghost"
                          size="sm"
                          className="size-8 p-0 text-muted-foreground hover:text-foreground"
                          onClick={(e) => {
                            e.stopPropagation()
                            setSelectedItem(row)
                          }}
                          title="Ver ficha técnica"
                          aria-label={`Ver ficha técnica do #${row.id}`}
                        >
                          <Eye />
                        </Button>
                        {hasRole('Gestor', 'Técnico') && (
                          <Button
                            variant="ghost"
                            size="sm"
                            className="size-8 p-0 text-muted-foreground hover:text-foreground"
                            onClick={(e) => {
                              e.stopPropagation()
                              navigate(`/edit/${row.id}`)
                            }}
                            title="Editar equipamento"
                            aria-label={`Editar equipamento #${row.id}`}
                          >
                            <Pencil />
                          </Button>
                        )}
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        <div className="border-t border-border p-4">
          <Pagination
            page={page}
            pageSize={pageSize}
            total={total}
            onPageChange={setPage}
            onPageSizeChange={(n) => {
              setPageSize(n)
              setPage(1)
            }}
            isFetching={isFetching}
          />
        </div>
      </div>

      {/* Modal de detalhe */}
      <ItemDetailsModal item={selectedItem} onClose={() => setSelectedItem(null)} />
    </div>
  )
}
