import { useEffect, useState } from 'react'
import { useQuery, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import type { ColumnDef } from '@tanstack/react-table'
import { listItemsPaginated, type Item } from '@/api/items'
import { listUnidades } from '@/api/unidades'
import { DataTable } from '@/components/ui/DataTable'
import { StatusBadge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { ItemDetailsModal } from '@/components/equipment/ItemDetailsModal'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { useAuth } from '@/contexts/AuthContext'
import { useConstants } from '@/hooks/useConstants'
import { formatDate } from '@/lib/utils'
import { Plus, Pencil, RefreshCw, Eye, Loader2, X } from 'lucide-react'

const STATUS_OPTIONS = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução']
const PAGE_SIZE = 10

export default function StockPage() {
  const { hasRole } = useAuth()
  const navigate = useNavigate()
  const queryClient = useQueryClient()
  const { equipmentTypes, isLoading: constantsLoading } = useConstants()
  const [filterTipo, setFilterTipo] = useState('all')
  const [filterStatus, setFilterStatus] = useState('all')
  const [filterRevenda, setFilterRevenda] = useState('all')
  const [pageIndex, setPageIndex] = useState(0)
  // `search` é o que o usuário digita; `searchAplicado` é o que vai para a
  // query, com atraso, para não disparar uma requisição por tecla (mesmo
  // padrão do HistoryPage).
  const [search, setSearch] = useState('')
  const [searchAplicado, setSearchAplicado] = useState('')
  const [isRefreshing, setIsRefreshing] = useState(false)
  const [selectedItem, setSelectedItem] = useState<Item | null>(null)

  const hasFilters =
    filterTipo !== 'all' || filterStatus !== 'all' || filterRevenda !== 'all' || searchAplicado !== ''

  // Espera o usuário parar de digitar antes de consultar o servidor, e volta
  // para a primeira página (mesmo padrão do HistoryPage): o resultado da nova
  // busca não tem relação com a página em que ele estava.
  useEffect(() => {
    const timer = setTimeout(() => {
      setSearchAplicado(search.trim())
      setPageIndex(0)
    }, 400)
    return () => clearTimeout(timer)
  }, [search])

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades-filter'],
    queryFn: () => listUnidades(),
  })

  // Lista paginada de verdade: 10 itens por página, filtros e busca resolvidos
  // no SQL (backend/app/db/inventory_manager_db.py). Antes disto a tela
  // buscava `limit: 500` (o teto do backend, settings.MAX_PAGE_SIZE) e
  // renderizava tudo — acima de 500 itens ativos a tabela truncava em
  // silêncio, sem aviso nenhum ao usuário.
  const { data, isLoading, isFetching } = useQuery({
    queryKey: ['items', 'list', filterTipo, filterStatus, filterRevenda, pageIndex, searchAplicado],
    queryFn: () =>
      listItemsPaginated({
        tipo: filterTipo !== 'all' ? filterTipo : undefined,
        status: filterStatus !== 'all' ? filterStatus : undefined,
        revenda: filterRevenda !== 'all' ? filterRevenda : undefined,
        search: searchAplicado || undefined,
        limit: PAGE_SIZE,
        offset: pageIndex * PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })

  const items = data?.items ?? []
  const total = data?.total ?? 0

  function handleTipoChange(value: string) {
    setFilterTipo(value)
    setPageIndex(0)
  }

  function handleStatusChange(value: string) {
    setFilterStatus(value)
    setPageIndex(0)
  }

  function handleRevendaChange(value: string) {
    setFilterRevenda(value)
    setPageIndex(0)
  }

  function handleClearFilters() {
    setFilterTipo('all')
    setFilterStatus('all')
    setFilterRevenda('all')
    setSearch('')
    setSearchAplicado('')
    setPageIndex(0)
  }

  const handleRefresh = async () => {
    setIsRefreshing(true)
    // Invalida a lista paginada e todos os cartões de indicador de uma vez —
    // todas as queries desta tela usam o prefixo 'items'.
    await queryClient.invalidateQueries({ queryKey: ['items'] })
    setTimeout(() => setIsRefreshing(false), 500)
  }

  const columns: ColumnDef<Item, unknown>[] = [
    {
      accessorKey: 'id',
      header: 'ID',
      size: 60,
      cell: ({ row }) => (
        <button
          onClick={() => setSelectedItem(row.original)}
          className="num font-semibold text-foreground hover:underline text-body-sm"
          title="Clique para ver detalhes do item"
        >
          #{row.original.id}
        </button>
      ),
    },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'model', header: 'Modelo' },
    {
      accessorKey: 'status',
      header: 'Status',
      cell: ({ row }) => <StatusBadge status={row.original.status} />,
    },
    { accessorKey: 'assigned_to', header: 'Usuário Alocado', cell: ({ getValue }) => (getValue() as string) || '-' },
    { accessorKey: 'revenda', header: 'Unidade' },
    {
      accessorKey: 'date_registered',
      header: 'Data Cadastro',
      cell: ({ getValue }) => <span className="num">{formatDate(getValue() as string)}</span>,
    },
    {
      id: 'actions',
      header: '',
      cell: ({ row }: { row: { original: Item } }) => (
        <div className="flex items-center gap-1">
          <Button
            variant="ghost"
            size="sm"
            onClick={() => setSelectedItem(row.original)}
            className="h-8 w-8 p-0 text-muted-foreground hover:text-foreground hover:bg-surface-alt rounded"
            title="Ver detalhes do equipamento"
          >
            <Eye size={14} />
          </Button>
          {hasRole('Gestor', 'Técnico') && (
            <Button
              variant="ghost"
              size="sm"
              onClick={() => navigate(`/edit/${row.original.id}`)}
              className="h-8 w-8 p-0 text-muted-foreground hover:text-foreground hover:bg-surface-alt rounded"
              title="Editar equipamento"
            >
              <Pencil size={14} />
            </Button>
          )}
        </div>
      ),
    },
  ]

  return (
    <div className="space-y-6">
      {/* Header Bar */}
      <PageHeader
        eyebrow="Inventário"
        eyebrowDetail="Painel Operacional de Equipamentos"
        title={
          <span className="inline-flex items-baseline gap-3">
            Estoque de Equipamentos
            {/* Substitui os antigos cartões de indicador (Total, Disponíveis, Em
                Empréstimo, Ações Pendentes): a contagem já vem de graça do
                `total` da resposta paginada, sem requisição extra — a
                referência (origin/redesign-frontend) mostra o mesmo texto no
                lugar dos cartões, que ela mantém só no Dashboard. */}
            <span className="text-caption font-mono num text-muted-foreground">({total} ativos registrados)</span>
          </span>
        }
        description="Gerenciamento centralizado de hardware e insumos de TI"
        actions={
          <>
            <Button variant="outline" size="sm" onClick={handleRefresh}>
              <RefreshCw size={14} className={isRefreshing || isFetching ? 'animate-spin' : ''} />
              Atualizar
            </Button>
            {hasRole('Gestor', 'Técnico') && (
              <Button variant="gradient" size="sm" onClick={() => navigate('/register')}>
                <Plus size={16} />
                Novo Equipamento
              </Button>
            )}
          </>
        }
      />

      {/* Filter Bar */}
      <div className="surface-panel p-3 space-y-3">
        <PanelHeader title="Filtros de Consulta" />
        <div className="flex flex-wrap items-center gap-3">
          <Select value={filterTipo} onValueChange={handleTipoChange} disabled={constantsLoading}>
            <SelectTrigger className="w-48">
              <SelectValue placeholder="Tipo de Equipamento" />
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

          <Select value={filterStatus} onValueChange={handleStatusChange}>
            <SelectTrigger className="w-56">
              <SelectValue placeholder="Status do Item" />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all">Todos os status</SelectItem>
              {STATUS_OPTIONS.map((s) => (
                <SelectItem key={s} value={s}>
                  {s}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>

          <Select value={filterRevenda} onValueChange={handleRevendaChange}>
            <SelectTrigger className="w-48">
              <SelectValue placeholder="Unidade" />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all">Todas as unidades</SelectItem>
              {unidades.map((u) => (
                <SelectItem key={u.id} value={u.nome}>
                  {u.nome}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>

          {hasFilters && (
            <Button
              variant="ghost"
              size="sm"
              onClick={handleClearFilters}
              className="text-xs text-muted-foreground hover:text-foreground"
            >
              <X size={14} />
              Limpar Filtros
            </Button>
          )}
        </div>
      </div>

      {/* Data Table */}
      {isLoading ? (
        <div className="surface-panel p-12 flex flex-col items-center gap-3">
          <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
          <p className="text-body-sm text-muted-foreground">Carregando estoque...</p>
        </div>
      ) : (
        <>
          {/* Estado vazio distingue "não há item cadastrado" (frota realmente
              vazia) de "nenhum resultado para este filtro" (há itens, mas os
              filtros/busca atuais não encontram nenhum) — o segundo sugere
              limpar os filtros. Renderizado além da DataTable (não no lugar
              dela) para não tirar do usuário a busca/paginação já em uso. */}
          {items.length === 0 && (
            <div className="surface-panel p-6 flex flex-col items-center gap-2 text-center">
              <p className="text-body-sm text-foreground font-medium">
                {hasFilters ? 'Nenhum resultado para este filtro.' : 'Nenhum item cadastrado.'}
              </p>
              <p className="text-body-sm text-muted-foreground">
                {hasFilters
                  ? 'Ajuste os filtros ou a busca para encontrar o que procura.'
                  : 'Cadastre um equipamento para começar a usar o estoque.'}
              </p>
              {hasFilters && (
                <Button variant="outline" size="sm" onClick={handleClearFilters}>
                  <X size={14} />
                  Limpar Filtros
                </Button>
              )}
            </div>
          )}
          <DataTable
            data={items}
            columns={columns}
            onRowClick={(item) => setSelectedItem(item)}
            searchPlaceholder="Buscar por patrimônio, marca, modelo, usuário, setor, IP..."
            pagination={{
              total,
              pageIndex,
              pageSize: PAGE_SIZE,
              onPageChange: setPageIndex,
              search,
              onSearchChange: setSearch,
            }}
          />
        </>
      )}

      {/* Modal de Detalhes do Equipamento */}
      <ItemDetailsModal item={selectedItem} onClose={() => setSelectedItem(null)} />
    </div>
  )
}
