import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import type { ColumnDef } from '@tanstack/react-table'
import { listItemsPaginated, type Item } from '@/api/items'
import { DataTable } from '@/components/ui/DataTable'
import { StatusBadge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { ItemDetailsModal } from '@/components/equipment/ItemDetailsModal'
import { useAuth } from '@/contexts/AuthContext'
import { useConstants } from '@/hooks/useConstants'
import { formatDate } from '@/lib/utils'
import { Plus, Pencil, RefreshCw, Package, CheckCircle2, Clock, Laptop, Eye, Loader2 } from 'lucide-react'

const STATUS_OPTIONS = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução']
const FETCH_ALL_LIMIT = 500

export default function StockPage() {
  const { hasRole } = useAuth()
  const navigate = useNavigate()
  const { equipmentTypes, isLoading: constantsLoading } = useConstants()
  const [filterTipo, setFilterTipo] = useState('all')
  const [filterStatus, setFilterStatus] = useState('all')
  const [isRefreshing, setIsRefreshing] = useState(false)
  const [selectedItem, setSelectedItem] = useState<Item | null>(null)

  const { data, isLoading, refetch } = useQuery({
    queryKey: ['items', filterTipo, filterStatus],
    queryFn: () =>
      listItemsPaginated({
        tipo: filterTipo !== 'all' ? filterTipo : undefined,
        status: filterStatus !== 'all' ? filterStatus : undefined,
        limit: FETCH_ALL_LIMIT,
      }),
  })

  const items = data?.items ?? []
  const total = data?.total ?? 0

  const disponiveisCount = items.filter((i) => i.status === 'Disponível').length
  const indisponiveisCount = items.filter((i) => i.status === 'Indisponível').length
  const pendentesCount = items.filter((i) => i.status?.startsWith('Pendente')).length

  const handleRefresh = async () => {
    setIsRefreshing(true)
    await refetch()
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
    {
      accessorKey: 'peripheral_count',
      header: 'Periféricos',
      size: 90,
      cell: ({ getValue }) => <span className="num">{getValue() as number}</span>,
    },
    { accessorKey: 'assigned_to', header: 'Usuário Alocado', cell: ({ getValue }) => (getValue() as string) || '-' },
    { accessorKey: 'revenda', header: 'Unidade' },
    {
      accessorKey: 'identificador',
      header: 'Identificador',
      cell: ({ getValue }) => <span className="num">{(getValue() as string) || '-'}</span>,
    },
    { accessorKey: 'setor', header: 'Setor', cell: ({ getValue }) => (getValue() as string) || '-' },
    {
      accessorKey: 'ip',
      header: 'IP',
      cell: ({ getValue }) => <span className="num">{(getValue() as string) || '-'}</span>,
    },
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
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-border pb-4">
        <div>
          <h2 className="text-heading text-foreground">
            Estoque de Equipamentos
          </h2>
          <p className="text-body-sm text-muted-foreground">
            Gerenciamento centralizado de hardware e insumos de TI
          </p>
        </div>
        <div className="flex items-center gap-3">
          <Button variant="outline" size="sm" onClick={handleRefresh}>
            <RefreshCw size={14} className={isRefreshing ? 'animate-spin' : ''} />
            Atualizar
          </Button>
          {hasRole('Gestor', 'Técnico') && (
            <Button variant="gradient" size="sm" onClick={() => navigate('/register')}>
              <Plus size={16} />
              Novo Equipamento
            </Button>
          )}
        </div>
      </div>

      {/* KPI Metric Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="surface-panel p-5 flex items-center gap-4">
          <div className="h-12 w-12 rounded-md bg-surface-alt border border-border text-foreground flex items-center justify-center">
            <Package size={24} />
          </div>
          <div>
            <p className="text-caption text-muted-foreground">Total em Estoque</p>
            <p className="text-heading-lg num text-foreground">{total}</p>
          </div>
        </div>

        <div className="surface-panel p-5 flex items-center gap-4">
          <div className="h-12 w-12 rounded-md bg-surface-alt border border-border text-foreground flex items-center justify-center">
            <CheckCircle2 size={24} />
          </div>
          <div>
            <p className="text-caption text-muted-foreground">Disponíveis</p>
            <p className="text-heading-lg num text-foreground">{disponiveisCount}</p>
          </div>
        </div>

        <div className="surface-panel p-5 flex items-center gap-4">
          <div className="h-12 w-12 rounded-md bg-surface-alt border border-border text-foreground flex items-center justify-center">
            <Laptop size={24} />
          </div>
          <div>
            <p className="text-caption text-muted-foreground">Em Empréstimo</p>
            <p className="text-heading-lg num text-foreground">{indisponiveisCount}</p>
          </div>
        </div>

        <div className="surface-panel p-5 flex items-center gap-4">
          <div className="h-12 w-12 rounded-md bg-surface-alt border border-border text-foreground flex items-center justify-center">
            <Clock size={24} />
          </div>
          <div>
            <p className="text-caption text-muted-foreground">Ações Pendentes</p>
            <p className="text-heading-lg num text-foreground">{pendentesCount}</p>
          </div>
        </div>
      </div>

      {/* Filter Bar */}
      <div className="surface-panel p-3 flex flex-wrap items-center gap-3">
        <span className="text-caption text-muted-foreground px-2">Filtros:</span>
        <Select value={filterTipo} onValueChange={setFilterTipo} disabled={constantsLoading}>
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

        <Select value={filterStatus} onValueChange={setFilterStatus}>
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

        {(filterTipo !== 'all' || filterStatus !== 'all') && (
          <Button
            variant="ghost"
            size="sm"
            onClick={() => {
              setFilterTipo('all')
              setFilterStatus('all')
            }}
            className="text-xs text-muted-foreground hover:text-foreground"
          >
            Limpar Filtros
          </Button>
        )}
      </div>

      {/* Data Table */}
      {isLoading ? (
        <div className="surface-panel p-12 flex flex-col items-center gap-3">
          <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
          <p className="text-body-sm text-muted-foreground">Carregando estoque...</p>
        </div>
      ) : (
        <DataTable
          data={items}
          columns={columns}
          onRowClick={(item) => setSelectedItem(item)}
          searchPlaceholder="Buscar por patrimônio, marca, modelo, usuário, setor, IP..."
        />
      )}

      {/* Modal de Detalhes do Equipamento */}
      <ItemDetailsModal item={selectedItem} onClose={() => setSelectedItem(null)} />
    </div>
  )
}


