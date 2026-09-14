import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import type { ColumnDef } from '@tanstack/react-table'
import { listItemsPaginated, type Item } from '@/api/items'
import { DataTable } from '@/components/ui/DataTable'
import { StatusBadge } from '@/components/ui/badge'
import { StatBlock } from '@/components/ui/StatBlock'
import { Button } from '@/components/ui/button'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { ItemDetailsModal } from '@/components/equipment/ItemDetailsModal'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { useAuth } from '@/contexts/AuthContext'
import { useConstants } from '@/hooks/useConstants'
import { formatDate } from '@/lib/utils'
import { Plus, Pencil, RefreshCw, Eye, Loader2 } from 'lucide-react'

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
      <PageHeader
        eyebrow="Inventário"
        eyebrowDetail="Painel Operacional de Equipamentos"
        title="Estoque de Equipamentos"
        description="Gerenciamento centralizado de hardware e insumos de TI"
        actions={
          <>
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
          </>
        }
      />

      {/* KPI Metric Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="surface-panel p-4">
          <StatBlock label="Total em Estoque" value={total} symbol="#" />
        </div>

        <div className="surface-panel p-4">
          <StatBlock label="Disponíveis" value={disponiveisCount} symbol="○" />
        </div>

        <div className="surface-panel p-4">
          <StatBlock label="Em Empréstimo" value={indisponiveisCount} />
        </div>

        <div className="surface-panel p-4">
          <StatBlock label="Ações Pendentes" value={pendentesCount} symbol="!" />
        </div>
      </div>

      {/* Filter Bar */}
      <div className="surface-panel p-3 space-y-3">
        <PanelHeader title="Filtros de Consulta" />
        <div className="flex flex-wrap items-center gap-3">
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


