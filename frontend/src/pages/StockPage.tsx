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
import { Plus, Pencil, RefreshCw, Package, CheckCircle2, Clock, Laptop, Eye } from 'lucide-react'

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
          className="font-mono font-bold text-indigo-600 hover:text-indigo-800 hover:underline text-xs"
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
    { accessorKey: 'peripheral_count', header: 'Periféricos', size: 90 },
    { accessorKey: 'assigned_to', header: 'Usuário Alocado', cell: ({ getValue }) => (getValue() as string) || '-' },
    { accessorKey: 'revenda', header: 'Unidade' },
    { accessorKey: 'identificador', header: 'Identificador', cell: ({ getValue }) => (getValue() as string) || '-' },
    { accessorKey: 'setor', header: 'Setor', cell: ({ getValue }) => (getValue() as string) || '-' },
    { accessorKey: 'ip', header: 'IP', cell: ({ getValue }) => (getValue() as string) || '-' },
    {
      accessorKey: 'date_registered',
      header: 'Data Cadastro',
      cell: ({ getValue }) => formatDate(getValue() as string),
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
            className="h-8 w-8 p-0 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg"
            title="Ver detalhes do equipamento"
          >
            <Eye size={14} />
          </Button>
          {hasRole('Gestor', 'Técnico') && (
            <Button
              variant="ghost"
              size="sm"
              onClick={() => navigate(`/edit/${row.original.id}`)}
              className="h-8 w-8 p-0 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg"
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
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold tracking-tight text-slate-900 font-heading">
            Estoque de Equipamentos
          </h2>
          <p className="text-sm text-slate-500 font-medium">
            Gerenciamento centralizado de hardware e insumos de TI
          </p>
        </div>
        <div className="flex items-center gap-3">
          <Button
            variant="outline"
            size="sm"
            onClick={handleRefresh}
            className="rounded-xl border-slate-200 bg-white shadow-sm"
          >
            <RefreshCw size={14} className={`mr-2 ${isRefreshing ? 'animate-spin text-indigo-600' : ''}`} />
            Atualizar
          </Button>
          {hasRole('Gestor', 'Técnico') && (
            <Button
              variant="gradient"
              size="sm"
              onClick={() => navigate('/register')}
              className="rounded-xl"
            >
              <Plus size={16} className="mr-1.5" />
              Novo Equipamento
            </Button>
          )}
        </div>
      </div>

      {/* KPI Metric Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="glass-card rounded-2xl p-4 flex items-center gap-4">
          <div className="h-12 w-12 rounded-xl bg-indigo-500/10 text-indigo-600 flex items-center justify-center font-bold shadow-inner">
            <Package size={24} />
          </div>
          <div>
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Total em Estoque</p>
            <p className="text-2xl font-bold text-slate-900 font-heading">{total}</p>
          </div>
        </div>

        <div className="glass-card rounded-2xl p-4 flex items-center gap-4">
          <div className="h-12 w-12 rounded-xl bg-emerald-500/10 text-emerald-600 flex items-center justify-center font-bold shadow-inner">
            <CheckCircle2 size={24} />
          </div>
          <div>
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Disponíveis</p>
            <p className="text-2xl font-bold text-emerald-600 font-heading">{disponiveisCount}</p>
          </div>
        </div>

        <div className="glass-card rounded-2xl p-4 flex items-center gap-4">
          <div className="h-12 w-12 rounded-xl bg-sky-500/10 text-sky-600 flex items-center justify-center font-bold shadow-inner">
            <Laptop size={24} />
          </div>
          <div>
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Em Empréstimo</p>
            <p className="text-2xl font-bold text-sky-600 font-heading">{indisponiveisCount}</p>
          </div>
        </div>

        <div className="glass-card rounded-2xl p-4 flex items-center gap-4">
          <div className="h-12 w-12 rounded-xl bg-amber-500/10 text-amber-600 flex items-center justify-center font-bold shadow-inner">
            <Clock size={24} />
          </div>
          <div>
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Ações Pendentes</p>
            <p className="text-2xl font-bold text-amber-600 font-heading">{pendentesCount}</p>
          </div>
        </div>
      </div>

      {/* Filter Bar */}
      <div className="flex flex-wrap items-center gap-3 p-3 rounded-2xl bg-white/70 backdrop-blur-md border border-slate-200/80 shadow-sm">
        <span className="text-xs font-bold text-slate-400 uppercase tracking-wider px-2">Filtros:</span>
        <Select value={filterTipo} onValueChange={setFilterTipo} disabled={constantsLoading}>
          <SelectTrigger className="w-48 bg-white rounded-xl border-slate-200">
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
          <SelectTrigger className="w-56 bg-white rounded-xl border-slate-200">
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
            className="text-xs text-slate-500 hover:text-slate-900"
          >
            Limpar Filtros
          </Button>
        )}
      </div>

      {/* Data Table */}
      {isLoading ? (
        <div className="glass-card rounded-2xl p-12 text-center text-slate-400 flex flex-col items-center gap-3">
          <div className="h-8 w-8 animate-spin rounded-full border-4 border-indigo-200 border-t-indigo-600" />
          <p className="text-sm font-medium">Carregando estoque...</p>
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


