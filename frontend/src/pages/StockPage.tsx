import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import type { ColumnDef } from '@tanstack/react-table'
import { listItemsPaginated, type Item } from '@/api/items'
import { DataTable } from '@/components/ui/DataTable'
import { StatusBadge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { useAuth } from '@/contexts/AuthContext'
import { useConstants } from '@/hooks/useConstants'
import { formatDate } from '@/lib/utils'
import { Plus, Pencil, RefreshCw } from 'lucide-react'

// Status não vem de /api/constants (é um enum interno do fluxo de estoque,
// não uma lista de domínio cadastrável) — continua fixo aqui.
const STATUS_OPTIONS = ['Disponível','Indisponível','Pendente','Pendente Devolução']

// Sem UI de paginação por enquanto (DataTable ainda não suporta — em
// desenvolvimento em paralelo). Usamos o teto de página do backend para
// preservar o comportamento atual (mostrar todo o estoque filtrado, sem
// truncar em 50 itens) enquanto consumimos o novo contrato {items, total}.
const FETCH_ALL_LIMIT = 500

export default function StockPage() {
  const { hasRole } = useAuth()
  const navigate = useNavigate()
  const { equipmentTypes, isLoading: constantsLoading } = useConstants()
  const [filterTipo, setFilterTipo] = useState('all')
  const [filterStatus, setFilterStatus] = useState('all')

  const { data, isLoading, refetch } = useQuery({
    queryKey: ['items', filterTipo, filterStatus],
    queryFn: () => listItemsPaginated({
      tipo: filterTipo !== 'all' ? filterTipo : undefined,
      status: filterStatus !== 'all' ? filterStatus : undefined,
      limit: FETCH_ALL_LIMIT,
    }),
  })
  const items = data?.items ?? []
  const total = data?.total ?? 0

  const columns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'model', header: 'Modelo' },
    {
      accessorKey: 'status',
      header: 'Status',
      cell: ({ row }) => <StatusBadge status={row.original.status} />,
    },
    { accessorKey: 'peripheral_count', header: 'Periféricos', size: 90 },
    { accessorKey: 'assigned_to', header: 'Usuário', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'identificador', header: 'Identificador', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'setor', header: 'Setor', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'ip', header: 'IP', cell: ({ getValue }) => getValue() as string || '-' },
    {
      accessorKey: 'date_registered',
      header: 'Cadastro',
      cell: ({ getValue }) => formatDate(getValue() as string),
    },
    ...(hasRole('Gestor', 'Técnico') ? [{
      id: 'actions',
      header: '',
      cell: ({ row }: { row: { original: Item } }) => (
        <Button
          variant="ghost"
          size="sm"
          onClick={() => navigate(`/edit/${row.original.id}`)}
        >
          <Pencil size={14} />
        </Button>
      ),
    } as ColumnDef<Item, unknown>] : []),
  ]

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-xl font-semibold text-slate-900">Estoque</h2>
          <p className="text-sm text-slate-500">{total} itens ativos</p>
        </div>
        <div className="flex items-center gap-2">
          <Button variant="outline" size="sm" onClick={() => refetch()}>
            <RefreshCw size={14} className="mr-1" />
            Atualizar
          </Button>
          {hasRole('Gestor', 'Técnico') && (
            <Button size="sm" onClick={() => navigate('/register')}>
              <Plus size={14} className="mr-1" />
              Novo Item
            </Button>
          )}
        </div>
      </div>

      {/* Filtros */}
      <div className="flex gap-3 flex-wrap">
        <Select value={filterTipo} onValueChange={setFilterTipo} disabled={constantsLoading}>
          <SelectTrigger className="w-44">
            <SelectValue placeholder="Tipo" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="all">Todos os tipos</SelectItem>
            {equipmentTypes.map((t) => <SelectItem key={t} value={t}>{t}</SelectItem>)}
          </SelectContent>
        </Select>
        <Select value={filterStatus} onValueChange={setFilterStatus}>
          <SelectTrigger className="w-52">
            <SelectValue placeholder="Status" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="all">Todos os status</SelectItem>
            {STATUS_OPTIONS.map((s) => <SelectItem key={s} value={s}>{s}</SelectItem>)}
          </SelectContent>
        </Select>
      </div>

      {isLoading ? (
        <div className="py-12 text-center text-slate-400">Carregando...</div>
      ) : (
        <DataTable data={items} columns={columns} searchPlaceholder="Buscar por marca, modelo, usuário..." />
      )}
    </div>
  )
}
