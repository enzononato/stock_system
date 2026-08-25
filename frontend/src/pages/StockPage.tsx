import { useMemo, useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import type { ColumnDef } from '@tanstack/react-table'
import { PieChart, Pie, Cell, ResponsiveContainer, AreaChart, Area, XAxis, Tooltip } from 'recharts'
import { listItemsPaginated, type Item } from '@/api/items'
import { listUnidades } from '@/api/unidades'
import { getRegistrationsChart } from '@/api/reports'
import { DataTable } from '@/components/ui/DataTable'
import { StatusBadge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { KpiCard } from '@/components/ui/KpiCard'
import { ProgressRow } from '@/components/ui/ProgressRow'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { ItemDetailsModal } from '@/components/equipment/ItemDetailsModal'
import SpecularButton from '@/components/effects/SpecularButton'
import { useAuth } from '@/contexts/AuthContext'
import { useConstants } from '@/hooks/useConstants'
import { formatDate, formatDateTime } from '@/lib/utils'
import { Plus, Pencil, RefreshCw, Eye, ArrowRight } from 'lucide-react'

const STATUS_OPTIONS = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução']
const FETCH_ALL_LIMIT = 500
const DONUT_COLORS = { Disponível: '#34d399', Indisponível: '#38bdf8', Pendente: '#fbbf24' }

export default function StockPage() {
  const { hasRole } = useAuth()
  const navigate = useNavigate()
  const { equipmentTypes, isLoading: constantsLoading } = useConstants()
  const [filterTipo, setFilterTipo] = useState('all')
  const [filterStatus, setFilterStatus] = useState('all')
  const [filterRevenda, setFilterRevenda] = useState('all')
  const [isRefreshing, setIsRefreshing] = useState(false)
  const [selectedItem, setSelectedItem] = useState<Item | null>(null)

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades-filter'],
    queryFn: () => listUnidades(),
  })

  const { data, isLoading, refetch, dataUpdatedAt } = useQuery({
    queryKey: ['items', filterTipo, filterStatus, filterRevenda],
    queryFn: () =>
      listItemsPaginated({
        tipo: filterTipo !== 'all' ? filterTipo : undefined,
        status: filterStatus !== 'all' ? filterStatus : undefined,
        revenda: filterRevenda !== 'all' ? filterRevenda : undefined,
        limit: FETCH_ALL_LIMIT,
      }),
  })

  const now = new Date()
  const { data: registrations } = useQuery({
    queryKey: ['chart-registrations', now.getFullYear(), now.getMonth() + 1],
    queryFn: () => getRegistrationsChart(now.getFullYear(), now.getMonth() + 1),
  })

  const items = data?.items ?? []
  const total = data?.total ?? 0

  const disponiveisCount = items.filter((i) => i.status === 'Disponível').length
  const indisponiveisCount = items.filter((i) => i.status === 'Indisponível').length
  const pendentesCount = items.filter((i) => i.status?.startsWith('Pendente')).length

  // Trend series driven by real registration data (last ~14 points) for the
  // KPI sparklines — cumulative growth rather than a fabricated shape.
  const trend = useMemo(() => {
    if (!registrations?.values?.length) return []
    let acc = 0
    return registrations.values.slice(-14).map((v) => {
      acc += v
      return acc
    })
  }, [registrations])

  const areaData = useMemo(() => {
    if (!registrations?.days?.length) return []
    return registrations.days.map((day, i) => ({ dia: `${day}`, cadastros: registrations.values[i] }))
  }, [registrations])

  const donutData = [
    { name: 'Disponível', value: disponiveisCount },
    { name: 'Indisponível', value: indisponiveisCount },
    { name: 'Pendente', value: pendentesCount },
  ].filter((d) => d.value > 0)

  // Grouped by unit ("revenda") — mirrors a warehouse/sync overview table:
  // status dot, % available as the progress metric, last movement as the
  // trailing timestamp.
  const unidadeRows = useMemo(() => {
    const map = new Map<string, { total: number; disponivel: number; lastDate?: string }>()
    for (const it of items) {
      const key = it.revenda || 'Sem unidade'
      const entry = map.get(key) ?? { total: 0, disponivel: 0 }
      entry.total += 1
      if (it.status === 'Disponível') entry.disponivel += 1
      if (it.date_registered && (!entry.lastDate || it.date_registered > entry.lastDate)) {
        entry.lastDate = it.date_registered
      }
      map.set(key, entry)
    }
    return [...map.entries()]
      .map(([revenda, v]) => ({
        revenda,
        total: v.total,
        pct: v.total ? Math.round((v.disponivel / v.total) * 100) : 0,
        lastDate: v.lastDate,
      }))
      .sort((a, b) => b.total - a.total)
      .slice(0, 6)
  }, [items])

  const recentItems = useMemo(
    () =>
      [...items]
        .filter((i) => i.date_registered)
        .sort((a, b) => new Date(b.date_registered!).getTime() - new Date(a.date_registered!).getTime())
        .slice(0, 6),
    [items]
  )

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
          className="font-mono font-semibold text-primary hover:underline text-xs"
          title="Clique para ver detalhes do item"
        >
          #{row.original.id}
        </button>
      ),
    },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'model', header: 'Modelo' },
    { accessorKey: 'status', header: 'Status', cell: ({ row }) => <StatusBadge status={row.original.status} /> },
    { accessorKey: 'peripheral_count', header: 'Periféricos', size: 90 },
    { accessorKey: 'assigned_to', header: 'Usuário Alocado', cell: ({ getValue }) => (getValue() as string) || '-' },
    { accessorKey: 'revenda', header: 'Unidade' },
    { accessorKey: 'identificador', header: 'Identificador', cell: ({ getValue }) => (getValue() as string) || '-' },
    { accessorKey: 'setor', header: 'Setor', cell: ({ getValue }) => (getValue() as string) || '-' },
    { accessorKey: 'ip', header: 'IP', cell: ({ getValue }) => (getValue() as string) || '-' },
    { accessorKey: 'date_registered', header: 'Data Cadastro', cell: ({ getValue }) => formatDate(getValue() as string) },
    {
      id: 'actions',
      header: '',
      cell: ({ row }: { row: { original: Item } }) => (
        <div className="flex items-center gap-1">
          <Button
            variant="ghost"
            size="sm"
            onClick={() => setSelectedItem(row.original)}
            className="h-8 w-8 p-0 text-muted-foreground hover:text-primary hover:bg-accent rounded-md"
            title="Ver detalhes do equipamento"
          >
            <Eye size={14} />
          </Button>
          {hasRole('Gestor', 'Técnico') && (
            <Button
              variant="ghost"
              size="sm"
              onClick={() => navigate(`/edit/${row.original.id}`)}
              className="h-8 w-8 p-0 text-muted-foreground hover:text-primary hover:bg-accent rounded-md"
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
    <div className="space-y-5">
      {/* Header Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-h1 text-foreground">Estoque de Equipamentos</h2>
          <p className="text-caption mt-1">Sincronização e visão geral do inventário de TI em tempo real</p>
        </div>
        <div className="flex items-center gap-2">
          <Button variant="outline" size="sm" onClick={handleRefresh}>
            <RefreshCw size={14} className={`mr-2 ${isRefreshing ? 'animate-spin text-primary' : ''}`} />
            Atualizar
          </Button>
          {hasRole('Gestor', 'Técnico') && (
            <SpecularButton size="md" radius={8} lineColor="#38bdf8" baseColor="#151b2b" onClick={() => navigate('/register')}>
              <Plus size={16} />
              Novo Equipamento
            </SpecularButton>
          )}
        </div>
      </div>

      {/* KPI row with sparklines */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <KpiCard
          label="Total em Estoque"
          value={total}
          hint={`Atualizado ${formatDateTime(new Date(dataUpdatedAt).toISOString())}`}
          tone="primary"
          sparkline={trend}
          highlighted
        />
        <KpiCard label="Disponíveis" value={disponiveisCount} hint={`${total ? Math.round((disponiveisCount / total) * 100) : 0}% do total`} tone="success" sparkline={trend.map((v) => v * 0.6)} />
        <KpiCard label="Em Empréstimo" value={indisponiveisCount} hint="Alocados a usuários" tone="info" sparkline={trend.map((v) => v * 0.3)} />
        <KpiCard label="Ações Pendentes" value={pendentesCount} hint="Aguardando confirmação" tone="warning" sparkline={trend.map((v) => v * 0.15)} />
      </div>

      {/* Unit overview — warehouse-style status table */}
      <div className="surface rounded-lg overflow-hidden">
        <div className="px-4 py-3 border-b border-border flex items-center justify-between">
          <h3 className="text-xs font-semibold text-foreground">Visão por Unidade de Revenda</h3>
          <button onClick={() => navigate('/unidades')} className="text-[11px] text-primary hover:underline flex items-center gap-1">
            Ver todas <ArrowRight size={11} />
          </button>
        </div>
        <div className="hidden sm:grid grid-cols-[1.4fr_0.6fr_1fr_auto] gap-3 px-4 py-2 text-[10px] font-semibold uppercase tracking-wider text-muted-foreground border-b border-border/60">
          <span>Unidade</span>
          <span>Disponibilidade</span>
          <span></span>
          <span className="text-right">Última movimentação</span>
        </div>
        {unidadeRows.length === 0 ? (
          <p className="px-4 py-6 text-center text-xs text-muted-foreground">Nenhum dado de unidade ainda.</p>
        ) : (
          unidadeRows.map((u) => (
            <ProgressRow
              key={u.revenda}
              label={u.revenda}
              sublabel={`${u.total} ${u.total === 1 ? 'item' : 'itens'}`}
              percent={u.pct}
              tone={u.pct > 60 ? 'success' : u.pct > 30 ? 'warning' : 'danger'}
              trailing={<span className="text-[11px] text-muted-foreground tabular">{formatDate(u.lastDate)}</span>}
            />
          ))
        )}
      </div>

      {/* Area chart + donut, side by side */}
      <div className="grid grid-cols-1 lg:grid-cols-[1fr_320px] gap-3">
        <div className="surface rounded-lg p-4">
          <h3 className="text-xs font-semibold text-foreground mb-3">Cadastros no Período</h3>
          <div className="h-[180px]">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={areaData} margin={{ top: 4, right: 8, left: -20, bottom: 0 }}>
                <defs>
                  <linearGradient id="stockAreaGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#38bdf8" stopOpacity={0.35} />
                    <stop offset="100%" stopColor="#38bdf8" stopOpacity={0} />
                  </linearGradient>
                </defs>
                <XAxis dataKey="dia" tick={{ fontSize: 10, fill: 'var(--muted-foreground)' }} axisLine={false} tickLine={false} />
                <Tooltip
                  contentStyle={{ background: 'var(--popover)', border: '1px solid var(--border)', borderRadius: 8, fontSize: 12 }}
                  labelStyle={{ color: 'var(--foreground)' }}
                />
                <Area type="monotone" dataKey="cadastros" stroke="#38bdf8" strokeWidth={2} fill="url(#stockAreaGrad)" />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="surface rounded-lg p-4 flex flex-col">
          <h3 className="text-xs font-semibold text-foreground mb-2">Status do Estoque</h3>
          <div className="flex-1 flex items-center gap-4">
            <div className="h-[140px] w-[140px] relative shrink-0">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={donutData} dataKey="value" innerRadius={44} outerRadius={62} paddingAngle={2} stroke="none">
                    {donutData.map((d) => (
                      <Cell key={d.name} fill={DONUT_COLORS[d.name as keyof typeof DONUT_COLORS]} />
                    ))}
                  </Pie>
                </PieChart>
              </ResponsiveContainer>
              <div className="absolute inset-0 flex flex-col items-center justify-center">
                <span className="text-lg font-semibold text-foreground tabular">{total}</span>
                <span className="text-[9px] text-muted-foreground">Total</span>
              </div>
            </div>
            <div className="space-y-2 min-w-0">
              {donutData.map((d) => (
                <div key={d.name} className="flex items-center gap-2 text-[11px]">
                  <span className="h-2 w-2 rounded-full shrink-0" style={{ background: DONUT_COLORS[d.name as keyof typeof DONUT_COLORS] }} />
                  <span className="text-muted-foreground truncate">{d.name}</span>
                  <span className="ml-auto font-medium text-foreground tabular">{d.value}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* Filter Bar */}
      <div className="flex flex-wrap items-center gap-3 p-3 rounded-lg surface">
        <span className="text-[11px] font-semibold text-muted-foreground uppercase tracking-wider px-1">Filtros</span>
        <Select value={filterTipo} onValueChange={setFilterTipo} disabled={constantsLoading}>
          <SelectTrigger className="w-48">
            <SelectValue placeholder="Tipo de Equipamento" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="all">Todos os tipos</SelectItem>
            {equipmentTypes.map((t) => (
              <SelectItem key={t} value={t}>{t}</SelectItem>
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
              <SelectItem key={s} value={s}>{s}</SelectItem>
            ))}
          </SelectContent>
        </Select>
        <Select value={filterRevenda} onValueChange={setFilterRevenda}>
          <SelectTrigger className="w-56">
            <SelectValue placeholder="Unidade de Revenda" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="all">Todas as unidades</SelectItem>
            {unidades.map((u) => (
              <SelectItem key={u.id} value={u.nome}>{u.nome}</SelectItem>
            ))}
          </SelectContent>
        </Select>
        {(filterTipo !== 'all' || filterStatus !== 'all' || filterRevenda !== 'all') && (
          <Button variant="ghost" size="sm" onClick={() => { setFilterTipo('all'); setFilterStatus('all'); setFilterRevenda('all') }} className="text-xs text-muted-foreground hover:text-foreground">
            Limpar Filtros
          </Button>
        )}
      </div>

      {/* Data Table */}
      {isLoading ? (
        <div className="surface rounded-lg p-12 text-center text-muted-foreground flex flex-col items-center gap-3">
          <div className="h-8 w-8 animate-spin rounded-full border-4 border-muted border-t-primary" />
          <p className="text-sm font-medium">Carregando estoque...</p>
        </div>
      ) : (
        <DataTable data={items} columns={columns} onRowClick={(item) => setSelectedItem(item)} searchPlaceholder="Buscar por patrimônio, marca, modelo, usuário, setor, IP..." />
      )}

      {/* Recent activity log, in the style of a "sync log" table */}
      <div className="surface rounded-lg overflow-hidden">
        <div className="px-4 py-3 border-b border-border flex items-center justify-between">
          <h3 className="text-xs font-semibold text-foreground">Log de Cadastros Recentes</h3>
        </div>
        {recentItems.length === 0 ? (
          <p className="px-4 py-6 text-center text-xs text-muted-foreground">Nenhum item recente.</p>
        ) : (
          <div className="divide-y divide-border/60">
            {recentItems.map((it) => (
              <button
                key={it.id}
                onClick={() => setSelectedItem(it)}
                className="w-full grid grid-cols-[auto_1fr_auto_auto] items-center gap-3 px-4 py-2.5 text-left hover:bg-accent/50 transition-colors"
              >
                <span className="text-[11px] text-muted-foreground tabular w-16">{formatDate(it.date_registered)}</span>
                <span className="text-xs text-foreground truncate">
                  <span className="text-muted-foreground">#{it.id}</span> {it.brand} {it.model} cadastrado
                </span>
                <StatusBadge status={it.status} />
                <span className="text-[11px] text-muted-foreground">{it.revenda || '-'}</span>
              </button>
            ))}
          </div>
        )}
      </div>

      <ItemDetailsModal item={selectedItem} onClose={() => setSelectedItem(null)} />
    </div>
  )
}
