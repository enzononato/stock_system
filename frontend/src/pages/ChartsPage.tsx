import { useMemo, useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import {
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
} from 'recharts'
import { Calendar } from 'lucide-react'

import { getLoansChart, getRegistrationsChart, getMonthlyReport } from '@/api/reports'
import { listUnidades } from '@/api/unidades'
import { useConstants } from '@/hooks/useConstants'
import { LoadingState, EmptyState, ErrorState } from '@/components/ui/StateBlocks'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import { PageHeader } from '@/components/layout/PageHeader'
import { useTheme } from '@/lib/theme'

const MONTHS = [
  'Janeiro',
  'Fevereiro',
  'Março',
  'Abril',
  'Maio',
  'Junho',
  'Julho',
  'Agosto',
  'Setembro',
  'Outubro',
  'Novembro',
  'Dezembro',
]

const TOOLTIP_STYLE = {
  backgroundColor: 'var(--surface)',
  borderColor: 'var(--border)',
  borderRadius: 4,
  fontSize: 12,
  fontWeight: 600,
  color: 'var(--foreground)',
  boxShadow: 'var(--shadow-overlay)',
}

export default function ChartsPage() {
  useTheme()

  const now = new Date()
  const [year, setYear] = useState(String(now.getFullYear()))
  const [month, setMonth] = useState(String(now.getMonth() + 1))
  const [filterRevenda, setFilterRevenda] = useState('all')
  const [params, setParams] = useState({
    year: now.getFullYear(),
    month: now.getMonth() + 1,
    revenda: 'all',
  })

  const { revendas = [] } = useConstants()
  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades-charts'],
    queryFn: () => listUnidades(),
  })

  const revendaOptions = useMemo(() => {
    const set = new Set<string>()
    revendas.forEach((r) => r && set.add(r))
    unidades.forEach((u) => u.nome && set.add(u.nome))
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))
  }, [revendas, unidades])

  const loans = useQuery({
    queryKey: ['chart-loans', params.year, params.month],
    queryFn: () => getLoansChart(params.year, params.month),
  })
  const registrations = useQuery({
    queryKey: ['chart-registrations', params.year, params.month],
    queryFn: () => getRegistrationsChart(params.year, params.month),
  })
  const monthly = useQuery({
    queryKey: ['chart-monthly-report', params.year, params.month],
    queryFn: () => getMonthlyReport(params.year, params.month),
    enabled: params.revenda !== 'all',
  })

  const isFiltered = params.revenda !== 'all'
  const daysInMonth = new Date(params.year, params.month, 0).getDate()
  const daysArray = Array.from({ length: daysInMonth }, (_, i) => i + 1)

  const { loansChartData, regChartData, totalEmprestimos, totalDevolucoes, totalCadastros } =
    useMemo(() => {
      if (!isFiltered) {
        const lData = (loans.data?.days ?? []).map((day, i) => ({
          dia: `Dia ${day}`,
          Empréstimos: loans.data?.values[i] ?? 0,
          Devoluções: loans.data?.values2?.[i] ?? 0,
        }))
        const rData = (registrations.data?.days ?? []).map((day, i) => ({
          dia: `Dia ${day}`,
          Cadastros: registrations.data?.values[i] ?? 0,
        }))
        return {
          loansChartData: lData,
          regChartData: rData,
          totalEmprestimos: loans.data?.values?.reduce((a, v) => a + v, 0) ?? 0,
          totalDevolucoes: loans.data?.values2?.reduce((a, v) => a + v, 0) ?? 0,
          totalCadastros: registrations.data?.values?.reduce((a, v) => a + v, 0) ?? 0,
        }
      }

      const rows = (monthly.data ?? []).filter((r) => r.revenda === params.revenda)
      const empByDay: Record<number, number> = {}
      const devByDay: Record<number, number> = {}
      const cadByDay: Record<number, number> = {}

      rows.forEach((r) => {
        if (r.data_emprestimo) {
          const d = new Date(r.data_emprestimo).getDate()
          if (d) empByDay[d] = (empByDay[d] ?? 0) + 1
        }
        if (r.data_devolucao) {
          const d = new Date(r.data_devolucao).getDate()
          if (d) devByDay[d] = (devByDay[d] ?? 0) + 1
        }
        if (r.operation_type === 'Cadastro') {
          const dateStr = r.data_confirmacao ?? r.data_emprestimo
          const d = dateStr ? new Date(dateStr).getDate() : null
          if (d) cadByDay[d] = (cadByDay[d] ?? 0) + 1
        }
      })

      const lData = daysArray.map((d) => ({
        dia: `Dia ${d}`,
        Empréstimos: empByDay[d] ?? 0,
        Devoluções: devByDay[d] ?? 0,
      }))
      const rData = daysArray.map((d) => ({
        dia: `Dia ${d}`,
        Cadastros: cadByDay[d] ?? 0,
      }))

      return {
        loansChartData: lData,
        regChartData: rData,
        totalEmprestimos: Object.values(empByDay).reduce((a, v) => a + v, 0),
        totalDevolucoes: Object.values(devByDay).reduce((a, v) => a + v, 0),
        totalCadastros: Object.values(cadByDay).reduce((a, v) => a + v, 0),
      }
    }, [isFiltered, loans.data, registrations.data, monthly.data, params.revenda, daysArray])

  const hasInvalidYear = !/^\d{4}$/.test(year) || Number(year) < 2000 || Number(year) > 2100
  const monthLabel = `${MONTHS[params.month - 1] ?? 'Mês'} / ${params.year}${
    params.revenda !== 'all' ? ` · ${params.revenda}` : ''
  }`
  const isLoading = isFiltered ? monthly.isLoading : loans.isLoading || registrations.isLoading
  const hasError = isFiltered ? monthly.error : loans.error || registrations.error

  function applyPeriod() {
    if (hasInvalidYear) return
    setParams({ year: Number(year), month: Number(month), revenda: filterRevenda })
  }

  // Distribuição por dia da semana (a partir de loans data)
  const weekdayDist = useMemo(() => {
    const days = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb']
    const counts: number[] = Array(7).fill(0)
    loansChartData.forEach((d, idx) => {
      const date = new Date(params.year, params.month - 1, idx + 1)
      counts[date.getDay()] = (counts[date.getDay()] ?? 0) + (d.Empréstimos ?? 0)
    })
    return days.map((d, i) => ({ dia: d, Empréstimos: counts[i] ?? 0 }))
  }, [loansChartData, params.year, params.month])

  return (
    <div className="page-container-dense space-y-6">
      <PageHeader
        eyebrow="Inteligência de Operações · Análise Comparativa"
        title="Indicadores Operacionais"
        description="Movimentação patrimonial comparativa por período, filial e tipo de operação."
      />

      {/* Toolbar de Filtro */}
      <div className="rounded-md border border-border bg-surface p-4">
        <div className="flex flex-wrap items-end gap-4">
          <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground self-end pb-1">
            Período
          </span>

          <div className="space-y-1">
            <Label className="text-caption font-medium">Ano</Label>
            <Input
              inputMode="numeric"
              maxLength={4}
              value={year}
              onChange={(e) => setYear(e.target.value.replace(/\D/g, '').slice(0, 4))}
              className="w-24 h-9 rounded-sm border-border bg-background text-body-sm font-mono"
              aria-invalid={hasInvalidYear}
            />
          </div>

          <div className="space-y-1">
            <Label className="text-caption font-medium">Mês</Label>
            <Select value={month} onValueChange={setMonth}>
              <SelectTrigger className="w-40 h-9 rounded-sm border-border bg-background text-body-sm">
                <SelectValue />
              </SelectTrigger>
              <SelectContent className="rounded-sm border-border bg-surface">
                {MONTHS.map((m, i) => (
                  <SelectItem key={m} value={String(i + 1)} className="text-body-sm">
                    {m}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          <div className="space-y-1">
            <Label className="text-caption font-medium">Unidade / Filial</Label>
            <Select value={filterRevenda} onValueChange={setFilterRevenda}>
              <SelectTrigger className="w-52 h-9 rounded-sm border-border bg-background text-body-sm">
                <SelectValue placeholder="Todas" />
              </SelectTrigger>
              <SelectContent className="rounded-sm border-border bg-surface">
                <SelectItem value="all" className="text-body-sm">
                  Todas as unidades
                </SelectItem>
                {revendaOptions.map((r) => (
                  <SelectItem key={r} value={r} className="text-body-sm">
                    {r}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          <Button
            size="sm"
            onClick={applyPeriod}
            disabled={hasInvalidYear}
            className="h-9 rounded-sm bg-primary text-primary-foreground hover:opacity-90 font-medium px-4"
          >
            <Calendar className="mr-1.5 size-4" />
            Aplicar Filtro
          </Button>

          {hasInvalidYear && (
            <p className="text-caption text-destructive self-end pb-1">
              Ano inválido (2000–2100)
            </p>
          )}
        </div>
      </div>

      {/* KPIs Numéricos */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        {[
          { label: 'Empréstimos no mês', value: totalEmprestimos, symbol: '↑' },
          { label: 'Devoluções no mês', value: totalDevolucoes, symbol: '↓' },
          { label: 'Cadastros no mês', value: totalCadastros, symbol: '+' },
        ].map((stat) => (
          <div key={stat.label} className="surface-panel p-5">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
              {stat.label}
            </span>
            <div className="flex items-baseline gap-2 mt-2">
              <span className="font-mono text-heading-lg font-bold text-foreground">
                {stat.value}
              </span>
              <span className="font-mono text-body-lg text-muted-foreground">{stat.symbol}</span>
            </div>
            <span className="text-caption text-muted-foreground mt-1 block">{monthLabel}</span>
          </div>
        ))}
      </div>

      {/* Gráfico 1: Área de Empréstimos × Devoluções */}
      <div className="surface-panel p-6 space-y-4">
        <div className="border-b border-border pb-3">
          <h2 className="text-body-lg font-semibold text-foreground">
            Movimentação de Empréstimos × Devoluções por Dia
          </h2>
          <p className="text-caption text-muted-foreground mt-0.5">
            Comparativo de saídas e retornos de equipamentos em {monthLabel}
          </p>
        </div>

        {isLoading ? (
          <LoadingState label="Carregando indicadores..." />
        ) : hasError ? (
          <ErrorState
            error={hasError}
            onRetry={() => (isFiltered ? void monthly.refetch() : void loans.refetch())}
          />
        ) : loansChartData.every((d) => !d.Empréstimos && !d.Devoluções) ? (
          <EmptyState
            title="Nenhuma movimentação registrada neste período"
            description="Tente selecionar outro mês ou ano no filtro acima."
          />
        ) : (
          <div className="min-w-0 overflow-x-auto">
            <div className="min-w-[540px]">
              <ResponsiveContainer width="100%" height={320}>
                <AreaChart
                  data={loansChartData}
                  margin={{ top: 10, right: 10, left: -20, bottom: 0 }}
                >
                  <defs>
                    <linearGradient id="gradEmp" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="var(--primary)" stopOpacity={0.25} />
                      <stop offset="95%" stopColor="var(--primary)" stopOpacity={0.02} />
                    </linearGradient>
                    <linearGradient id="gradDev" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="var(--text-muted)" stopOpacity={0.2} />
                      <stop offset="95%" stopColor="var(--text-muted)" stopOpacity={0.01} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid
                    strokeDasharray="3 3"
                    stroke="var(--border)"
                    vertical={false}
                  />
                  <XAxis
                    dataKey="dia"
                    tick={{ fontSize: 11, fill: 'var(--muted-foreground)' }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <YAxis
                    allowDecimals={false}
                    tick={{ fontSize: 11, fill: 'var(--muted-foreground)' }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip contentStyle={TOOLTIP_STYLE} />
                  <Area
                    type="monotone"
                    dataKey="Empréstimos"
                    stroke="var(--foreground)"
                    strokeWidth={2}
                    fill="url(#gradEmp)"
                    dot={{ r: 2.5, fill: 'var(--foreground)' }}
                    activeDot={{ r: 4.5 }}
                  />
                  <Area
                    type="monotone"
                    dataKey="Devoluções"
                    stroke="var(--muted-foreground)"
                    strokeWidth={2}
                    fill="url(#gradDev)"
                    dot={{ r: 2.5, fill: 'var(--muted-foreground)' }}
                    activeDot={{ r: 4.5 }}
                  />
                </AreaChart>
              </ResponsiveContainer>

              {/* Legenda manual monocromática */}
              <div className="flex items-center gap-6 mt-4 pt-3 border-t border-border text-caption">
                <div className="flex items-center gap-2">
                  <div className="w-5 h-1 rounded bg-foreground" />
                  <span className="text-foreground font-medium">Empréstimos</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-5 h-1 rounded bg-muted-foreground" />
                  <span className="text-muted-foreground">Devoluções</span>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Grid: Gráfico Barras (Cadastros por dia) + Distribuição Semanal */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Gráfico 2: Área de Cadastros por Dia */}
        <div className="surface-panel p-5">
          <div className="border-b border-border pb-3 mb-4">
            <h2 className="text-body-lg font-semibold text-foreground">Cadastros por Dia</h2>
            <p className="text-caption text-muted-foreground mt-0.5">
              Novos patrimônios registrados no período.
            </p>
          </div>

          {isLoading ? (
            <LoadingState label="Carregando cadastros..." />
          ) : regChartData.every((d) => !d.Cadastros) ? (
            <EmptyState title="Nenhum cadastro neste período." />
          ) : (
            <div className="overflow-x-auto">
              <div className="min-w-[300px]">
                <ResponsiveContainer width="100%" height={230}>
                  <AreaChart data={regChartData} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
                    <defs>
                      <linearGradient id="gradCad" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor="var(--foreground)" stopOpacity={0.22} />
                        <stop offset="95%" stopColor="var(--foreground)" stopOpacity={0.01} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid
                      strokeDasharray="3 3"
                      stroke="var(--border)"
                      vertical={false}
                    />
                    <XAxis
                      dataKey="dia"
                      tick={{ fontSize: 10, fill: 'var(--muted-foreground)' }}
                      axisLine={false}
                      tickLine={false}
                    />
                    <YAxis
                      allowDecimals={false}
                      tick={{ fontSize: 10, fill: 'var(--muted-foreground)' }}
                      axisLine={false}
                      tickLine={false}
                    />
                    <Tooltip contentStyle={TOOLTIP_STYLE} />
                    <Area
                      type="monotone"
                      dataKey="Cadastros"
                      stroke="var(--foreground)"
                      strokeWidth={2}
                      fill="url(#gradCad)"
                      dot={{ r: 2.5, fill: 'var(--foreground)' }}
                      activeDot={{ r: 4.5 }}
                    />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}
        </div>

        {/* Gráfico 3: Área de Distribuição por dia da semana */}
        <div className="surface-panel p-5">
          <div className="border-b border-border pb-3 mb-4">
            <h2 className="text-body-lg font-semibold text-foreground">
              Distribuição Semanal de Empréstimos
            </h2>
            <p className="text-caption text-muted-foreground mt-0.5">
              Concentração de saídas por dia da semana.
            </p>
          </div>

          {weekdayDist.every((d) => !d.Empréstimos) ? (
            <EmptyState title="Nenhum dado disponível para este período." />
          ) : (
            <ResponsiveContainer width="100%" height={230}>
              <AreaChart data={weekdayDist} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
                <defs>
                  <linearGradient id="gradSem" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="var(--foreground)" stopOpacity={0.2} />
                    <stop offset="95%" stopColor="var(--foreground)" stopOpacity={0.01} />
                  </linearGradient>
                </defs>
                <CartesianGrid
                  strokeDasharray="3 3"
                  stroke="var(--border)"
                  vertical={false}
                />
                <XAxis
                  dataKey="dia"
                  tick={{ fontSize: 11, fill: 'var(--muted-foreground)' }}
                  axisLine={false}
                  tickLine={false}
                />
                <YAxis
                  allowDecimals={false}
                  tick={{ fontSize: 10, fill: 'var(--muted-foreground)' }}
                  axisLine={false}
                  tickLine={false}
                />
                <Tooltip contentStyle={TOOLTIP_STYLE} />
                <Area
                  type="monotone"
                  dataKey="Empréstimos"
                  stroke="var(--foreground)"
                  strokeWidth={2}
                  fill="url(#gradSem)"
                  dot={{ r: 2.5, fill: 'var(--foreground)' }}
                  activeDot={{ r: 4.5 }}
                />
              </AreaChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>
    </div>
  )
}
