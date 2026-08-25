import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { getLoansChart, getRegistrationsChart } from '@/api/reports'
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
} from 'recharts'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Calendar, Filter, ArrowUpRight, ArrowDownLeft, PackagePlus } from 'lucide-react'

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

export default function ChartsPage() {
  const now = new Date()
  const [year, setYear] = useState(String(now.getFullYear()))
  const [month, setMonth] = useState(String(now.getMonth() + 1))
  const [params, setParams] = useState({ year: now.getFullYear(), month: now.getMonth() + 1 })

  const { data: loansData, isLoading: loansLoading } = useQuery({
    queryKey: ['chart-loans', params.year, params.month],
    queryFn: () => getLoansChart(params.year, params.month),
  })

  const { data: registrationsData, isLoading: regLoading } = useQuery({
    queryKey: ['chart-registrations', params.year, params.month],
    queryFn: () => getRegistrationsChart(params.year, params.month),
  })

  function buildLoansChartData() {
    if (!loansData) return []
    return loansData.days.map((day, i) => ({
      dia: `Dia ${day}`,
      Empréstimos: loansData.values[i],
      Devoluções: loansData.values2?.[i] ?? 0,
    }))
  }

  function buildRegChartData() {
    if (!registrationsData) return []
    return registrationsData.days.map((day, i) => ({
      dia: `Dia ${day}`,
      Cadastros: registrationsData.values[i],
    }))
  }

  // Totais do mês selecionado
  const totalEmprestimos = loansData?.values?.reduce((acc, v) => acc + v, 0) ?? 0
  const totalDevolucoes = loansData?.values2?.reduce((acc, v) => acc + v, 0) ?? 0
  const totalCadastros = registrationsData?.values?.reduce((acc, v) => acc + v, 0) ?? 0

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold tracking-tight text-foreground font-heading">
            Dashboard & Indicadores
          </h2>
          <p className="text-sm text-muted-foreground font-medium">
            Análise temporal de empréstimos, devoluções e novos cadastros de equipamentos
          </p>
        </div>
      </div>

      {/* Filter Control Bar */}
      <div className="flex flex-wrap items-end gap-4 p-4 rounded-2xl bg-card backdrop-blur-md border border-border shadow-sm">
        <div className="flex items-center gap-2 text-primary font-bold text-xs uppercase tracking-wider pr-2">
          <Filter size={16} />
          <span>Filtrar Período</span>
        </div>
        <div className="flex flex-col gap-1.5">
          <Label className="text-xs font-semibold text-muted-foreground">Ano</Label>
          <Input
            value={year}
            onChange={(e) => setYear(e.target.value)}
            className="w-28 rounded-xl bg-card border-border"
          />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label className="text-xs font-semibold text-muted-foreground">Mês</Label>
          <Select value={month} onValueChange={setMonth}>
            <SelectTrigger className="w-44 rounded-xl bg-card border-border">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              {MONTHS.map((m, i) => (
                <SelectItem key={i + 1} value={String(i + 1)}>
                  {m}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
        <Button
          variant="default"
          onClick={() => setParams({ year: Number(year), month: Number(month) })}
          className="rounded-xl"
        >
          <Calendar size={15} className="mr-1.5" />
          Aplicar Filtro
        </Button>
      </div>

      {/* Summary KPI Cards for Selected Month */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="surface rounded-lg p-5 flex items-center gap-4">
          <div className="h-12 w-12 rounded-xl bg-primary/10 text-primary flex items-center justify-center font-bold">
            <ArrowUpRight size={24} />
          </div>
          <div>
            <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider">Empréstimos no Mês</p>
            <p className="text-2xl font-bold text-primary font-heading">{totalEmprestimos}</p>
          </div>
        </div>

        <div className="surface rounded-lg p-5 flex items-center gap-4">
          <div className="h-12 w-12 rounded-xl bg-success/10 text-success flex items-center justify-center font-bold">
            <ArrowDownLeft size={24} />
          </div>
          <div>
            <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider">Devoluções no Mês</p>
            <p className="text-2xl font-bold text-success font-heading">{totalDevolucoes}</p>
          </div>
        </div>

        <div className="surface rounded-lg p-5 flex items-center gap-4">
          <div className="h-12 w-12 rounded-xl bg-purple/10 text-purple-light flex items-center justify-center font-bold">
            <PackagePlus size={24} />
          </div>
          <div>
            <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider">Cadastros no Mês</p>
            <p className="text-2xl font-bold text-purple-light font-heading">{totalCadastros}</p>
          </div>
        </div>
      </div>

      {/* Gráfico 1: Empréstimos x Devoluções */}
      <div className="surface rounded-lg p-6 space-y-4">
        <div className="flex items-center justify-between">
          <div>
            <h3 className="text-base font-bold text-foreground font-heading">
              Movimentação de Empréstimos × Devoluções por Dia
            </h3>
            <p className="text-xs text-muted-foreground">
              Comparativo de saídas e retornos de equipamentos em {MONTHS[params.month - 1]} de {params.year}
            </p>
          </div>
        </div>

        {loansLoading ? (
          <div className="h-72 flex items-center justify-center text-muted-foreground text-sm">Carregando gráfico...</div>
        ) : buildLoansChartData().every((d) => !d.Empréstimos && !d.Devoluções) ? (
          <div className="h-72 flex flex-col items-center justify-center gap-2 text-muted-foreground text-sm">
            <PackagePlus size={22} className="opacity-40" />
            Nenhuma movimentação registrada em {MONTHS[params.month - 1]} de {params.year}.
          </div>
        ) : (
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={buildLoansChartData()} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" vertical={false} />
              <XAxis dataKey="dia" tick={{ fontSize: 11, fill: 'var(--muted-foreground)' }} axisLine={false} tickLine={false} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: 'var(--muted-foreground)' }} axisLine={false} tickLine={false} />
              <Tooltip
                contentStyle={{
                  backgroundColor: 'var(--popover)',
                  borderColor: 'var(--border)',
                  borderRadius: '10px',
                  boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.4)',
                  color: 'var(--foreground)',
                  fontSize: '12px',
                  fontWeight: '600',
                }}
                labelStyle={{ color: 'var(--foreground)' }}
              />
              <Legend wrapperStyle={{ paddingTop: '10px', fontSize: '12px', fontWeight: '600', color: 'var(--muted-foreground)' }} />
              <Bar dataKey="Empréstimos" fill="#38bdf8" radius={[6, 6, 0, 0]} maxBarSize={28} />
              <Bar dataKey="Devoluções" fill="#34d399" radius={[6, 6, 0, 0]} maxBarSize={28} />
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>

      {/* Gráfico 2: Novos Cadastros */}
      <div className="surface rounded-lg p-6 space-y-4">
        <div>
          <h3 className="text-base font-bold text-foreground font-heading">
            Novos Cadastros de Equipamentos por Dia
          </h3>
          <p className="text-xs text-muted-foreground">
            Volume de inclusões de patrimônios no acervo em {MONTHS[params.month - 1]} de {params.year}
          </p>
        </div>

        {regLoading ? (
          <div className="h-72 flex items-center justify-center text-muted-foreground text-sm">Carregando gráfico...</div>
        ) : buildRegChartData().every((d) => !d.Cadastros) ? (
          <div className="h-72 flex flex-col items-center justify-center gap-2 text-muted-foreground text-sm">
            <PackagePlus size={22} className="opacity-40" />
            Nenhum cadastro registrado em {MONTHS[params.month - 1]} de {params.year}.
          </div>
        ) : (
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={buildRegChartData()} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" vertical={false} />
              <XAxis dataKey="dia" tick={{ fontSize: 11, fill: 'var(--muted-foreground)' }} axisLine={false} tickLine={false} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: 'var(--muted-foreground)' }} axisLine={false} tickLine={false} />
              <Tooltip
                contentStyle={{
                  backgroundColor: 'var(--popover)',
                  borderColor: 'var(--border)',
                  borderRadius: '10px',
                  boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.4)',
                  color: 'var(--foreground)',
                  fontSize: '12px',
                  fontWeight: '600',
                }}
                labelStyle={{ color: 'var(--foreground)' }}
              />
              <Bar dataKey="Cadastros" fill="#A855F7" radius={[6, 6, 0, 0]} maxBarSize={32} />
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>
    </div>
  )
}

