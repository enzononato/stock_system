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
import { StatBlock } from '@/components/ui/StatBlock'
import { Calendar, Filter, Loader2 } from 'lucide-react'
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

/*
 * Paleta dos gráficos em escala de cinza, lida dos tokens CSS para acompanhar
 * o modo claro/escuro. recharts recebe cor por prop (fill/stroke), então não dá
 * para usar classe do Tailwind aqui — é preciso resolver o valor computado.
 */
function corDoToken(nome: string, alternativa: string): string {
  if (typeof window === 'undefined') return alternativa
  const valor = getComputedStyle(document.documentElement).getPropertyValue(nome).trim()
  return valor || alternativa
}

const CORES_GRAFICO = {
  serie1: () => corDoToken('--text-primary', '#111111'),
  serie2: () => corDoToken('--text-muted', '#85857E'),
  grade: () => corDoToken('--border', '#D9D9D4'),
  eixo: () => corDoToken('--text-secondary', '#5F5F5A'),
}

export default function ChartsPage() {
  // Assina o contexto de tema para que esta página re-renderize quando o
  // usuário alterna claro/escuro. Sem isso, CORES_GRAFICO/corDoToken (que
  // leem getComputedStyle no momento do render) nunca seriam chamados de
  // novo, e os gráficos ficariam com a paleta do tema anterior até algo
  // não relacionado forçar um re-render (ex.: "Aplicar Filtro").
  //
  // NÃO REMOVA esta chamada mesmo que pareça "não usada": o valor de retorno
  // não é lido em lugar nenhum (por isso não é desestruturado), mas a
  // CHAMADA de useTheme() é o que registra a assinatura ao contexto —
  // é ela quem dispara o re-render. tsconfig.json tem noUnusedLocals:false
  // e o projeto não tem eslint, então nada aqui acusaria erro de build se
  // esta linha for apagada numa limpeza de "variável não usada": tsc, os
  // 132+ testes e o vite build continuariam verdes, e os gráficos
  // simplesmente parariam de acompanhar a troca de tema silenciosamente
  // (nenhum teste cobre ChartsPage, e o jsdom não calcula CSS mesmo que
  // cobrisse).
  useTheme()

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
      <PageHeader
        eyebrow="Business Intelligence"
        eyebrowDetail="Indicadores Operacionais"
        title="Dashboard & Indicadores"
        description="Análise temporal de empréstimos, devoluções e novos cadastros de equipamentos"
      />

      {/* Filter Control Bar */}
      <div className="flex flex-wrap items-end gap-4 surface-panel p-4">
        <div className="flex items-center gap-2 text-foreground text-caption pr-2">
          <Filter size={16} />
          <span>Filtrar Período</span>
        </div>
        <div className="flex flex-col gap-1.5">
          <Label>Ano</Label>
          <Input
            value={year}
            onChange={(e) => setYear(e.target.value)}
            className="w-28"
          />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label>Mês</Label>
          <Select value={month} onValueChange={setMonth}>
            <SelectTrigger className="w-44">
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
          variant="gradient"
          onClick={() => setParams({ year: Number(year), month: Number(month) })}
        >
          <Calendar size={15} />
          Aplicar Filtro
        </Button>
      </div>

      {/* Summary KPI Cards for Selected Month */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="surface-panel p-4">
          <StatBlock label="Empréstimos no Mês" value={totalEmprestimos} symbol="↑" />
        </div>

        <div className="surface-panel p-4">
          <StatBlock label="Devoluções no Mês" value={totalDevolucoes} symbol="↓" />
        </div>

        <div className="surface-panel p-4">
          <StatBlock label="Cadastros no Mês" value={totalCadastros} symbol="#" />
        </div>
      </div>

      {/* Gráfico 1: Empréstimos x Devoluções */}
      <div className="surface-panel p-6 space-y-4">
        <div className="flex items-center justify-between">
          <div>
            <h3 className="text-body-lg font-semibold text-foreground">
              Movimentação de Empréstimos × Devoluções por Dia
            </h3>
            <p className="text-caption text-muted-foreground">
              Comparativo de saídas e retornos de equipamentos em {MONTHS[params.month - 1]} de {params.year}
            </p>
          </div>
        </div>

        {loansLoading ? (
          <div className="h-72 flex items-center justify-center gap-2 text-body-sm text-muted-foreground">
            <Loader2 className="h-4 w-4 animate-spin" />
            Carregando gráfico...
          </div>
        ) : (
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={buildLoansChartData()} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={CORES_GRAFICO.grade()} vertical={false} />
              <XAxis dataKey="dia" tick={{ fontSize: 11, fill: CORES_GRAFICO.eixo() }} axisLine={false} tickLine={false} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: CORES_GRAFICO.eixo() }} axisLine={false} tickLine={false} />
              <Tooltip
                contentStyle={{
                  backgroundColor: corDoToken('--surface', '#FFFFFF'),
                  borderColor: corDoToken('--border', '#D9D9D4'),
                  borderRadius: 'var(--radius-md)',
                  color: corDoToken('--text-primary', '#111111'),
                  fontSize: '12px',
                  fontWeight: '600',
                }}
              />
              <Legend wrapperStyle={{ paddingTop: '10px', fontSize: '12px', fontWeight: '600' }} />
              <Bar dataKey="Empréstimos" fill={CORES_GRAFICO.serie1()} radius={[6, 6, 0, 0]} maxBarSize={28} />
              <Bar dataKey="Devoluções" fill={CORES_GRAFICO.serie2()} radius={[6, 6, 0, 0]} maxBarSize={28} />
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>

      {/* Gráfico 2: Novos Cadastros */}
      <div className="surface-panel p-6 space-y-4">
        <div>
          <h3 className="text-body-lg font-semibold text-foreground">
            Novos Cadastros de Equipamentos por Dia
          </h3>
          <p className="text-caption text-muted-foreground">
            Volume de inclusões de patrimônios no acervo em {MONTHS[params.month - 1]} de {params.year}
          </p>
        </div>

        {regLoading ? (
          <div className="h-72 flex items-center justify-center gap-2 text-body-sm text-muted-foreground">
            <Loader2 className="h-4 w-4 animate-spin" />
            Carregando gráfico...
          </div>
        ) : (
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={buildRegChartData()} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={CORES_GRAFICO.grade()} vertical={false} />
              <XAxis dataKey="dia" tick={{ fontSize: 11, fill: CORES_GRAFICO.eixo() }} axisLine={false} tickLine={false} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: CORES_GRAFICO.eixo() }} axisLine={false} tickLine={false} />
              <Tooltip
                contentStyle={{
                  backgroundColor: corDoToken('--surface', '#FFFFFF'),
                  borderColor: corDoToken('--border', '#D9D9D4'),
                  borderRadius: 'var(--radius-md)',
                  color: corDoToken('--text-primary', '#111111'),
                  fontSize: '12px',
                  fontWeight: '600',
                }}
              />
              <Bar dataKey="Cadastros" fill={CORES_GRAFICO.serie1()} radius={[6, 6, 0, 0]} maxBarSize={32} />
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>
    </div>
  )
}

