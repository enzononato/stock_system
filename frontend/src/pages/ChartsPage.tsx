import { useMemo, useState } from 'react'
import { isAxiosError } from 'axios'
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

import { getLoansChart, getRegistrationsChart, getMonthlyReport, type ReportRow } from '@/api/reports'
import { listUnidades } from '@/api/unidades'
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
import { useAuth } from '@/contexts/AuthContext'

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

const DIAS_SEMANA = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb']

/**
 * Mensagem de negócio para o 403 concreto desta tela: o filtro de unidade usa
 * GET /reports/monthly, que exige Gestor/Técnico no backend (ver reports.py),
 * enquanto os endpoints de gráfico não têm essa guarda de papel. A defesa
 * principal é esconder o seletor de unidade para quem não pode usá-lo (Jovem
 * Aprendiz, ver `podeFiltrarPorUnidade` abaixo); esta é a defesa secundária,
 * caso mesmo assim a requisição filtrada volte com 403 — mesmo padrão do 403
 * de senha errada em HistoryPage: mensagem de negócio no caso conhecido,
 * `getErrorMessage` (via ErrorState) em qualquer outro erro.
 */
function mensagem403Unidade(error: unknown): string | undefined {
  if (isAxiosError(error) && error.response?.status === 403) {
    return 'Seu perfil de acesso não permite ver este relatório filtrado por unidade.'
  }
  return undefined
}

/*
 * Paleta dos gráficos em escala de cinza, lida dos tokens CSS para acompanhar
 * o modo claro/escuro. recharts recebe cor por prop (fill/stroke), então não dá
 * para usar classe do Tailwind aqui — é preciso resolver o valor computado.
 *
 * Os valores reproduzem a FT_STC, que usava `var(--foreground)`,
 * `var(--muted-foreground)`, `var(--brand-text)` (antes --primary, ilegível em azul sobre o preto) e `var(--text-muted)` direto nas
 * props. --foreground e --muted-foreground são aliases de --text-primary e
 * --text-secondary (index.css), então lemos os tokens de base.
 */
function corDoToken(nome: string, alternativa: string): string {
  if (typeof window === 'undefined') return alternativa
  const valor = getComputedStyle(document.documentElement).getPropertyValue(nome).trim()
  return valor || alternativa
}

const CORES_GRAFICO = {
  /** Linha e pontos da série principal (equivale a --foreground). */
  serie1: () => corDoToken('--text-primary', '#111111'),
  /** Linha e pontos da série secundária (equivale a --muted-foreground). */
  serie2: () => corDoToken('--text-secondary', '#5F5F5A'),
  /** Degradê sob a série de empréstimos do gráfico principal. */
  degradePrimario: () => corDoToken('--brand-text', '#0131CD'),
  /** Degradê sob a série de devoluções do gráfico principal. */
  degradeSecundario: () => corDoToken('--text-muted', '#85857E'),
  grade: () => corDoToken('--border', '#D9D9D4'),
  eixo: () => corDoToken('--text-secondary', '#5F5F5A'),
}

function estiloTooltip() {
  return {
    backgroundColor: corDoToken('--surface', '#FFFFFF'),
    borderColor: corDoToken('--border', '#D9D9D4'),
    borderRadius: 4,
    fontSize: 12,
    fontWeight: 600,
    color: corDoToken('--text-primary', '#111111'),
    boxShadow: corDoToken('--shadow-overlay', '0 2px 12px rgba(0, 0, 0, 0.08)'),
  }
}

interface PontoDia {
  dia: string
  Empréstimos: number
  Devoluções: number
}

interface PontoCadastro {
  dia: string
  Cadastros: number
}

interface PontoSemana {
  dia: string
  Empréstimos: number
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
  // testes e o vite build continuariam verdes, e os gráficos simplesmente
  // parariam de acompanhar a troca de tema silenciosamente (nenhum teste
  // cobre ChartsPage, e o jsdom não calcula CSS mesmo que cobrisse).
  useTheme()

  const { hasRole } = useAuth()
  // GET /reports/monthly (usado pelo filtro de unidade, abaixo) é restrito a
  // Gestor/Técnico no backend (reports.py) — os endpoints de gráfico não têm
  // essa guarda. Jovem Aprendiz é o único papel sem acesso; em vez de deixar
  // essa pessoa escolher uma unidade e receber um 403 na tela de Indicadores
  // (que ela abre normalmente), o seletor nem aparece para esse papel.
  const podeFiltrarPorUnidade = !hasRole('Jovem Aprendiz')

  const now = new Date()
  const [year, setYear] = useState(String(now.getFullYear()))
  const [month, setMonth] = useState(String(now.getMonth() + 1))
  const [filtroUnidade, setFiltroUnidade] = useState('all')
  const [params, setParams] = useState({
    year: now.getFullYear(),
    month: now.getMonth() + 1,
    unidade: 'all',
  })

  const anoValido = /^\d{4}$/.test(year) && Number(year) >= 2000 && Number(year) <= 2100

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades-charts'],
    queryFn: () => listUnidades(),
    // Só interessa a quem de fato pode usar o filtro — evita uma requisição
    // que o Jovem Aprendiz nunca vai usar (o seletor não aparece para ele).
    enabled: podeFiltrarPorUnidade,
  })

  const filtrado = params.unidade !== 'all'

  const loansQuery = useQuery({
    queryKey: ['chart-loans', params.year, params.month],
    queryFn: () => getLoansChart(params.year, params.month),
  })

  const regQuery = useQuery({
    queryKey: ['chart-registrations', params.year, params.month],
    queryFn: () => getRegistrationsChart(params.year, params.month),
  })

  // O filtro de unidade não é aceito por /charts/loans nem /charts/registrations
  // (ver reports.py) — a saída é buscar o relatório mensal completo, que traz
  // `revenda`, e reagregar por dia no cliente só quando uma unidade específica
  // está selecionada.
  const monthlyQuery = useQuery({
    queryKey: ['chart-monthly-report', params.year, params.month],
    queryFn: () => getMonthlyReport(params.year, params.month),
    enabled: filtrado,
  })

  const isLoading = filtrado ? monthlyQuery.isLoading : loansQuery.isLoading || regQuery.isLoading
  const loansError = filtrado ? monthlyQuery.error : loansQuery.error
  const regError = filtrado ? monthlyQuery.error : regQuery.error

  function repetirLoans() {
    if (filtrado) void monthlyQuery.refetch()
    else void loansQuery.refetch()
  }

  function repetirReg() {
    if (filtrado) void monthlyQuery.refetch()
    else void regQuery.refetch()
  }

  const diasNoMes = new Date(params.year, params.month, 0).getDate()

  const { loansChartData, regChartData } = useMemo(() => {
    if (!filtrado) {
      const loansData = loansQuery.data
      const registrationsData = regQuery.data
      const lData: PontoDia[] = (loansData?.days ?? []).map((day, i) => ({
        dia: `Dia ${day}`,
        Empréstimos: loansData?.values[i] ?? 0,
        Devoluções: loansData?.values2?.[i] ?? 0,
      }))
      const rData: PontoCadastro[] = (registrationsData?.days ?? []).map((day, i) => ({
        dia: `Dia ${day}`,
        Cadastros: registrationsData?.values[i] ?? 0,
      }))
      return { loansChartData: lData, regChartData: rData }
    }

    const linhas: ReportRow[] = (monthlyQuery.data ?? []).filter((r) => r.revenda === params.unidade)
    const emprestimosPorDia: Record<number, number> = {}
    const devolucoesPorDia: Record<number, number> = {}
    const cadastrosPorDia: Record<number, number> = {}

    linhas.forEach((r) => {
      // `data_emprestimo` é o mesmo slot de coluna reaproveitado pelo UNION do
      // relatório mensal para TODA linha (cadastro, periférico, exclusão...),
      // não só empréstimo — por isso é obrigatório checar `operation_type`
      // antes de contar, senão cadastros e outras operações também entrariam
      // aqui como se fossem empréstimos.
      if (r.operation_type === 'Empréstimo' && r.data_emprestimo) {
        const dia = new Date(r.data_emprestimo).getDate()
        emprestimosPorDia[dia] = (emprestimosPorDia[dia] ?? 0) + 1
      }

      // `data_devolucao` vem de uma subconsulta sem limite de data (ver
      // inventory_manager_db.py) — uma máquina emprestada no mês selecionado
      // mas devolvida em outro mês apareceria aqui como devolução do dia
      // correspondente NAQUELE outro mês. Só conta se a devolução caiu de
      // fato dentro do ano/mês filtrado, senão este gráfico discordaria do
      // não-filtrado (que usa /charts/loans, sempre restrito ao mês).
      if (r.data_devolucao) {
        const dataDevolucao = new Date(r.data_devolucao)
        if (dataDevolucao.getFullYear() === params.year && dataDevolucao.getMonth() + 1 === params.month) {
          const dia = dataDevolucao.getDate()
          devolucoesPorDia[dia] = (devolucoesPorDia[dia] ?? 0) + 1
        }
      }

      if (r.operation_type === 'Cadastro' && r.data_emprestimo) {
        const dia = new Date(r.data_emprestimo).getDate()
        cadastrosPorDia[dia] = (cadastrosPorDia[dia] ?? 0) + 1
      }
    })

    const lData: PontoDia[] = Array.from({ length: diasNoMes }, (_, i) => {
      const dia = i + 1
      return {
        dia: `Dia ${dia}`,
        Empréstimos: emprestimosPorDia[dia] ?? 0,
        Devoluções: devolucoesPorDia[dia] ?? 0,
      }
    })
    const rData: PontoCadastro[] = Array.from({ length: diasNoMes }, (_, i) => {
      const dia = i + 1
      return { dia: `Dia ${dia}`, Cadastros: cadastrosPorDia[dia] ?? 0 }
    })

    return { loansChartData: lData, regChartData: rData }
  }, [filtrado, loansQuery.data, regQuery.data, monthlyQuery.data, params.unidade, params.year, params.month, diasNoMes])

  // Gráfico 3: distribuição por dia da semana, derivada do próprio
  // `loansChartData` acima (que já é a fonte de dados de empréstimos em uso,
  // filtrada ou não) — sem nenhuma requisição nova.
  const weekdayChartData: PontoSemana[] = useMemo(() => {
    const contagens = Array(7).fill(0)
    loansChartData.forEach((ponto, idx) => {
      const dia = idx + 1
      const data = new Date(params.year, params.month - 1, dia)
      contagens[data.getDay()] += ponto.Empréstimos
    })
    return DIAS_SEMANA.map((label, i) => ({ dia: label, Empréstimos: contagens[i] ?? 0 }))
  }, [loansChartData, params.year, params.month])

  // Totais do mês selecionado
  const totalEmprestimos = loansChartData.reduce((acc, d) => acc + d.Empréstimos, 0)
  const totalDevolucoes = loansChartData.reduce((acc, d) => acc + d.Devoluções, 0)
  const totalCadastros = regChartData.reduce((acc, d) => acc + d.Cadastros, 0)

  const periodoLegenda = `${MONTHS[params.month - 1] ?? 'Mês'} / ${params.year}${
    filtrado ? ` · ${params.unidade}` : ''
  }`

  const loansVazio = loansChartData.every((d) => !d.Empréstimos && !d.Devoluções)
  const regVazio = regChartData.every((d) => !d.Cadastros)
  const weekdayVazio = weekdayChartData.every((d) => !d.Empréstimos)

  function aplicarFiltro() {
    if (!anoValido) return
    setParams({ year: Number(year), month: Number(month), unidade: filtroUnidade })
  }

  const tooltip = estiloTooltip()

  return (
    <div className="page-container-dense space-y-6">
      <PageHeader
        eyebrow="Inteligência de Operações · Análise Comparativa"
        title="Indicadores Operacionais"
        description="Movimentação patrimonial comparativa por período, filial e tipo de operação."
      />

      {/* Barra de filtro */}
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
              aria-invalid={!anoValido}
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

          {/* Seletor de unidade só para quem pode chamar GET /reports/monthly. */}
          {podeFiltrarPorUnidade && (
            <div className="space-y-1">
              <Label className="text-caption font-medium">Unidade / Filial</Label>
              <Select value={filtroUnidade} onValueChange={setFiltroUnidade}>
                <SelectTrigger className="w-52 h-9 rounded-sm border-border bg-background text-body-sm">
                  <SelectValue placeholder="Todas" />
                </SelectTrigger>
                <SelectContent className="rounded-sm border-border bg-surface">
                  <SelectItem value="all" className="text-body-sm">
                    Todas as unidades
                  </SelectItem>
                  {unidades.map((u) => (
                    <SelectItem key={u.id} value={u.nome} className="text-body-sm">
                      {u.nome}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
          )}

          <Button
            size="sm"
            onClick={aplicarFiltro}
            disabled={!anoValido}
            className="h-9 rounded-sm bg-primary text-primary-foreground hover:opacity-90 font-medium px-4"
          >
            <Calendar className="mr-1.5 size-4" />
            Aplicar Filtro
          </Button>

          {!anoValido && (
            <p className="text-caption text-destructive self-end pb-1">Ano inválido (2000–2100)</p>
          )}
        </div>
      </div>

      {/* KPIs numéricos */}
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
              <span className="font-mono text-heading-lg font-bold text-foreground">{stat.value}</span>
              <span className="font-mono text-body-lg text-muted-foreground">{stat.symbol}</span>
            </div>
            <span className="text-caption text-muted-foreground mt-1 block">{periodoLegenda}</span>
          </div>
        ))}
      </div>

      {/* Gráfico 1: área de empréstimos × devoluções */}
      <div className="surface-panel p-6 space-y-4">
        <div className="border-b border-border pb-3">
          <h2 className="text-body-lg font-semibold text-foreground">
            Movimentação de Empréstimos × Devoluções por Dia
          </h2>
          <p className="text-caption text-muted-foreground mt-0.5">
            Comparativo de saídas e retornos de equipamentos em {periodoLegenda}
          </p>
        </div>

        {isLoading ? (
          <LoadingState label="Carregando indicadores..." />
        ) : loansError ? (
          <ErrorState error={loansError} message={mensagem403Unidade(loansError)} onRetry={repetirLoans} />
        ) : loansVazio ? (
          <EmptyState
            title="Nenhuma movimentação registrada neste período"
            description="Tente selecionar outro mês ou ano no filtro acima."
          />
        ) : (
          <div className="min-w-0 overflow-x-auto">
            <div className="min-w-[540px]">
              <ResponsiveContainer width="100%" height={320}>
                <AreaChart data={loansChartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                  <defs>
                    <linearGradient id="gradEmp" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor={CORES_GRAFICO.degradePrimario()} stopOpacity={0.25} />
                      <stop offset="95%" stopColor={CORES_GRAFICO.degradePrimario()} stopOpacity={0.02} />
                    </linearGradient>
                    <linearGradient id="gradDev" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor={CORES_GRAFICO.degradeSecundario()} stopOpacity={0.2} />
                      <stop offset="95%" stopColor={CORES_GRAFICO.degradeSecundario()} stopOpacity={0.01} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={CORES_GRAFICO.grade()} vertical={false} />
                  <XAxis
                    dataKey="dia"
                    tick={{ fontSize: 11, fill: CORES_GRAFICO.eixo() }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <YAxis
                    allowDecimals={false}
                    tick={{ fontSize: 11, fill: CORES_GRAFICO.eixo() }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip contentStyle={tooltip} />
                  <Area
                    type="monotone"
                    dataKey="Empréstimos"
                    stroke={CORES_GRAFICO.serie1()}
                    strokeWidth={2}
                    fill="url(#gradEmp)"
                    dot={{ r: 2.5, fill: CORES_GRAFICO.serie1() }}
                    activeDot={{ r: 4.5 }}
                  />
                  <Area
                    type="monotone"
                    dataKey="Devoluções"
                    stroke={CORES_GRAFICO.serie2()}
                    strokeWidth={2}
                    fill="url(#gradDev)"
                    dot={{ r: 2.5, fill: CORES_GRAFICO.serie2() }}
                    activeDot={{ r: 4.5 }}
                  />
                </AreaChart>
              </ResponsiveContainer>

              {/* Legenda manual monocromática — classes de token, acompanham o tema. */}
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

      {/* Grade: cadastros por dia + distribuição semanal */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Gráfico 2: área de cadastros por dia */}
        <div className="surface-panel p-5">
          <div className="border-b border-border pb-3 mb-4">
            <h2 className="text-body-lg font-semibold text-foreground">Cadastros por Dia</h2>
            <p className="text-caption text-muted-foreground mt-0.5">
              Novos patrimônios registrados no período.
            </p>
          </div>

          {isLoading ? (
            <LoadingState label="Carregando cadastros..." />
          ) : regError ? (
            <ErrorState error={regError} message={mensagem403Unidade(regError)} onRetry={repetirReg} />
          ) : regVazio ? (
            <EmptyState title="Nenhum cadastro neste período." />
          ) : (
            <div className="overflow-x-auto">
              <div className="min-w-[300px]">
                <ResponsiveContainer width="100%" height={230}>
                  <AreaChart data={regChartData} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
                    <defs>
                      <linearGradient id="gradCad" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%" stopColor={CORES_GRAFICO.serie1()} stopOpacity={0.22} />
                        <stop offset="95%" stopColor={CORES_GRAFICO.serie1()} stopOpacity={0.01} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={CORES_GRAFICO.grade()} vertical={false} />
                    <XAxis
                      dataKey="dia"
                      tick={{ fontSize: 10, fill: CORES_GRAFICO.eixo() }}
                      axisLine={false}
                      tickLine={false}
                    />
                    <YAxis
                      allowDecimals={false}
                      tick={{ fontSize: 10, fill: CORES_GRAFICO.eixo() }}
                      axisLine={false}
                      tickLine={false}
                    />
                    <Tooltip contentStyle={tooltip} />
                    <Area
                      type="monotone"
                      dataKey="Cadastros"
                      stroke={CORES_GRAFICO.serie1()}
                      strokeWidth={2}
                      fill="url(#gradCad)"
                      dot={{ r: 2.5, fill: CORES_GRAFICO.serie1() }}
                      activeDot={{ r: 4.5 }}
                    />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}
        </div>

        {/* Gráfico 3: área de distribuição por dia da semana */}
        <div className="surface-panel p-5">
          <div className="border-b border-border pb-3 mb-4">
            <h2 className="text-body-lg font-semibold text-foreground">Distribuição Semanal de Empréstimos</h2>
            <p className="text-caption text-muted-foreground mt-0.5">
              Concentração de saídas por dia da semana.
            </p>
          </div>

          {isLoading ? (
            <LoadingState label="Carregando indicadores..." />
          ) : loansError ? (
            <ErrorState error={loansError} message={mensagem403Unidade(loansError)} onRetry={repetirLoans} />
          ) : weekdayVazio ? (
            <EmptyState title="Nenhum dado disponível para este período." />
          ) : (
            <ResponsiveContainer width="100%" height={230}>
              <AreaChart data={weekdayChartData} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
                <defs>
                  <linearGradient id="gradSem" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor={CORES_GRAFICO.serie1()} stopOpacity={0.2} />
                    <stop offset="95%" stopColor={CORES_GRAFICO.serie1()} stopOpacity={0.01} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke={CORES_GRAFICO.grade()} vertical={false} />
                <XAxis
                  dataKey="dia"
                  tick={{ fontSize: 11, fill: CORES_GRAFICO.eixo() }}
                  axisLine={false}
                  tickLine={false}
                />
                <YAxis
                  allowDecimals={false}
                  tick={{ fontSize: 10, fill: CORES_GRAFICO.eixo() }}
                  axisLine={false}
                  tickLine={false}
                />
                <Tooltip contentStyle={tooltip} />
                <Area
                  type="monotone"
                  dataKey="Empréstimos"
                  stroke={CORES_GRAFICO.serie1()}
                  strokeWidth={2}
                  fill="url(#gradSem)"
                  dot={{ r: 2.5, fill: CORES_GRAFICO.serie1() }}
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
