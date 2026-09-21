import { useMemo, useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { getLoansChart, getRegistrationsChart, getMonthlyReport, type ReportRow } from '@/api/reports'
import { listUnidades } from '@/api/unidades'
import {
  AreaChart,
  Area,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
} from 'recharts'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { StatBlock } from '@/components/ui/StatBlock'
import { Button } from '@/components/ui/button'
import { LoadingState, EmptyState, ErrorState } from '@/components/ui/StateBlocks'
import { Calendar, Filter } from 'lucide-react'
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
  // 132+ testes e o vite build continuariam verdes, e os gráficos
  // simplesmente parariam de acompanhar a troca de tema silenciosamente
  // (nenhum teste cobre ChartsPage, e o jsdom não calcula CSS mesmo que
  // cobrisse).
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

  const loansVazio = loansChartData.every((d) => !d.Empréstimos && !d.Devoluções)
  const regVazio = regChartData.every((d) => !d.Cadastros)
  const weekdayVazio = weekdayChartData.every((d) => !d.Empréstimos)

  function aplicarFiltro() {
    if (!anoValido) return
    setParams({ year: Number(year), month: Number(month), unidade: filtroUnidade })
  }

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
            inputMode="numeric"
            maxLength={4}
            onChange={(e) => setYear(e.target.value.replace(/\D/g, '').slice(0, 4))}
            className="w-28"
            aria-invalid={!anoValido}
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
        {podeFiltrarPorUnidade && (
          <div className="flex flex-col gap-1.5">
            <Label>Unidade</Label>
            <Select value={filtroUnidade} onValueChange={setFiltroUnidade}>
              <SelectTrigger className="w-48">
                <SelectValue placeholder="Todas as unidades" />
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
          </div>
        )}
        <Button variant="gradient" onClick={aplicarFiltro} disabled={!anoValido}>
          <Calendar size={15} />
          Aplicar Filtro
        </Button>
        {!anoValido && (
          <p className="text-caption text-muted-foreground pb-2">Ano inválido (2000–2100)</p>
        )}
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
              {filtrado ? ` · ${params.unidade}` : ''}
            </p>
          </div>
        </div>

        {isLoading ? (
          <LoadingState label="Carregando gráfico..." className="h-72" />
        ) : loansError ? (
          <ErrorState
            error={loansError}
            onRetry={repetirLoans}
            title="Não foi possível carregar este gráfico"
            className="h-72"
          />
        ) : loansVazio ? (
          <EmptyState title="Nenhuma movimentação neste período." className="h-72" />
        ) : (
          <ResponsiveContainer width="100%" height={320}>
            <AreaChart data={loansChartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <defs>
                <linearGradient id="gradEmprestimos" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor={CORES_GRAFICO.serie1()} stopOpacity={0.25} />
                  <stop offset="95%" stopColor={CORES_GRAFICO.serie1()} stopOpacity={0.02} />
                </linearGradient>
                <linearGradient id="gradDevolucoes" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor={CORES_GRAFICO.serie2()} stopOpacity={0.25} />
                  <stop offset="95%" stopColor={CORES_GRAFICO.serie2()} stopOpacity={0.02} />
                </linearGradient>
              </defs>
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
              <Area
                type="monotone"
                dataKey="Empréstimos"
                stroke={CORES_GRAFICO.serie1()}
                strokeWidth={2}
                fill="url(#gradEmprestimos)"
                dot={{ r: 2, fill: CORES_GRAFICO.serie1() }}
                activeDot={{ r: 4 }}
              />
              <Area
                type="monotone"
                dataKey="Devoluções"
                stroke={CORES_GRAFICO.serie2()}
                strokeWidth={2}
                fill="url(#gradDevolucoes)"
                dot={{ r: 2, fill: CORES_GRAFICO.serie2() }}
                activeDot={{ r: 4 }}
              />
            </AreaChart>
          </ResponsiveContainer>
        )}

        {/* Legenda manual — só este primeiro gráfico tem legenda (igual à
            referência, origin/redesign-frontend); os outros dois ficam sem.
            As cores vêm de CORES_GRAFICO/corDoToken (mesmo mecanismo das
            séries do gráfico acima) via style inline, nunca de classe
            Tailwind fixa — senão não acompanharia a troca de tema. */}
        {!isLoading && !loansError && !loansVazio && (
          <div className="flex items-center gap-6 pt-3 border-t border-border text-caption">
            <div className="flex items-center gap-1.5">
              <div className="w-6 h-0.5" style={{ backgroundColor: CORES_GRAFICO.serie1() }} />
              <span className="text-foreground font-medium">Empréstimos</span>
            </div>
            <div className="flex items-center gap-1.5">
              <div className="w-6 h-0.5" style={{ backgroundColor: CORES_GRAFICO.serie2() }} />
              <span className="text-muted-foreground">Devoluções</span>
            </div>
          </div>
        )}
      </div>

      {/* Gráfico 2: Novos Cadastros + Gráfico 3: Distribuição Semanal */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="surface-panel p-6 space-y-4">
          <div>
            <h3 className="text-body-lg font-semibold text-foreground">
              Novos Cadastros de Equipamentos por Dia
            </h3>
            <p className="text-caption text-muted-foreground">
              Volume de inclusões de patrimônios no acervo em {MONTHS[params.month - 1]} de {params.year}
              {filtrado ? ` · ${params.unidade}` : ''}
            </p>
          </div>

          {isLoading ? (
            <LoadingState label="Carregando gráfico..." className="h-64" />
          ) : regError ? (
            <ErrorState
              error={regError}
              onRetry={repetirReg}
              title="Não foi possível carregar este gráfico"
              className="h-64"
            />
          ) : regVazio ? (
            <EmptyState title="Nenhuma movimentação neste período." className="h-64" />
          ) : (
            <ResponsiveContainer width="100%" height={256}>
              <BarChart data={regChartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
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

        {/* Gráfico 3: Distribuição Semanal de Empréstimos */}
        <div className="surface-panel p-6 space-y-4">
          <div>
            <h3 className="text-body-lg font-semibold text-foreground">
              Distribuição Semanal de Empréstimos
            </h3>
            <p className="text-caption text-muted-foreground">
              Concentração de movimentações por dia da semana em {MONTHS[params.month - 1]} de {params.year}
              {filtrado ? ` · ${params.unidade}` : ''}
            </p>
          </div>

          {isLoading ? (
            <LoadingState label="Carregando gráfico..." className="h-64" />
          ) : loansError ? (
            <ErrorState
              error={loansError}
              onRetry={repetirLoans}
              title="Não foi possível carregar este gráfico"
              className="h-64"
            />
          ) : weekdayVazio ? (
            <EmptyState title="Nenhuma movimentação neste período." className="h-64" />
          ) : (
            <ResponsiveContainer width="100%" height={256}>
              <BarChart data={weekdayChartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
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
                <Bar dataKey="Empréstimos" fill={CORES_GRAFICO.serie1()} radius={[6, 6, 0, 0]} maxBarSize={32} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>
    </div>
  )
}
