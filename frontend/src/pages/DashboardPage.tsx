import { useMemo } from 'react'
import { useIsFetching, useQuery, useQueryClient } from '@tanstack/react-query'
import { Link } from 'react-router-dom'
import {
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
} from 'recharts'
import {
  ArrowRight,
  Boxes,
  FileSignature,
  Loader2,
  PackageCheck,
  RefreshCw,
  TrendingUp,
  Undo2,
  Wrench,
} from 'lucide-react'

import { getAlertasPainel, getResumoPainel } from '@/api/painel'
import { getLoansChart } from '@/api/reports'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { useAuth } from '@/contexts/AuthContext'
import { EquipamentosPorUnidade } from '@/features/painel/EquipamentosPorUnidade'
import { IndicadoresPainel } from '@/features/painel/IndicadoresPainel'
import { ListaAlertas } from '@/features/painel/ListaAlertas'
import { useTheme } from '@/lib/theme'
import { formatDateTime } from '@/lib/utils'

interface PontoDia {
  dia: string
  Empréstimos: number
  Devoluções: number
}

/**
 * Mesma leitura de cor por token CSS usada em ChartsPage: recharts recebe cor
 * via prop (fill/stroke), não por classe Tailwind, então é preciso resolver o
 * valor computado da custom property em vez de vincular uma classe.
 */
function corDoToken(nome: string, alternativa: string): string {
  if (typeof window === 'undefined') return alternativa
  const valor = getComputedStyle(document.documentElement).getPropertyValue(nome).trim()
  return valor || alternativa
}

/*
 * Mesmas cores da FT_STC, que usava `var(--foreground)` e
 * `var(--muted-foreground)` direto nas props do recharts. Esses dois tokens
 * são aliases de --text-primary e --text-secondary (index.css), então lemos
 * os tokens de base — o resultado visual é idêntico e continua acompanhando
 * a troca de tema.
 */
const CORES_GRAFICO = {
  serie1: () => corDoToken('--text-primary', '#111111'),
  serie2: () => corDoToken('--text-secondary', '#5F5F5A'),
  grade: () => corDoToken('--border', '#D9D9D4'),
  eixo: () => corDoToken('--text-secondary', '#5F5F5A'),
}

interface LinkRapido {
  to: string
  label: string
  description: string
  icon: typeof Boxes
  /** Sem `roles`: visível para qualquer papel (mesma convenção da Sidebar). */
  roles?: string[]
}

const LINKS_RAPIDOS: LinkRapido[] = [
  {
    to: '/stock',
    label: 'Consultar Estoque',
    description: 'Tabela densa de patrimônios',
    icon: Boxes,
  },
  {
    to: '/loan',
    label: 'Novo Empréstimo',
    description: 'Atribuir ativo e gerar termo',
    icon: PackageCheck,
    roles: ['Gestor', 'Técnico'],
  },
  {
    to: '/return',
    label: 'Devolução Física',
    description: 'Conferir periféricos e estado',
    icon: Undo2,
    roles: ['Gestor', 'Técnico'],
  },
  {
    to: '/terms',
    label: 'Termos de Responsabilidade',
    description: 'Assinaturas e PDFs arquivados',
    icon: FileSignature,
    roles: ['Gestor', 'Técnico'],
  },
  {
    to: '/manutencoes',
    label: 'Manutenções',
    description: 'Ordens abertas e custo',
    icon: Wrench,
    roles: ['Gestor', 'Técnico'],
  },
]

export default function DashboardPage() {
  // Assina o tema só para re-renderizar quando o usuário troca claro/escuro —
  // mesmo mecanismo do ChartsPage. Sem isto o gráfico de área congelaria na
  // paleta anterior: corDoToken() só é chamado de novo dentro de um render, e
  // nada mais aqui dispararia um quando o tema muda sozinho.
  useTheme()

  const { hasRole } = useAuth()
  // Os alertas são pendências que só Gestor e Técnico resolvem (e trazem nome de
  // colaborador desligado): o endpoint é restrito a eles, então o Jovem Aprendiz nem
  // dispara a consulta (sem 403 previsível) e vê só os indicadores.
  const podeVerAlertas = hasRole('Gestor', 'Técnico')

  const queryClient = useQueryClient()
  const atualizando = useIsFetching({ queryKey: ['painel'] }) > 0

  const now = new Date()
  const currentYear = now.getFullYear()
  const currentMonth = now.getMonth() + 1

  // Todos os números vêm do servidor (sem o limite de 500 itens do cliente).
  const resumoQuery = useQuery({
    queryKey: ['painel', 'resumo'],
    queryFn: getResumoPainel,
  })
  const alertasQuery = useQuery({
    queryKey: ['painel', 'alertas'],
    queryFn: getAlertasPainel,
    enabled: podeVerAlertas,
  })

  const loansChartQuery = useQuery({
    queryKey: ['dashboard', 'chart-loans', currentYear, currentMonth],
    queryFn: () => getLoansChart(currentYear, currentMonth),
  })

  // Últimos 15 dias do gráfico mensal de empréstimos x devoluções. O
  // endpoint (`/reports/charts/loans`) devolve o mês inteiro; o recorte para
  // os últimos 15 dias é feito aqui, DEPOIS de parear cada dia com seu valor
  // pelo índice — recortar `days.slice(-15)` primeiro e só então indexar o
  // array `values` original pelo novo índice desalinha dia e valor em
  // qualquer mês com mais de 15 dias de movimentação.
  const chartData = useMemo<PontoDia[]>(() => {
    const dias = loansChartQuery.data?.days ?? []
    const valores = loansChartQuery.data?.values ?? []
    const valores2 = loansChartQuery.data?.values2 ?? []
    // Sem dados no mês: em vez de uma mensagem vazia, o gráfico mostra 15 dias
    // zerados (eixo e grade continuam visíveis).
    if (dias.length === 0) {
      return Array.from({ length: 15 }, (_, i) => ({
        dia: `D${i + 1}`,
        Empréstimos: 0,
        Devoluções: 0,
      }))
    }
    const completo: PontoDia[] = dias.map((dia, i) => ({
      dia: `D${dia}`,
      Empréstimos: valores[i] ?? 0,
      Devoluções: valores2[i] ?? 0,
    }))
    return completo.slice(-15)
  }, [loansChartQuery.data])

  const linksVisiveis = LINKS_RAPIDOS.filter((link) => !link.roles || hasRole(...link.roles))

  const classePainel = 'bg-surface border border-border rounded-md p-5 space-y-4'
  const resumo = resumoQuery.data

  return (
    <div className="space-y-8">
      <PageHeader
        eyebrow="SISTEMA DE GESTÃO PATRIMONIAL"
        title="Painel"
        description={resumo ? `Atualizado em ${formatDateTime(resumo.gerado_em)}` : undefined}
        actions={
          <Button
            type="button"
            variant="outline"
            size="sm"
            loading={atualizando}
            onClick={() => queryClient.refetchQueries({ queryKey: ['painel'] })}
          >
            {!atualizando && <RefreshCw />}
            Atualizar
          </Button>
        }
      />

      {resumoQuery.isLoading ? (
        <LoadingState label="Carregando indicadores..." />
      ) : resumoQuery.isError || !resumo ? (
        <ErrorState
          error={resumoQuery.error}
          title="Não foi possível carregar o painel"
          onRetry={() => resumoQuery.refetch()}
        />
      ) : (
        <>
          <IndicadoresPainel resumo={resumo} />

          <div className="grid grid-cols-1 items-start gap-6 lg:grid-cols-12">
            {podeVerAlertas && (
              <div className={`lg:col-span-7 ${classePainel}`}>
                {alertasQuery.isLoading ? (
                  <LoadingState label="Carregando alertas..." className="py-8" />
                ) : alertasQuery.isError || !alertasQuery.data ? (
                  <ErrorState
                    error={alertasQuery.error}
                    title="Não foi possível carregar os alertas"
                    onRetry={() => alertasQuery.refetch()}
                    className="py-8"
                  />
                ) : (
                  <ListaAlertas dados={alertasQuery.data} />
                )}
              </div>
            )}

            <div className={`${podeVerAlertas ? 'lg:col-span-5' : 'lg:col-span-12'} ${classePainel}`}>
              <div className="border-b border-border pb-3">
                <h2 className="text-body font-semibold text-foreground">Equipamentos por unidade</h2>
                <p className="text-body-sm text-muted-foreground">Ativos cadastrados e situação atual</p>
              </div>
              <EquipamentosPorUnidade unidades={resumo.equipamentos.por_unidade} />
            </div>
          </div>
        </>
      )}

      {/* Seção inferior: movimentação e acesso rápido */}
      <div className="grid grid-cols-1 items-start gap-6 lg:grid-cols-12">
        <div className={`lg:col-span-8 ${classePainel}`}>
          <div className="flex items-center justify-between border-b border-border pb-3">
            <div>
              <h2 className="text-body font-semibold text-foreground flex items-center gap-2">
                <TrendingUp className="size-4 text-muted-foreground" />
                Movimentação Operacional Diária
              </h2>
              <p className="text-body-sm text-muted-foreground">
                Fluxo de empréstimos e devoluções nos últimos dias
              </p>
            </div>
            <div className="flex items-center gap-3 text-xs font-mono">
              <span className="flex items-center gap-1">
                <span className="size-2 rounded-full bg-foreground" /> Empréstimos
              </span>
              <span className="flex items-center gap-1 text-muted-foreground">
                <span className="size-2 rounded-full bg-muted-foreground border border-border" /> Devoluções
              </span>
            </div>
          </div>

          <div className="h-64 w-full pt-2">
            {loansChartQuery.isLoading ? (
              <div className="h-full flex items-center justify-center gap-2 text-body-sm text-muted-foreground">
                <Loader2 className="h-4 w-4 animate-spin" />
                Carregando gráfico...
              </div>
            ) : loansChartQuery.isError ? (
              <div className="h-full flex flex-col items-center justify-center gap-3 text-body-sm text-muted-foreground">
                <p>Não foi possível carregar este gráfico.</p>
                <Button variant="outline" size="sm" onClick={() => loansChartQuery.refetch()}>
                  <RefreshCw size={14} />
                  Tentar novamente
                </Button>
              </div>
            ) : (
              <ResponsiveContainer width="100%" height="100%">
                <AreaChart data={chartData} margin={{ top: 5, right: 10, left: -20, bottom: 0 }}>
                  {/* Degradê vertical sob cada linha; as cores vêm de corDoToken
                      para acompanhar a troca de tema (ver useTheme acima). */}
                  <defs>
                    <linearGradient id="colorEmp" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor={CORES_GRAFICO.serie1()} stopOpacity={0.25} />
                      <stop offset="95%" stopColor={CORES_GRAFICO.serie1()} stopOpacity={0} />
                    </linearGradient>
                    <linearGradient id="colorDev" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor={CORES_GRAFICO.serie2()} stopOpacity={0.2} />
                      <stop offset="95%" stopColor={CORES_GRAFICO.serie2()} stopOpacity={0} />
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
                  <Tooltip
                    contentStyle={{
                      backgroundColor: corDoToken('--surface', '#FFFFFF'),
                      borderColor: corDoToken('--border', '#D9D9D4'),
                      borderRadius: 6,
                      fontSize: 12,
                      fontWeight: 600,
                      color: corDoToken('--text-primary', '#111111'),
                      boxShadow: corDoToken('--shadow-overlay', '0 2px 12px rgba(0, 0, 0, 0.08)'),
                    }}
                  />
                  <Area
                    type="monotone"
                    dataKey="Empréstimos"
                    stroke={CORES_GRAFICO.serie1()}
                    strokeWidth={2}
                    fill="url(#colorEmp)"
                  />
                  <Area
                    type="monotone"
                    dataKey="Devoluções"
                    stroke={CORES_GRAFICO.serie2()}
                    strokeWidth={1.5}
                    strokeDasharray="3 3"
                    fill="url(#colorDev)"
                  />
                </AreaChart>
              </ResponsiveContainer>
            )}
          </div>

          <div className="pt-3 border-t border-border flex items-center justify-between text-xs text-muted-foreground">
            <span>Valores consolidados em tempo real pelo banco de dados</span>
            <Link
              to="/charts"
              className="text-foreground hover:underline inline-flex items-center gap-1 font-medium"
            >
              Ver indicadores detalhados <ArrowRight className="size-3" />
            </Link>
          </div>
        </div>

        {/* Acesso operacional rápido (4 colunas) */}
        <div className={`lg:col-span-4 ${classePainel}`}>
          <div className="border-b border-border pb-3">
            <h2 className="text-body font-semibold text-foreground">Fluxos Rápidos de TI</h2>
            <p className="text-body-sm text-muted-foreground">
              Acesso direto às operações rotineiras
            </p>
          </div>

          <div className="space-y-2">
            {linksVisiveis.map(({ to, label, description, icon: Icon }) => (
              <Link
                key={to}
                to={to}
                className="flex items-center justify-between p-3 rounded border border-border hover:bg-surface-alt transition-colors group"
              >
                <div className="flex items-center gap-2.5">
                  <Icon className="size-4 text-muted-foreground group-hover:text-foreground" />
                  <div>
                    <p className="text-body-sm font-medium text-foreground">{label}</p>
                    <p className="text-[11px] text-muted-foreground">{description}</p>
                  </div>
                </div>
                <ArrowRight className="size-3.5 text-muted-foreground group-hover:text-foreground" />
              </Link>
            ))}
          </div>
        </div>
      </div>
    </div>
  )
}
