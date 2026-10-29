import { useMemo } from 'react'
import { useQuery } from '@tanstack/react-query'
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
  AlertCircle,
  ArrowRight,
  Boxes,
  Clock,
  ExternalLink,
  FileSignature,
  Loader2,
  PackageCheck,
  RefreshCw,
  TrendingUp,
  Undo2,
} from 'lucide-react'

import { listItemsPaginated } from '@/api/items'
import { getLoansChart } from '@/api/reports'
import { listHistoryPaginated } from '@/api/history'
import { StatBlock } from '@/components/ui/StatBlock'
import { Button } from '@/components/ui/button'
import { useAuth } from '@/contexts/AuthContext'
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
]

export default function DashboardPage() {
  // Assina o tema só para re-renderizar quando o usuário troca claro/escuro —
  // mesmo mecanismo do ChartsPage. Sem isto o gráfico de área congelaria na
  // paleta anterior: corDoToken() só é chamado de novo dentro de um render, e
  // nada mais aqui dispararia um quando o tema muda sozinho.
  useTheme()

  const { hasRole } = useAuth()
  // "Atenção Operacional" (botões para /terms e /return) e "Últimas
  // Movimentações" (GET /history) exigem Gestor ou Técnico no backend —
  // Jovem Aprendiz não acessa nem as rotas de destino (App.tsx) nem o
  // endpoint de histórico (history.py, gestor_or_tecnico), que devolveria
  // 403. Em vez de um link morto ou uma requisição fadada a falhar, as duas
  // seções somem inteiras para esse papel — os quatro totais do topo
  // continuam visíveis para todos, pois vêm de GET /items (liberado a
  // qualquer usuário autenticado).
  const podeVerOperacoes = hasRole('Gestor', 'Técnico')

  const now = new Date()
  const currentYear = now.getFullYear()
  const currentMonth = now.getMonth() + 1

  // Quatro totais do topo: cada um lê só o `total` de uma consulta dedicada
  // com `limit` mínimo — nunca soma um array carregado no cliente. Mesmo
  // padrão de StockPage (ver comentário lá): antes, telas como esta
  // agregavam sobre `listItemsPaginated({ limit: 500 })`, e acima de 500
  // itens ativos os números ficavam errados em silêncio.
  const { data: totalGeralData, isLoading: totalGeralLoading } = useQuery({
    queryKey: ['dashboard', 'items', 'count', 'total'],
    queryFn: () => listItemsPaginated({ limit: 1 }),
  })
  const { data: disponiveisData, isLoading: disponiveisLoading } = useQuery({
    queryKey: ['dashboard', 'items', 'count', 'Disponível'],
    queryFn: () => listItemsPaginated({ status: 'Disponível', limit: 1 }),
  })
  const { data: indisponiveisData, isLoading: indisponiveisLoading } = useQuery({
    queryKey: ['dashboard', 'items', 'count', 'Indisponível'],
    queryFn: () => listItemsPaginated({ status: 'Indisponível', limit: 1 }),
  })
  // A lista "Atenção Operacional" mostra só os primeiros pendentes de cada
  // tipo; o `total` desta mesma consulta paginada também alimenta o quarto
  // bloco do topo ("Pendências Operacionais") — uma única requisição por
  // status cobre as duas necessidades, sem buscar a lista inteira.
  const { data: pendenteData, isLoading: pendenteLoading } = useQuery({
    queryKey: ['dashboard', 'items', 'count', 'Pendente'],
    queryFn: () => listItemsPaginated({ status: 'Pendente', limit: 4 }),
  })
  const { data: pendenteDevolucaoData, isLoading: pendenteDevolucaoLoading } = useQuery({
    queryKey: ['dashboard', 'items', 'count', 'Pendente Devolução'],
    queryFn: () => listItemsPaginated({ status: 'Pendente Devolução', limit: 3 }),
  })

  const loansChartQuery = useQuery({
    queryKey: ['dashboard', 'chart-loans', currentYear, currentMonth],
    queryFn: () => getLoansChart(currentYear, currentMonth),
  })

  const historyQuery = useQuery({
    queryKey: ['dashboard', 'recent-history'],
    queryFn: () => listHistoryPaginated({ limit: 6, offset: 0 }),
    // Jovem Aprendiz não vê esta seção (ver `podeVerOperacoes`) — desabilitar
    // a consulta evita um 403 previsível contra o backend só para descartar
    // a resposta.
    enabled: podeVerOperacoes,
  })

  const totalGeral = totalGeralData?.total ?? 0
  const disponiveisCount = disponiveisData?.total ?? 0
  const indisponiveisCount = indisponiveisData?.total ?? 0
  const pendentesConfirmacao = pendenteData?.items ?? []
  const pendentesDevolucao = pendenteDevolucaoData?.items ?? []
  const totalAtencao = (pendenteData?.total ?? 0) + (pendenteDevolucaoData?.total ?? 0)

  const totaisCarregando =
    totalGeralLoading || disponiveisLoading || indisponiveisLoading || pendenteLoading || pendenteDevolucaoLoading

  const percentualDisponivel = totalGeral > 0 ? Math.round((disponiveisCount / totalGeral) * 100) : 0

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

  const recentEvents = historyQuery.data?.items ?? []

  const linksVisiveis = LINKS_RAPIDOS.filter((link) => !link.roles || hasRole(...link.roles))

  const classePainel = 'bg-surface border border-border rounded-md p-5 space-y-4'

  return (
    <div className="space-y-8">
      {/* Cabeçalho editorial com contexto e horário */}
      <div className="border-b border-border pb-5">
        <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-2">
          <div>
            <span className="text-caption text-muted-foreground font-mono">
              SISTEMA DE GESTÃO PATRIMONIAL
            </span>
            <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
              Panorama Operacional
            </h1>
          </div>
          <div className="flex items-center gap-3 text-xs text-muted-foreground font-mono">
            <span>
              Mês Atual: {String(currentMonth).padStart(2, '0')}/{currentYear}
            </span>
            <span>·</span>
            <span className="flex items-center gap-1.5">
              <span className="size-1.5 rounded-full bg-emerald-500 animate-pulse" />
              Sincronizado
            </span>
          </div>
        </div>

        {/* Faixa superior de totais — cada um lê só o `total` de uma consulta dedicada. */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-6 pt-6 mt-4 border-t border-border">
          <StatBlock
            label="Total de Patrimônios"
            value={totaisCarregando ? '...' : totalGeral}
            hint="Ativos cadastrados"
          />
          <StatBlock
            label="Disponíveis para Alocação"
            value={totaisCarregando ? '...' : disponiveisCount}
            hint={`${percentualDisponivel}% da frota`}
          />
          <StatBlock
            label="Empréstimos Ativos"
            value={totaisCarregando ? '...' : indisponiveisCount}
            hint="Em uso com colaboradores"
          />
          <StatBlock
            label="Pendências Operacionais"
            value={totaisCarregando ? '...' : totalAtencao}
            hint={totalAtencao > 0 ? '! Atenção requerida' : '✓ Nenhuma pendência'}
          />
        </div>
      </div>

      {/* Bloco assimétrico principal: tendência em ondas + atenção operacional */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        {/* Tendência de movimentações (7 colunas; ocupa a linha toda quando
            "Atenção Operacional" não aparece para o papel do usuário) */}
        <div className={`${podeVerOperacoes ? 'lg:col-span-7' : 'lg:col-span-12'} ${classePainel}`}>
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

        {/* Atenção Operacional (5 colunas) — some para quem não acessa /terms nem /return. */}
        {podeVerOperacoes && (
          <div className={`lg:col-span-5 ${classePainel}`}>
            <div className="flex items-center justify-between border-b border-border pb-3">
              <div>
                <h2 className="text-body font-semibold text-foreground flex items-center gap-2">
                  <AlertCircle className="size-4 text-foreground" />
                  Atenção Operacional
                </h2>
                <p className="text-body-sm text-muted-foreground">
                  Ações imediatas e termos aguardando validação
                </p>
              </div>
              <span className="text-caption font-mono px-2 py-0.5 rounded-sm border border-border text-foreground font-bold">
                {totalAtencao}
              </span>
            </div>

            {totalAtencao === 0 ? (
              <div className="py-8 text-center text-body-sm text-muted-foreground space-y-1">
                <p className="font-medium text-foreground">Operação em dia</p>
                <p>Não há termos ou devoluções pendentes no momento.</p>
              </div>
            ) : (
              <div className="divide-y divide-border">
                {pendentesConfirmacao.map((p) => (
                  <div key={`confirmacao-${p.id}`} className="py-2.5 flex items-start justify-between gap-3 text-xs">
                    <div className="space-y-0.5 min-w-0">
                      <div className="flex items-center gap-2">
                        <span className="font-mono font-semibold text-foreground">#{p.id}</span>
                        <span className="text-[11px] px-1.5 py-0.5 rounded-sm border border-border font-mono text-foreground font-medium">
                          Termo Pendente
                        </span>
                      </div>
                      <p className="font-medium text-foreground truncate">
                        {p.brand} {p.model} {p.tipo ? `(${p.tipo})` : ''}
                      </p>
                      <p className="text-muted-foreground truncate">
                        Responsável: {p.assigned_to || 'Não informado'}
                      </p>
                    </div>
                    <Button asChild variant="outline" size="sm" className="h-7 px-2.5 text-xs shrink-0">
                      <Link to="/terms">Conferir</Link>
                    </Button>
                  </div>
                ))}

                {pendentesDevolucao.map((p) => (
                  <div key={`devolucao-${p.id}`} className="py-2.5 flex items-start justify-between gap-3 text-xs">
                    <div className="space-y-0.5 min-w-0">
                      <div className="flex items-center gap-2">
                        <span className="font-mono font-semibold text-foreground">#{p.id}</span>
                        <span className="text-[11px] px-1.5 py-0.5 rounded-sm border border-border font-mono text-foreground font-medium">
                          Devolução Pendente
                        </span>
                      </div>
                      <p className="font-medium text-foreground truncate">
                        {p.brand} {p.model}
                      </p>
                      <p className="text-muted-foreground truncate">{p.revenda}</p>
                    </div>
                    <Button asChild variant="outline" size="sm" className="h-7 px-2.5 text-xs shrink-0">
                      <Link to="/return">Devolver</Link>
                    </Button>
                  </div>
                ))}
              </div>
            )}

            <div className="pt-2 border-t border-border flex justify-between items-center text-xs">
              <Link to="/terms" className="text-muted-foreground hover:text-foreground">
                Ver todos os termos →
              </Link>
              <Link to="/return" className="text-muted-foreground hover:text-foreground">
                Ver devoluções →
              </Link>
            </div>
          </div>
        )}
      </div>

      {/* Seção inferior: atividade recente e acesso rápido */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        {/* Linha do tempo compacta de eventos recentes (8 colunas) */}
        {podeVerOperacoes && (
          <div className={`lg:col-span-8 ${classePainel}`}>
            <div className="flex items-center justify-between border-b border-border pb-3">
              <div>
                <h2 className="text-body font-semibold text-foreground flex items-center gap-2">
                  <Clock className="size-4 text-muted-foreground" />
                  Últimas Movimentações no Sistema
                </h2>
                <p className="text-body-sm text-muted-foreground">
                  Auditoria cronológica das ações mais recentes
                </p>
              </div>
              <Link
                to="/history"
                className="text-xs text-foreground hover:underline flex items-center gap-1 font-medium"
              >
                Histórico completo <ExternalLink className="size-3" />
              </Link>
            </div>

            {historyQuery.isLoading ? (
              <div className="py-6 flex items-center justify-center gap-2 text-body-sm text-muted-foreground">
                <Loader2 className="h-4 w-4 animate-spin" />
                Carregando movimentações...
              </div>
            ) : historyQuery.isError ? (
              <div className="py-6 flex flex-col items-center justify-center gap-3 text-body-sm text-muted-foreground">
                <p>Não foi possível carregar as últimas movimentações.</p>
                <Button variant="outline" size="sm" onClick={() => historyQuery.refetch()}>
                  <RefreshCw size={14} />
                  Tentar novamente
                </Button>
              </div>
            ) : recentEvents.length === 0 ? (
              <p className="text-body-sm text-muted-foreground py-6 text-center">
                Nenhuma movimentação registrada recentemente.
              </p>
            ) : (
              <div className="divide-y divide-border">
                {recentEvents.map((evt) => (
                  <div key={evt.id} className="py-2.5 flex items-center justify-between gap-4 text-xs">
                    <div className="flex items-center gap-3 min-w-0">
                      <span className="font-mono text-muted-foreground text-[11px] shrink-0">
                        {formatDateTime(evt.data_operacao)}
                      </span>
                      <span className="text-[11px] font-mono px-1.5 py-0.5 rounded-sm border border-border text-foreground shrink-0 font-medium">
                        {evt.operation}
                      </span>
                      <span className="truncate text-foreground font-medium">
                        {evt.tipo ? `${evt.tipo} · ` : ''}
                        {evt.marca || evt.modelo
                          ? `${evt.marca ?? ''} ${evt.modelo ?? ''}`.trim()
                          : `Item #${evt.item_id ?? evt.id}`}
                      </span>
                    </div>

                    <div className="flex items-center gap-4 shrink-0 text-muted-foreground">
                      <span className="hidden sm:inline truncate max-w-[140px]">
                        Operador:{' '}
                        <strong className="text-foreground font-medium">{evt.operador}</strong>
                      </span>
                      {evt.usuario && (
                        <span className="hidden md:inline truncate max-w-[140px]">
                          Usuário: <span className="text-foreground">{evt.usuario}</span>
                        </span>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* Acesso operacional rápido (4 colunas; linha toda sem o histórico) */}
        <div className={`${podeVerOperacoes ? 'lg:col-span-4' : 'lg:col-span-12'} ${classePainel}`}>
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
