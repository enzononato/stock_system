import { useMemo } from "react";
import { useQuery } from "@tanstack/react-query";
import { Link } from "@tanstack/react-router";
import {
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
} from "recharts";
import {
  AlertCircle,
  ArrowRight,
  Boxes,
  Clock,
  ExternalLink,
  FileSignature,
  PackageCheck,
  TrendingUp,
  Undo2,
} from "lucide-react";

import { listItemsPaginated } from "@/api/items";
import { getLoansChart, getRegistrationsChart } from "@/api/reports";
import { listHistoryPaginated } from "@/api/history";
import { StatBlock } from "@/components/data/stat-block";
import { Button } from "@/components/ui/button";
import { TOKENS } from "@/lib/tokens";
import { useTheme } from "@/lib/theme";
import { formatDateTime } from "@/lib/utils";

export function DashboardPage() {
  const { theme } = useTheme();
  const tokens = theme === "dark" ? TOKENS.dark : TOKENS.light;

  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;

  // Consultas reais de dados
  const itemsQuery = useQuery({
    queryKey: ["dashboard-items"],
    queryFn: () => listItemsPaginated({ limit: 500 }),
  });

  const loansChart = useQuery({
    queryKey: ["dashboard-chart-loans", currentYear, currentMonth],
    queryFn: () => getLoansChart(currentYear, currentMonth),
  });

  const historyQuery = useQuery({
    queryKey: ["dashboard-recent-history"],
    queryFn: () => listHistoryPaginated({ limit: 6, offset: 0 }),
  });

  const items = itemsQuery.data?.items ?? [];
  const totalPatrimonios = itemsQuery.data?.total ?? items.length;

  const disponiveis = items.filter((i) => i.status === "Disponível").length;
  const emUso = items.filter((i) => i.status === "Indisponível" && Boolean(i.assigned_to)).length;
  const pendentesConfirmacao = items.filter((i) => i.status === "Pendente");
  const pendentesDevolucao = items.filter((i) => i.status === "Pendente Devolução");

  const totalAtencao = pendentesConfirmacao.length + pendentesDevolucao.length;

  // Dados do gráfico de tendência monocromático
  const chartData = useMemo(() => {
    const days = loansChart.data?.days ?? [];
    const values = loansChart.data?.values ?? [];
    const values2 = loansChart.data?.values2 ?? [];

    if (days.length === 0) {
      // Fallback sutil de dias se vazio
      return Array.from({ length: 15 }, (_, i) => ({
        dia: `D${i + 1}`,
        Empréstimos: 0,
        Devoluções: 0,
      }));
    }

    return days.slice(-15).map((d, i) => ({
      dia: `D${d}`,
      Empréstimos: values[i] ?? 0,
      Devoluções: values2[i] ?? 0,
    }));
  }, [loansChart.data]);

  const recentEvents = historyQuery.data?.items ?? [];

  return (
    <div className="space-y-8">
      {/* Cabeçalho Editorial com Contexto e Horário */}
      <div className="border-b border-border pb-5">
        <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-2">
          <div>
            <span className="text-caption text-muted-foreground">Sistema de Gestão Patrimonial</span>
            <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
              Panorama Operacional
            </h1>
          </div>
          <div className="flex items-center gap-3 text-xs text-muted-foreground font-mono num">
            <span>Mês Atual: {String(currentMonth).padStart(2, "0")}/{currentYear}</span>
            <span>·</span>
            <span className="flex items-center gap-1.5">
              <span className="size-1.5 rounded-full bg-foreground" />
              Sincronizado
            </span>
          </div>
        </div>

        {/* Faixa Superior de Stat Blocks Tipográficos (Sem Cards de KPI Repetitivos - Seção 23) */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-6 pt-6 mt-4 border-t border-border/60">
          <StatBlock
            label="Total de Patrimônios"
            value={itemsQuery.isLoading ? "—" : totalPatrimonios}
            hint="Ativos cadastrados"
          />
          <StatBlock
            label="Disponíveis para Alocação"
            value={itemsQuery.isLoading ? "—" : disponiveis}
            hint={`${totalPatrimonios > 0 ? Math.round((disponiveis / totalPatrimonios) * 100) : 0}% da frota`}
          />
          <StatBlock
            label="Empréstimos Ativos"
            value={itemsQuery.isLoading ? "—" : emUso}
            hint="Em uso com colaboradores"
          />
          <StatBlock
            label="Pendências Operacionais"
            value={itemsQuery.isLoading ? "—" : totalAtencao}
            hint={totalAtencao > 0 ? "! Atenção requerida" : "✓ Nenhuma pendência"}
          />
        </div>
      </div>

      {/* BLOCO ASSIMÉTRICO PRINCIPAL (Seção 23: Tendência + Atenção Operacional) */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        {/* Lado Esquerdo: Tendência de Movimentações (7 colunas) */}
        <div className="lg:col-span-7 bg-surface border border-border rounded-md p-5 space-y-4">
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
            <div className="flex items-center gap-3 text-xs font-mono num">
              <span className="flex items-center gap-1">
                <span className="size-2 rounded-full bg-foreground" /> Empréstimos
              </span>
              <span className="flex items-center gap-1 text-muted-foreground">
                <span className="size-2 rounded-full bg-muted-foreground border border-border" /> Devoluções
              </span>
            </div>
          </div>

          <div className="h-64 w-full pt-2">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={chartData} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
                <CartesianGrid strokeDasharray="2 2" stroke={tokens.border} vertical={false} />
                <XAxis dataKey="dia" tick={{ fontSize: 11, fill: tokens.textSecondary }} axisLine={false} tickLine={false} />
                <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: tokens.textSecondary }} axisLine={false} tickLine={false} />
                <Tooltip
                  contentStyle={{
                    backgroundColor: tokens.surface,
                    borderColor: tokens.borderStrong,
                    borderRadius: 4,
                    fontSize: 12,
                    color: tokens.textPrimary,
                  }}
                />
                <Area
                  type="monotone"
                  dataKey="Empréstimos"
                  stroke={tokens.textPrimary}
                  strokeWidth={2}
                  fill={tokens.surfaceAlt}
                  fillOpacity={0.6}
                />
                <Area
                  type="monotone"
                  dataKey="Devoluções"
                  stroke={tokens.textSecondary}
                  strokeWidth={1.5}
                  strokeDasharray="3 3"
                  fill="transparent"
                />
              </AreaChart>
            </ResponsiveContainer>
          </div>

          <div className="pt-3 border-t border-border flex items-center justify-between text-xs text-muted-foreground">
            <span>Valores consolidados em tempo real pelo banco de dados</span>
            <Link to="/charts" className="text-foreground hover:underline inline-flex items-center gap-1">
              Ver indicadores detalhados <ArrowRight className="size-3" />
            </Link>
          </div>
        </div>

        {/* Lado Direito: Atenção Operacional com Divisores Hairline (5 colunas) */}
        <div className="lg:col-span-5 bg-surface border border-border rounded-md p-5 space-y-4">
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
            <span className="text-caption font-mono num px-2 py-0.5 rounded-sm border border-border-strong text-foreground font-bold">
              {totalAtencao}
            </span>
          </div>

          {totalAtencao === 0 ? (
            <div className="py-8 text-center text-body-sm text-muted-foreground space-y-1">
              <p className="font-medium text-foreground">Operação em dia</p>
              <p>Não há termos ou devoluções pendentes no momento.</p>
            </div>
          ) : (
            <div className="divide-y divide-border/60">
              {pendentesConfirmacao.slice(0, 4).map((p) => (
                <div key={p.id} className="py-2.5 flex items-start justify-between gap-3 text-xs">
                  <div className="space-y-0.5 min-w-0">
                    <div className="flex items-center gap-2">
                      <span className="font-mono font-semibold text-foreground">#{p.id}</span>
                      <span className="text-caption px-1 rounded-sm border border-border text-foreground">
                        Termo Pendente
                      </span>
                    </div>
                    <p className="font-medium text-foreground truncate">
                      {p.brand} {p.model} ({p.tipo})
                    </p>
                    <p className="text-muted-foreground truncate">
                      Responsável: {p.assigned_to || "Não informado"}
                    </p>
                  </div>
                  <Link to="/terms">
                    <Button variant="outline" size="sm" className="h-7 px-2 text-xs shrink-0">
                      Conferir
                    </Button>
                  </Link>
                </div>
              ))}

              {pendentesDevolucao.slice(0, 3).map((p) => (
                <div key={p.id} className="py-2.5 flex items-start justify-between gap-3 text-xs">
                  <div className="space-y-0.5 min-w-0">
                    <div className="flex items-center gap-2">
                      <span className="font-mono font-semibold text-foreground">#{p.id}</span>
                      <span className="text-caption px-1 rounded-sm border border-border text-foreground">
                        Devolução Pendente
                      </span>
                    </div>
                    <p className="font-medium text-foreground truncate">
                      {p.brand} {p.model}
                    </p>
                    <p className="text-muted-foreground truncate">{p.revenda}</p>
                  </div>
                  <Link to="/return">
                    <Button variant="outline" size="sm" className="h-7 px-2 text-xs shrink-0">
                      Devolver
                    </Button>
                  </Link>
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
      </div>

      {/* SEÇÃO INFERIOR: Atividade Recente e Acesso Rápido */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        {/* Linha do Tempo Compacta de Eventos Recentes (8 colunas) */}
        <div className="lg:col-span-8 bg-surface border border-border rounded-md p-5 space-y-4">
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
            <Link to="/history" className="text-xs text-foreground hover:underline flex items-center gap-1 font-medium">
              Histórico completo <ExternalLink className="size-3" />
            </Link>
          </div>

          {recentEvents.length === 0 ? (
            <p className="text-body-sm text-muted-foreground py-6 text-center">
              Nenhuma movimentação registrada recentemente.
            </p>
          ) : (
            <div className="divide-y divide-border/60">
              {recentEvents.map((evt) => (
                <div key={evt.id} className="py-2.5 flex items-center justify-between gap-4 text-xs">
                  <div className="flex items-center gap-3 min-w-0">
                    <span className="font-mono text-muted-foreground text-[11px] shrink-0">
                      {formatDateTime(evt.data_operacao)}
                    </span>
                    <span className="text-caption px-1.5 py-0.5 rounded-sm border border-border text-foreground shrink-0 font-medium">
                      {evt.operation}
                    </span>
                    <span className="truncate text-foreground font-medium">
                      {evt.tipo ? `${evt.tipo} · ` : ""}
                      {evt.marca || evt.modelo ? `${evt.marca ?? ""} ${evt.modelo ?? ""}`.trim() : `Item #${evt.item_id || evt.id}`}
                    </span>
                  </div>

                  <div className="flex items-center gap-4 shrink-0 text-muted-foreground">
                    <span className="hidden sm:inline truncate max-w-[140px]">
                      Operador: <strong className="text-foreground font-medium">{evt.operador}</strong>
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

        {/* Acesso Operacional Rápido (4 colunas) */}
        <div className="lg:col-span-4 bg-surface border border-border rounded-md p-5 space-y-4">
          <div className="border-b border-border pb-3">
            <h2 className="text-body font-semibold text-foreground">Fluxos Rápidos de TI</h2>
            <p className="text-body-sm text-muted-foreground">
              Acesso direto às operações rotineiras
            </p>
          </div>

          <div className="space-y-2">
            <Link to="/stock" className="flex items-center justify-between p-3 rounded border border-border hover:bg-surface-alt transition-colors group">
              <div className="flex items-center gap-2.5">
                <Boxes className="size-4 text-muted-foreground group-hover:text-foreground" />
                <div>
                  <p className="text-body-sm font-medium text-foreground">Consultar Estoque</p>
                  <p className="text-[11px] text-muted-foreground">Tabela densa de patrimônios</p>
                </div>
              </div>
              <ArrowRight className="size-3.5 text-muted-foreground group-hover:text-foreground" />
            </Link>

            <Link to="/loan" className="flex items-center justify-between p-3 rounded border border-border hover:bg-surface-alt transition-colors group">
              <div className="flex items-center gap-2.5">
                <PackageCheck className="size-4 text-muted-foreground group-hover:text-foreground" />
                <div>
                  <p className="text-body-sm font-medium text-foreground">Novo Empréstimo</p>
                  <p className="text-[11px] text-muted-foreground">Atribuir ativo e gerar termo</p>
                </div>
              </div>
              <ArrowRight className="size-3.5 text-muted-foreground group-hover:text-foreground" />
            </Link>

            <Link to="/return" className="flex items-center justify-between p-3 rounded border border-border hover:bg-surface-alt transition-colors group">
              <div className="flex items-center gap-2.5">
                <Undo2 className="size-4 text-muted-foreground group-hover:text-foreground" />
                <div>
                  <p className="text-body-sm font-medium text-foreground">Devolução Física</p>
                  <p className="text-[11px] text-muted-foreground">Conferir periféricos e estado</p>
                </div>
              </div>
              <ArrowRight className="size-3.5 text-muted-foreground group-hover:text-foreground" />
            </Link>

            <Link to="/terms" className="flex items-center justify-between p-3 rounded border border-border hover:bg-surface-alt transition-colors group">
              <div className="flex items-center gap-2.5">
                <FileSignature className="size-4 text-muted-foreground group-hover:text-foreground" />
                <div>
                  <p className="text-body-sm font-medium text-foreground">Termos de Responsabilidade</p>
                  <p className="text-[11px] text-muted-foreground">Assinaturas e PDFs arquivados</p>
                </div>
              </div>
              <ArrowRight className="size-3.5 text-muted-foreground group-hover:text-foreground" />
            </Link>
          </div>
        </div>
      </div>
    </div>
  );
}
