import { useMemo, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import {
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  BarChart,
  Bar,
  LineChart,
  Line,
} from "recharts";
import { Calendar, Filter } from "lucide-react";

import { getLoansChart, getRegistrationsChart, getMonthlyReport } from "@/api/reports";
import { listUnidades } from "@/api/unidades";
import { useConstants } from "@/hooks/useConstants";
import { LoadingState, EmptyState, ErrorState } from "@/components/app/StateBlocks";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

const MONTHS = [
  "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
];

const TOOLTIP_STYLE = {
  backgroundColor: "var(--color-surface)",
  borderColor: "var(--color-border)",
  borderRadius: 4,
  fontSize: 12,
  fontWeight: 600,
  color: "var(--color-foreground)",
};

export function ChartsPage() {
  const now = new Date();
  const [year, setYear] = useState(String(now.getFullYear()));
  const [month, setMonth] = useState(String(now.getMonth() + 1));
  const [filterRevenda, setFilterRevenda] = useState("all");
  const [params, setParams] = useState({
    year: now.getFullYear(),
    month: now.getMonth() + 1,
    revenda: "all",
  });

  const { revendas = [] } = useConstants();
  const { data: unidades = [] } = useQuery({
    queryKey: ["unidades-charts"],
    queryFn: () => listUnidades(),
  });

  const revendaOptions = useMemo(() => {
    const set = new Set<string>();
    revendas.forEach((r) => r && set.add(r));
    unidades.forEach((u) => u.nome && set.add(u.nome));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "pt-BR"));
  }, [revendas, unidades]);

  const loans = useQuery({
    queryKey: ["chart-loans", params.year, params.month],
    queryFn: () => getLoansChart(params.year, params.month),
  });
  const registrations = useQuery({
    queryKey: ["chart-registrations", params.year, params.month],
    queryFn: () => getRegistrationsChart(params.year, params.month),
  });
  const monthly = useQuery({
    queryKey: ["chart-monthly-report", params.year, params.month],
    queryFn: () => getMonthlyReport(params.year, params.month),
    enabled: params.revenda !== "all",
  });

  const isFiltered = params.revenda !== "all";
  const daysInMonth = new Date(params.year, params.month, 0).getDate();
  const daysArray = Array.from({ length: daysInMonth }, (_, i) => i + 1);

  const { loansChartData, regChartData, totalEmprestimos, totalDevolucoes, totalCadastros } =
    useMemo(() => {
      if (!isFiltered) {
        const lData = (loans.data?.days ?? []).map((day, i) => ({
          dia: `${day}`,
          Empréstimos: loans.data?.values[i] ?? 0,
          Devoluções: loans.data?.values2?.[i] ?? 0,
        }));
        const rData = (registrations.data?.days ?? []).map((day, i) => ({
          dia: `${day}`,
          Cadastros: registrations.data?.values[i] ?? 0,
        }));
        return {
          loansChartData: lData,
          regChartData: rData,
          totalEmprestimos: loans.data?.values?.reduce((a, v) => a + v, 0) ?? 0,
          totalDevolucoes: loans.data?.values2?.reduce((a, v) => a + v, 0) ?? 0,
          totalCadastros: registrations.data?.values?.reduce((a, v) => a + v, 0) ?? 0,
        };
      }

      const rows = (monthly.data ?? []).filter((r) => r.revenda === params.revenda);
      const empByDay: Record<number, number> = {};
      const devByDay: Record<number, number> = {};
      const cadByDay: Record<number, number> = {};

      rows.forEach((r) => {
        if (r.data_emprestimo) {
          const d = new Date(r.data_emprestimo).getDate();
          if (d) empByDay[d] = (empByDay[d] ?? 0) + 1;
        }
        if (r.data_devolucao) {
          const d = new Date(r.data_devolucao).getDate();
          if (d) devByDay[d] = (devByDay[d] ?? 0) + 1;
        }
        if (r.operation_type === "Cadastro") {
          const dateStr = r.data_confirmacao ?? r.data_emprestimo;
          const d = dateStr ? new Date(dateStr).getDate() : null;
          if (d) cadByDay[d] = (cadByDay[d] ?? 0) + 1;
        }
      });

      const lData = daysArray.map((d) => ({
        dia: `${d}`,
        Empréstimos: empByDay[d] ?? 0,
        Devoluções: devByDay[d] ?? 0,
      }));
      const rData = daysArray.map((d) => ({
        dia: `${d}`,
        Cadastros: cadByDay[d] ?? 0,
      }));

      return {
        loansChartData: lData,
        regChartData: rData,
        totalEmprestimos: Object.values(empByDay).reduce((a, v) => a + v, 0),
        totalDevolucoes: Object.values(devByDay).reduce((a, v) => a + v, 0),
        totalCadastros: Object.values(cadByDay).reduce((a, v) => a + v, 0),
      };
    }, [isFiltered, loans.data, registrations.data, monthly.data, params.revenda, daysArray]);

  const hasInvalidYear = !/^\d{4}$/.test(year) || Number(year) < 2000 || Number(year) > 2100;
  const monthLabel = `${MONTHS[params.month - 1] ?? "Mês"} / ${params.year}${params.revenda !== "all" ? ` · ${params.revenda}` : ""}`;
  const isLoading = isFiltered ? monthly.isLoading : loans.isLoading || registrations.isLoading;
  const hasError = isFiltered ? monthly.error : loans.error || registrations.error;

  function applyPeriod() {
    if (hasInvalidYear) return;
    setParams({ year: Number(year), month: Number(month), revenda: filterRevenda });
  }

  // Distribuição por dia da semana (a partir de loans data)
  const weekdayDist = useMemo(() => {
    const days = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
    const counts: number[] = Array(7).fill(0);
    loansChartData.forEach((d, idx) => {
      const date = new Date(params.year, params.month - 1, idx + 1);
      counts[date.getDay()] = (counts[date.getDay()] ?? 0) + (d.Empréstimos ?? 0);
    });
    return days.map((d, i) => ({ dia: d, Empréstimos: counts[i] ?? 0 }));
  }, [loansChartData, params.year, params.month]);

  return (
    <div className="page-container-dense space-y-6">
      {/* Header */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Inteligência de Operações
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Análise Comparativa</span>
          </div>
          <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">
            Indicadores Operacionais
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Movimentação patrimonial comparativa por período, filial e tipo de operação.
          </p>
        </div>
      </div>

      {/* Toolbar de Filtro */}
      <div className="rounded-[6px] border border-border bg-surface p-4">
        <div className="flex flex-wrap items-end gap-4">
          <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground self-end pb-0.5">
            Período
          </span>

          <div className="space-y-1">
            <Label className="text-caption font-medium">Ano</Label>
            <Input
              inputMode="numeric"
              maxLength={4}
              value={year}
              onChange={(e) => setYear(e.target.value.replace(/\D/g, "").slice(0, 4))}
              className="w-20 h-8 rounded-[4px] border-border bg-background text-body-sm font-mono"
              aria-invalid={hasInvalidYear}
            />
          </div>

          <div className="space-y-1">
            <Label className="text-caption font-medium">Mês</Label>
            <Select value={month} onValueChange={setMonth}>
              <SelectTrigger className="w-36 h-8 rounded-[4px] border-border bg-background text-body-sm">
                <SelectValue />
              </SelectTrigger>
              <SelectContent className="rounded-[4px] border-border bg-surface">
                {MONTHS.map((m, i) => (
                  <SelectItem key={m} value={String(i + 1)} className="text-body-sm rounded-[2px]">{m}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          <div className="space-y-1">
            <Label className="text-caption font-medium">Unidade / Filial</Label>
            <Select value={filterRevenda} onValueChange={setFilterRevenda}>
              <SelectTrigger className="w-44 h-8 rounded-[4px] border-border bg-background text-body-sm">
                <SelectValue placeholder="Todas" />
              </SelectTrigger>
              <SelectContent className="rounded-[4px] border-border bg-surface">
                <SelectItem value="all" className="text-body-sm rounded-[2px]">Todas as unidades</SelectItem>
                {revendaOptions.map((r) => (
                  <SelectItem key={r} value={r} className="text-body-sm rounded-[2px]">{r}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          <Button
            size="sm"
            onClick={applyPeriod}
            disabled={hasInvalidYear}
            className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 font-medium h-8"
          >
            <Calendar className="mr-1.5 size-3.5" />
            Aplicar
          </Button>

          {hasInvalidYear && (
            <p className="text-caption text-muted-foreground self-end pb-1">
              Ano inválido (2000–2100)
            </p>
          )}
        </div>
      </div>

      {/* KPIs Numéricos */}
      <div className="grid grid-cols-3 gap-4">
        {[
          { label: "Empréstimos no mês", value: totalEmprestimos, symbol: "↑" },
          { label: "Devoluções no mês", value: totalDevolucoes, symbol: "↓" },
          { label: "Cadastros no mês", value: totalCadastros, symbol: "+" },
        ].map((stat) => (
          <div key={stat.label} className="rounded-[6px] border border-border bg-surface p-5">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
              {stat.label}
            </span>
            <div className="flex items-baseline gap-2 mt-2">
              <span className="font-mono text-heading-lg font-bold text-foreground">{stat.value}</span>
              <span className="font-mono text-body-lg text-muted-foreground">{stat.symbol}</span>
            </div>
            <span className="text-caption text-muted-foreground">{monthLabel}</span>
          </div>
        ))}
      </div>

      {/* Gráfico 1: Área de Empréstimos × Devoluções */}
      <div className="rounded-[6px] border border-border bg-surface p-5">
        <div className="border-b border-border pb-3 mb-5">
          <h2 className="text-body-lg font-semibold text-foreground">
            Empréstimos × Devoluções — {monthLabel}
          </h2>
          <p className="text-caption text-muted-foreground">
            Comparativo diário de entradas e saídas do inventário ativo.
          </p>
        </div>

        {isLoading ? (
          <LoadingState label="Carregando indicadores…" />
        ) : hasError ? (
          <ErrorState
            error={hasError}
            onRetry={() => isFiltered ? void monthly.refetch() : void loans.refetch()}
          />
        ) : loansChartData.every((d) => !d.Empréstimos && !d.Devoluções) ? (
          <div className="py-8 text-center text-caption text-muted-foreground">
            Nenhuma movimentação neste período.
          </div>
        ) : (
          <div className="min-w-0 overflow-x-auto">
            <div className="min-w-[520px]">
              <ResponsiveContainer width="100%" height={300}>
                <AreaChart data={loansChartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                  <defs>
                    <linearGradient id="gradEmp" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="var(--color-foreground)" stopOpacity={0.15} />
                      <stop offset="95%" stopColor="var(--color-foreground)" stopOpacity={0.01} />
                    </linearGradient>
                    <linearGradient id="gradDev" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="var(--color-muted-foreground)" stopOpacity={0.15} />
                      <stop offset="95%" stopColor="var(--color-muted-foreground)" stopOpacity={0.01} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" vertical={false} />
                  <XAxis dataKey="dia" tick={{ fontSize: 10, fill: "var(--color-muted-foreground)" }} axisLine={false} tickLine={false} />
                  <YAxis allowDecimals={false} tick={{ fontSize: 10, fill: "var(--color-muted-foreground)" }} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={TOOLTIP_STYLE} />
                  <Area
                    type="monotone"
                    dataKey="Empréstimos"
                    stroke="var(--color-foreground)"
                    strokeWidth={2}
                    fill="url(#gradEmp)"
                    dot={{ r: 2, fill: "var(--color-foreground)" }}
                    activeDot={{ r: 4 }}
                  />
                  <Area
                    type="monotone"
                    dataKey="Devoluções"
                    stroke="var(--color-muted-foreground)"
                    strokeWidth={2}
                    fill="url(#gradDev)"
                    dot={{ r: 2, fill: "var(--color-muted-foreground)" }}
                    activeDot={{ r: 4 }}
                  />
                </AreaChart>
              </ResponsiveContainer>

              {/* Legenda manual monocromática */}
              <div className="flex items-center gap-6 mt-3 pt-3 border-t border-border text-caption">
                <div className="flex items-center gap-1.5">
                  <div className="w-6 h-0.5 bg-foreground" />
                  <span className="text-foreground font-medium">Empréstimos</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <div className="w-6 h-0.5 bg-muted-foreground" />
                  <span className="text-muted-foreground">Devoluções</span>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Grid: Gráfico Barras (Cadastros por dia) + Distribuição Semanal */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Gráfico 2: Barras de Cadastros */}
        <div className="rounded-[6px] border border-border bg-surface p-5">
          <div className="border-b border-border pb-3 mb-5">
            <h2 className="text-body font-semibold text-foreground">Cadastros por Dia</h2>
            <p className="text-caption text-muted-foreground">
              Novos patrimônios registrados no período.
            </p>
          </div>

          {isLoading ? (
            <div className="py-6 text-center text-caption text-muted-foreground">Carregando…</div>
          ) : regChartData.every((d) => !d.Cadastros) ? (
            <div className="py-8 text-center text-caption text-muted-foreground">
              Nenhum cadastro neste período.
            </div>
          ) : (
            <div className="overflow-x-auto">
              <div className="min-w-[300px]">
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={regChartData} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" vertical={false} />
                    <XAxis dataKey="dia" tick={{ fontSize: 9, fill: "var(--color-muted-foreground)" }} axisLine={false} tickLine={false} />
                    <YAxis allowDecimals={false} tick={{ fontSize: 9, fill: "var(--color-muted-foreground)" }} axisLine={false} tickLine={false} />
                    <Tooltip contentStyle={TOOLTIP_STYLE} />
                    <Bar dataKey="Cadastros" fill="var(--color-foreground)" radius={[2, 2, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}
        </div>

        {/* Gráfico 3: Distribuição por dia da semana */}
        <div className="rounded-[6px] border border-border bg-surface p-5">
          <div className="border-b border-border pb-3 mb-5">
            <h2 className="text-body font-semibold text-foreground">Distribuição Semanal de Empréstimos</h2>
            <p className="text-caption text-muted-foreground">
              Concentração de movimentações por dia da semana.
            </p>
          </div>

          {weekdayDist.every((d) => !d.Empréstimos) ? (
            <div className="py-8 text-center text-caption text-muted-foreground">
              Nenhum dado disponível para este período.
            </div>
          ) : (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={weekdayDist} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" vertical={false} />
                <XAxis dataKey="dia" tick={{ fontSize: 11, fill: "var(--color-muted-foreground)" }} axisLine={false} tickLine={false} />
                <YAxis allowDecimals={false} tick={{ fontSize: 9, fill: "var(--color-muted-foreground)" }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={TOOLTIP_STYLE} />
                <Bar dataKey="Empréstimos" fill="var(--color-foreground)" opacity={0.8} radius={[2, 2, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>
    </div>
  );
}
