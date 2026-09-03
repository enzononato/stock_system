import { useMemo, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import {
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
} from "recharts";
import { ArrowDownLeft, ArrowUpRight, Calendar, Filter, PackagePlus } from "lucide-react";

import { getLoansChart, getRegistrationsChart, getMonthlyReport } from "@/api/reports";
import { listUnidades } from "@/api/unidades";
import { useConstants } from "@/hooks/useConstants";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { KpiCard } from "@/components/app/KpiCard";
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
  "Janeiro",
  "Fevereiro",
  "Março",
  "Abril",
  "Maio",
  "Junho",
  "Julho",
  "Agosto",
  "Setembro",
  "Outubro",
  "Novembro",
  "Dezembro",
];

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
          dia: `Dia ${day}`,
          Empréstimos: loans.data?.values[i] ?? 0,
          Devoluções: loans.data?.values2?.[i] ?? 0,
        }));
        const rData = (registrations.data?.days ?? []).map((day, i) => ({
          dia: `Dia ${day}`,
          Cadastros: registrations.data?.values[i] ?? 0,
        }));
        const tEmp = loans.data?.values?.reduce((a, v) => a + v, 0) ?? 0;
        const tDev = loans.data?.values2?.reduce((a, v) => a + v, 0) ?? 0;
        const tCad = registrations.data?.values?.reduce((a, v) => a + v, 0) ?? 0;
        return {
          loansChartData: lData,
          regChartData: rData,
          totalEmprestimos: tEmp,
          totalDevolucoes: tDev,
          totalCadastros: tCad,
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
        dia: `Dia ${d}`,
        Empréstimos: empByDay[d] ?? 0,
        Devoluções: devByDay[d] ?? 0,
      }));
      const rData = daysArray.map((d) => ({
        dia: `Dia ${d}`,
        Cadastros: cadByDay[d] ?? 0,
      }));
      const tEmp = Object.values(empByDay).reduce((a, v) => a + v, 0);
      const tDev = Object.values(devByDay).reduce((a, v) => a + v, 0);
      const tCad = Object.values(cadByDay).reduce((a, v) => a + v, 0);

      return {
        loansChartData: lData,
        regChartData: rData,
        totalEmprestimos: tEmp,
        totalDevolucoes: tDev,
        totalCadastros: tCad,
      };
    }, [isFiltered, loans.data, registrations.data, monthly.data, params.revenda, daysArray]);

  const monthLabel = `${MONTHS[params.month - 1] ?? "Mês"} de ${params.year}${params.revenda !== "all" ? ` · ${params.revenda}` : ""}`;
  const hasInvalidYear = !/^\d{4}$/.test(year) || Number(year) < 2000 || Number(year) > 2100;

  const applyPeriod = () => {
    if (hasInvalidYear) return;
    setParams({ year: Number(year), month: Number(month), revenda: filterRevenda });
  };

  const tooltipStyle = {
    backgroundColor: "var(--color-popover)",
    borderColor: "var(--color-border)",
    borderRadius: 10,
    fontSize: 12,
    fontWeight: 600,
  };
  const isLoading = isFiltered ? monthly.isLoading : loans.isLoading || registrations.isLoading;
  const error = isFiltered ? monthly.error : loans.error || registrations.error;

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Gestão"
        title="Indicadores"
        description="Acompanhe a movimentação do patrimônio por período e unidade usando os dados disponíveis no sistema."
      />

      <div className="surface-panel p-4 sm:p-5">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-end">
          <div className="flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-primary lg:mr-2 lg:pb-2">
            <Filter className="size-4" aria-hidden />
            Filtros
          </div>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="charts-year">Ano</Label>
              <Input
                id="charts-year"
                inputMode="numeric"
                maxLength={4}
                value={year}
                onChange={(e) => setYear(e.target.value.replace(/\D/g, "").slice(0, 4))}
                className="w-full sm:w-28"
                aria-invalid={hasInvalidYear}
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Mês</Label>
              <Select value={month} onValueChange={setMonth}>
                <SelectTrigger className="w-full sm:w-40">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {MONTHS.map((m, i) => (
                    <SelectItem key={m} value={String(i + 1)}>
                      {m}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Unidade / Revenda</Label>
              <Select value={filterRevenda} onValueChange={setFilterRevenda}>
                <SelectTrigger className="w-full sm:w-48">
                  <SelectValue placeholder="Todas as unidades" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">Todas as unidades</SelectItem>
                  {revendaOptions.map((r) => (
                    <SelectItem key={r} value={r}>
                      {r}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
          </div>
          <Button onClick={applyPeriod} disabled={hasInvalidYear} className="w-full sm:w-auto">
            <Calendar className="mr-1.5 size-4" aria-hidden />
            Aplicar
          </Button>
          {hasInvalidYear && (
            <p className="text-xs text-destructive lg:pb-2">Informe um ano entre 2000 e 2100.</p>
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
        <KpiCard label="Empréstimos no mês" value={totalEmprestimos} icon={ArrowUpRight} />
        <KpiCard label="Devoluções no mês" value={totalDevolucoes} icon={ArrowDownLeft} />
        <KpiCard label="Cadastros no mês" value={totalCadastros} icon={PackagePlus} />
      </div>

      <Section
        title="Empréstimos × Devoluções"
        description={`Movimentação diária em ${monthLabel}`}
      >
        {isLoading ? (
          <LoadingState label="Carregando indicadores..." />
        ) : error ? (
          <ErrorState
            error={error}
            onRetry={() => (isFiltered ? monthly.refetch() : loans.refetch())}
          />
        ) : loansChartData.every((d) => !d.Empréstimos && !d.Devoluções) ? (
          <EmptyState
            title="Nenhuma movimentação registrada"
            description={`Não há empréstimos ou devoluções em ${monthLabel}.`}
            icon={<PackagePlus className="size-5" aria-hidden />}
          />
        ) : (
          <div className="min-w-0 overflow-x-auto">
            <div className="min-w-[560px]">
              <ResponsiveContainer width="100%" height={320}>
                <AreaChart
                  data={loansChartData}
                  margin={{ top: 10, right: 10, left: -20, bottom: 0 }}
                >
                  <defs>
                    <linearGradient id="gradEmprestimos" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="var(--color-chart-1)" stopOpacity={0.3} />
                      <stop offset="95%" stopColor="var(--color-chart-1)" stopOpacity={0.02} />
                    </linearGradient>
                    <linearGradient id="gradDevolucoes" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="var(--color-chart-2)" stopOpacity={0.3} />
                      <stop offset="95%" stopColor="var(--color-chart-2)" stopOpacity={0.02} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" className="stroke-border" vertical={false} />
                  <XAxis dataKey="dia" tick={{ fontSize: 11 }} axisLine={false} tickLine={false} />
                  <YAxis
                    allowDecimals={false}
                    tick={{ fontSize: 11 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip contentStyle={tooltipStyle} />
                  <Legend wrapperStyle={{ paddingTop: 10, fontSize: 12, fontWeight: 600 }} />
                  <Area
                    type="monotone"
                    dataKey="Empréstimos"
                    stroke="var(--color-chart-1)"
                    strokeWidth={2.5}
                    fill="url(#gradEmprestimos)"
                    dot={{ r: 3 }}
                    activeDot={{ r: 5 }}
                  />
                  <Area
                    type="monotone"
                    dataKey="Devoluções"
                    stroke="var(--color-chart-2)"
                    strokeWidth={2.5}
                    fill="url(#gradDevolucoes)"
                    dot={{ r: 3 }}
                    activeDot={{ r: 5 }}
                  />
                </AreaChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}
      </Section>

      <Section
        title="Novos cadastros"
        description={`Volume diário de novos patrimônios em ${monthLabel}`}
      >
        {isLoading ? (
          <LoadingState label="Carregando indicadores..." />
        ) : error ? (
          <ErrorState
            error={error}
            onRetry={() => (isFiltered ? monthly.refetch() : registrations.refetch())}
          />
        ) : regChartData.every((d) => !d.Cadastros) ? (
          <EmptyState
            title="Nenhum cadastro registrado"
            description={`Não há novos cadastros em ${monthLabel}.`}
            icon={<PackagePlus className="size-5" aria-hidden />}
          />
        ) : (
          <div className="min-w-0 overflow-x-auto">
            <div className="min-w-[560px]">
              <ResponsiveContainer width="100%" height={280}>
                <AreaChart
                  data={regChartData}
                  margin={{ top: 10, right: 10, left: -20, bottom: 0 }}
                >
                  <defs>
                    <linearGradient id="gradCadastros" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="var(--color-chart-4)" stopOpacity={0.3} />
                      <stop offset="95%" stopColor="var(--color-chart-4)" stopOpacity={0.02} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" className="stroke-border" vertical={false} />
                  <XAxis dataKey="dia" tick={{ fontSize: 11 }} axisLine={false} tickLine={false} />
                  <YAxis
                    allowDecimals={false}
                    tick={{ fontSize: 11 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip contentStyle={tooltipStyle} />
                  <Area
                    type="monotone"
                    dataKey="Cadastros"
                    stroke="var(--color-chart-4)"
                    strokeWidth={2.5}
                    fill="url(#gradCadastros)"
                    dot={{ r: 3 }}
                    activeDot={{ r: 5 }}
                  />
                </AreaChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}
      </Section>
    </div>
  );
}
