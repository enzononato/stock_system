import { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";
import { ArrowDownLeft, ArrowUpRight, Calendar, Filter, PackagePlus } from "lucide-react";

import { getLoansChart, getRegistrationsChart } from "@/api/reports";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { KpiCard } from "@/components/app/KpiCard";
import { LoadingState, EmptyState, ErrorState } from "@/components/app/StateBlocks";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

const MONTHS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

export function ChartsPage() {
  const now = new Date();
  const [year, setYear] = useState(String(now.getFullYear()));
  const [month, setMonth] = useState(String(now.getMonth() + 1));
  const [params, setParams] = useState({ year: now.getFullYear(), month: now.getMonth() + 1 });

  const loans = useQuery({ queryKey: ["chart-loans", params.year, params.month], queryFn: () => getLoansChart(params.year, params.month) });
  const registrations = useQuery({ queryKey: ["chart-registrations", params.year, params.month], queryFn: () => getRegistrationsChart(params.year, params.month) });

  const loansChartData = (loans.data?.days ?? []).map((day, i) => ({ dia: `Dia ${day}`, Empréstimos: loans.data?.values[i] ?? 0, Devoluções: loans.data?.values2?.[i] ?? 0 }));
  const regChartData = (registrations.data?.days ?? []).map((day, i) => ({ dia: `Dia ${day}`, Cadastros: registrations.data?.values[i] ?? 0 }));
  const totalEmprestimos = loans.data?.values?.reduce((a, v) => a + v, 0) ?? 0;
  const totalDevolucoes = loans.data?.values2?.reduce((a, v) => a + v, 0) ?? 0;
  const totalCadastros = registrations.data?.values?.reduce((a, v) => a + v, 0) ?? 0;
  const monthLabel = `${MONTHS[params.month - 1] ?? "Mês"} de ${params.year}`;
  const hasInvalidYear = !/^\d{4}$/.test(year) || Number(year) < 2000 || Number(year) > 2100;

  const applyPeriod = () => {
    if (hasInvalidYear) return;
    setParams({ year: Number(year), month: Number(month) });
  };

  const tooltipStyle = { backgroundColor: "var(--color-popover)", borderColor: "var(--color-border)", borderRadius: 10, fontSize: 12, fontWeight: 600 };

  return (
    <div className="space-y-6">
      <PageHeader eyebrow="Gestão" title="Indicadores" description="Acompanhe a movimentação do patrimônio por período usando os dados disponíveis no sistema." />

      <div className="surface-panel p-4 sm:p-5">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-end">
          <div className="flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-primary lg:mr-2 lg:pb-2"><Filter className="size-4" aria-hidden />Período</div>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div className="flex flex-col gap-1.5"><Label htmlFor="charts-year">Ano</Label><Input id="charts-year" inputMode="numeric" maxLength={4} value={year} onChange={(e) => setYear(e.target.value.replace(/\D/g, "").slice(0, 4))} className="w-full sm:w-28" aria-invalid={hasInvalidYear} /></div>
            <div className="flex flex-col gap-1.5"><Label>Mês</Label><Select value={month} onValueChange={setMonth}><SelectTrigger className="w-full sm:w-44"><SelectValue /></SelectTrigger><SelectContent>{MONTHS.map((m, i) => <SelectItem key={m} value={String(i + 1)}>{m}</SelectItem>)}</SelectContent></Select></div>
          </div>
          <Button onClick={applyPeriod} disabled={hasInvalidYear} className="w-full sm:w-auto"><Calendar className="mr-1.5 size-4" aria-hidden />Aplicar</Button>
          {hasInvalidYear && <p className="text-xs text-destructive lg:pb-2">Informe um ano entre 2000 e 2100.</p>}
        </div>
      </div>

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
        <KpiCard label="Empréstimos no mês" value={totalEmprestimos} icon={ArrowUpRight} />
        <KpiCard label="Devoluções no mês" value={totalDevolucoes} icon={ArrowDownLeft} />
        <KpiCard label="Cadastros no mês" value={totalCadastros} icon={PackagePlus} />
      </div>

      <Section title="Empréstimos × Devoluções" description={`Movimentação diária em ${monthLabel}`}>
        {loans.isLoading ? <LoadingState label="Carregando indicadores..." /> : loans.error ? <ErrorState error={loans.error} onRetry={() => loans.refetch()} /> : loansChartData.every((d) => !d.Empréstimos && !d.Devoluções) ? <EmptyState title="Nenhuma movimentação registrada" description={`Não há empréstimos ou devoluções em ${monthLabel}.`} icon={<PackagePlus className="size-5" aria-hidden />} /> : (
          <div className="min-w-0 overflow-x-auto"><div className="min-w-[560px]"><ResponsiveContainer width="100%" height={320}><AreaChart data={loansChartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}><defs><linearGradient id="gradEmprestimos" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="var(--color-chart-1)" stopOpacity={0.3} /><stop offset="95%" stopColor="var(--color-chart-1)" stopOpacity={0.02} /></linearGradient><linearGradient id="gradDevolucoes" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="var(--color-chart-2)" stopOpacity={0.3} /><stop offset="95%" stopColor="var(--color-chart-2)" stopOpacity={0.02} /></linearGradient></defs><CartesianGrid strokeDasharray="3 3" className="stroke-border" vertical={false} /><XAxis dataKey="dia" tick={{ fontSize: 11 }} axisLine={false} tickLine={false} /><YAxis allowDecimals={false} tick={{ fontSize: 11 }} axisLine={false} tickLine={false} /><Tooltip contentStyle={tooltipStyle} /><Legend wrapperStyle={{ paddingTop: 10, fontSize: 12, fontWeight: 600 }} /><Area type="monotone" dataKey="Empréstimos" stroke="var(--color-chart-1)" strokeWidth={2.5} fill="url(#gradEmprestimos)" dot={{ r: 3 }} activeDot={{ r: 5 }} /><Area type="monotone" dataKey="Devoluções" stroke="var(--color-chart-2)" strokeWidth={2.5} fill="url(#gradDevolucoes)" dot={{ r: 3 }} activeDot={{ r: 5 }} /></AreaChart></ResponsiveContainer></div></div>
        )}
      </Section>

      <Section title="Novos cadastros" description={`Volume diário de novos patrimônios em ${monthLabel}`}>
        {registrations.isLoading ? <LoadingState label="Carregando indicadores..." /> : registrations.error ? <ErrorState error={registrations.error} onRetry={() => registrations.refetch()} /> : regChartData.every((d) => !d.Cadastros) ? <EmptyState title="Nenhum cadastro registrado" description={`Não há novos cadastros em ${monthLabel}.`} icon={<PackagePlus className="size-5" aria-hidden />} /> : (
          <div className="min-w-0 overflow-x-auto"><div className="min-w-[560px]"><ResponsiveContainer width="100%" height={280}><AreaChart data={regChartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}><defs><linearGradient id="gradCadastros" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="var(--color-chart-4)" stopOpacity={0.3} /><stop offset="95%" stopColor="var(--color-chart-4)" stopOpacity={0.02} /></linearGradient></defs><CartesianGrid strokeDasharray="3 3" className="stroke-border" vertical={false} /><XAxis dataKey="dia" tick={{ fontSize: 11 }} axisLine={false} tickLine={false} /><YAxis allowDecimals={false} tick={{ fontSize: 11 }} axisLine={false} tickLine={false} /><Tooltip contentStyle={tooltipStyle} /><Area type="monotone" dataKey="Cadastros" stroke="var(--color-chart-4)" strokeWidth={2.5} fill="url(#gradCadastros)" dot={{ r: 3 }} activeDot={{ r: 5 }} /></AreaChart></ResponsiveContainer></div></div>
        )}
      </Section>
    </div>
  );
}
