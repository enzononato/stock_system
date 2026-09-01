import { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { Download } from "lucide-react";

import { getMonthlyReport, exportMonthlyReportCsv, type ReportRow } from "@/api/reports";
import { DataTable, type Column } from "@/components/app/DataTable";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { formatDateTime } from "@/lib/utils";
import { getErrorMessage } from "@/lib/api-error";
import { toast } from "sonner";

const MONTHS = [
  "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
];

export function ReportPage() {
  const currentDate = new Date();
  const [year, setYear] = useState(String(currentDate.getFullYear()));
  const [month, setMonth] = useState(String(currentDate.getMonth() + 1));
  const [queryParams, setQueryParams] = useState({ year: currentDate.getFullYear(), month: currentDate.getMonth() + 1 });
  const [isExporting, setIsExporting] = useState(false);
  const [filterRevenda, setFilterRevenda] = useState("all");

  const { data: report = [], isLoading, error, refetch } = useQuery({
    queryKey: ["report", queryParams.year, queryParams.month],
    queryFn: () => getMonthlyReport(queryParams.year, queryParams.month),
  });

  const revendaOptions = [...new Set(report.map((r) => r.revenda).filter(Boolean))] as string[];
  const filteredReport = filterRevenda === "all" ? report : report.filter((r) => r.revenda === filterRevenda);

  function handleGenerate() {
    setQueryParams({ year: Number(year), month: Number(month) });
    void refetch();
  }

  async function handleExport() {
    setIsExporting(true);
    try {
      await exportMonthlyReportCsv(queryParams.year, queryParams.month);
    } catch (err) {
      toast.error(getErrorMessage(err, "Erro ao exportar relatório."));
    } finally {
      setIsExporting(false);
    }
  }

  const columns: Column<ReportRow>[] = [
    { key: "item_id", header: "ID Item", cell: (r) => r.item_id ?? "-" },
    { key: "operador", header: "Operador", cell: (r) => r.operador ?? "-", primary: true },
    { key: "operation_type", header: "Operação", cell: (r) => r.operation_type ?? "-" },
    { key: "tipo", header: "Tipo", cell: (r) => r.tipo ?? "-", hideBelow: "md" },
    { key: "brand", header: "Marca", cell: (r) => r.brand ?? "-", hideBelow: "lg" },
    { key: "model", header: "Modelo", cell: (r) => r.model ?? "-", hideBelow: "lg" },
    { key: "identificador", header: "Identificador", cell: (r) => r.identificador || "-", hideBelow: "lg" },
    { key: "nota_fiscal", header: "Nota Fiscal", cell: (r) => r.nota_fiscal || "-", hideBelow: "xl" },
    { key: "fornecedor", header: "Fornecedor", cell: (r) => r.fornecedor || "-", hideBelow: "xl" },
    { key: "usuario", header: "Usuário", cell: (r) => r.usuario || "-", hideBelow: "md" },
    { key: "cpf", header: "CPF", cell: (r) => r.cpf || "-", hideBelow: "xl" },
    { key: "cargo", header: "Cargo", cell: (r) => r.cargo || "-", hideBelow: "xl" },
    { key: "setor", header: "Setor", cell: (r) => r.setor || "-", hideBelow: "lg" },
    { key: "revenda", header: "Revenda", cell: (r) => r.revenda || "-", hideBelow: "md" },
    { key: "center_cost", header: "C. Custo", cell: (r) => r.center_cost || "-", hideBelow: "xl" },
    { key: "data_emprestimo", header: "Data", cell: (r) => formatDateTime(r.data_emprestimo) },
    { key: "data_confirmacao", header: "Confirmação", cell: (r) => formatDateTime(r.data_confirmacao), hideBelow: "lg" },
    { key: "data_devolucao", header: "Devolução", cell: (r) => formatDateTime(r.data_devolucao), hideBelow: "lg" },
    { key: "details", header: "Detalhes", cell: (r) => r.details || "-", hideBelow: "xl" },
  ];

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Relatórios"
        title="Relatório Mensal"
        description="Visualize todas as operações de um determinado mês."
      />

      <Section>
        <div className="flex flex-wrap items-end gap-4">
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="report-year">Ano</Label>
            <Input id="report-year" value={year} onChange={(e) => setYear(e.target.value)} className="w-24" />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Mês</Label>
            <Select value={month} onValueChange={setMonth}>
              <SelectTrigger className="w-40">
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
          <Button onClick={handleGenerate}>Gerar Relatório</Button>
          <Button variant="outline" onClick={handleExport} disabled={isExporting}>
            <Download className="mr-2 size-4" aria-hidden />
            {isExporting ? "Exportando..." : "Exportar CSV"}
          </Button>
          <div className="ml-auto flex flex-col gap-1.5">
            <Label>Revenda</Label>
            <Select value={filterRevenda} onValueChange={setFilterRevenda}>
              <SelectTrigger className="w-48">
                <SelectValue placeholder="Todas as revendas" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todas as revendas</SelectItem>
                {revendaOptions.map((r) => (
                  <SelectItem key={r} value={r}>
                    {r}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>
      </Section>

      <Section title="Resultado">
        <DataTable
          data={filteredReport}
          columns={columns}
          rowKey={(r) => r.history_id ?? `${r.item_id ?? "x"}-${r.data_emprestimo ?? r.data_confirmacao ?? r.data_devolucao ?? "0"}-${r.usuario ?? ""}`}
          isLoading={isLoading}
          error={error}
          onRetry={() => void refetch()}
          emptyTitle="Nenhum registro no período"
          emptyDescription="Ajuste o ano/mês e gere o relatório novamente."
        />
      </Section>
    </div>
  );
}
