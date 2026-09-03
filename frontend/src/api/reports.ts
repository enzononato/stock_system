import api, { downloadAuthenticated } from "./client";

export interface ReportRow {
  history_id?: number;
  item_id?: number;
  peripheral_id?: number;
  operador?: string;
  usuario?: string;
  cpf?: string;
  cargo?: string;
  center_cost?: string;
  setor?: string;
  fornecedor?: string;
  revenda?: string;
  details?: string;
  data_emprestimo?: string;
  data_confirmacao?: string;
  data_devolucao?: string;
  operation_type?: string;
  tipo?: string;
  brand?: string;
  model?: string;
  identificador?: string;
  nota_fiscal?: string;
}

export interface ChartData {
  days: number[];
  values: number[];
  values2?: number[];
}

export async function getMonthlyReport(year: number, month: number) {
  const res = await api.get("/reports/monthly", { params: { year, month } });
  return res.data as ReportRow[];
}

/**
 * Exporta o relatório mensal em CSV baixando-o de forma autenticada (Bearer
 * token via axios) e disparando o download no navegador — substitui o antigo
 * `<a href={getMonthlyReportExportUrl(...)} download>`, que sempre retornava
 * 401 porque o navegador não envia o Authorization header numa navegação crua.
 */
export async function exportMonthlyReportCsv(year: number, month: number): Promise<void> {
  await downloadAuthenticated(
    `/reports/monthly/export?year=${year}&month=${month}`,
    `relatorio_mensal_${year}_${String(month).padStart(2, "0")}.csv`,
  );
}

export async function getLoansChart(year: number, month: number) {
  const res = await api.get("/reports/charts/loans", { params: { year, month } });
  return res.data as ChartData;
}

export async function getRegistrationsChart(year: number, month: number) {
  const res = await api.get("/reports/charts/registrations", { params: { year, month } });
  return res.data as ChartData;
}
