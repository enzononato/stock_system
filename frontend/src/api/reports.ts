import api from './client'

export interface ReportRow {
  history_id?: number
  item_id?: number
  peripheral_id?: number
  operador?: string
  usuario?: string
  cpf?: string
  cargo?: string
  center_cost?: string
  setor?: string
  fornecedor?: string
  revenda?: string
  details?: string
  data_emprestimo?: string
  data_confirmacao?: string
  data_devolucao?: string
  operation_type?: string
  tipo?: string
  brand?: string
  model?: string
  identificador?: string
  nota_fiscal?: string
}

export interface ChartData {
  days: number[]
  values: number[]
  values2?: number[]
}

export async function getMonthlyReport(year: number, month: number) {
  const res = await api.get('/reports/monthly', { params: { year, month } })
  return res.data as ReportRow[]
}

export function getMonthlyReportExportUrl(year: number, month: number) {
  return `/api/reports/monthly/export?year=${year}&month=${month}`
}

export async function getLoansChart(year: number, month: number) {
  const res = await api.get('/reports/charts/loans', { params: { year, month } })
  return res.data as ChartData
}

export async function getRegistrationsChart(year: number, month: number) {
  const res = await api.get('/reports/charts/registrations', { params: { year, month } })
  return res.data as ChartData
}
