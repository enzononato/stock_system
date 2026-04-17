import api from './client'

export interface HistoryEntry {
  id: number
  item_id?: number
  peripheral_id?: number
  operador?: string
  operation?: string
  revenda?: string
  data_operacao?: string
  tipo?: string
  marca?: string
  modelo?: string
  nota_fiscal?: string
  fornecedor?: string
  identificador?: string
  usuario?: string
  cpf?: string
  cargo?: string
  center_cost?: string
  setor?: string
  details?: string
}

export async function listHistory() {
  const res = await api.get('/history')
  return res.data as HistoryEntry[]
}

export async function reverseEntry(historyId: number) {
  const res = await api.post(`/history/${historyId}/reverse`)
  return res.data
}
