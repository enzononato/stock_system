import api, { type Paginated } from './client'

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
  /** Nome/caminho do anexo vinculado à operação (ex.: comprovante de remoção), quando houver. */
  operacao_anexo?: string
  /** Nome/caminho do termo assinado (empréstimo ou devolução) vinculado a esta entrada, quando houver. */
  termo_assinado_anexo?: string
}

export interface ListHistoryParams {
  limit?: number
  offset?: number
}

/**
 * `GET /api/history` agora responde `{ items, total }` (paginado) em vez de um
 * array cru, aceitando `limit`/`offset`. Contrato acordado com o backend (T4).
 */
export async function listHistory(params?: ListHistoryParams): Promise<Paginated<HistoryEntry>> {
  const res = await api.get('/history', { params })
  return res.data as Paginated<HistoryEntry>
}

/**
 * @deprecated Wrapper temporário que devolve o array cru para não quebrar em
 * runtime o HistoryPage, que ainda espera `HistoryEntry[]` de `listHistory`.
 * A página ainda vai falhar no `tsc` até ser migrada para o formato paginado
 * (Módulo 5/6) — remova este wrapper quando isso acontecer.
 */
export async function listHistoryArray(params?: ListHistoryParams): Promise<HistoryEntry[]> {
  const { items } = await listHistory(params)
  return items
}

/** O estorno agora exige a senha do operador para confirmar a ação (mesma regra do backend). */
export async function reverseEntry(historyId: number, password: string) {
  const res = await api.post(`/history/${historyId}/reverse`, { password })
  return res.data
}
