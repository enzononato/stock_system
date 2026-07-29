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
 * Contrato novo (paginado) de `GET /api/history`: aceita `limit`/`offset` e
 * devolve `{ items, total }`. Contrato acordado com o backend (T4) — use esta
 * função em código novo. `listHistory` (abaixo) é o shim de compatibilidade
 * com a assinatura antiga.
 */
export async function listHistoryPaginated(params?: ListHistoryParams): Promise<Paginated<HistoryEntry>> {
  const res = await api.get('/history', { params })
  return res.data as Paginated<HistoryEntry>
}

/**
 * @deprecated `GET /api/history` passou a responder `{ items, total }` (ver
 * `listHistoryPaginated`). Esta função mantém a assinatura antiga — sem
 * parâmetros, devolvendo `HistoryEntry[]` — desembrulhando `.items`
 * internamente. HistoryPage passa esta função *diretamente* como `queryFn`
 * (`queryFn: listHistory`, sem wrapper em arrow function), então a aridade
 * precisa casar exatamente com o `QueryFunction` esperado pelo TanStack Query;
 * um parâmetro opcional aqui já é suficiente para o `tsc` rejeitar por causa
 * da checagem de "weak type" (o objeto de contexto do TanStack Query não
 * compartilha nenhuma propriedade com `ListHistoryParams`). Remova esta
 * função e troque a chamada por `listHistoryPaginated` quando a página
 * migrar (Módulo 5/6).
 */
export async function listHistory(): Promise<HistoryEntry[]> {
  const { items } = await listHistoryPaginated()
  return items
}

/**
 * Contrato novo: o backend passou a exigir a senha do operador no corpo para
 * confirmar o estorno. Use esta função em código novo. `reverseEntry` (abaixo)
 * é o shim de compatibilidade com a assinatura antiga (sem senha).
 */
export async function reverseEntryWithPassword(historyId: number, password: string) {
  const res = await api.post(`/history/${historyId}/reverse`, { password })
  return res.data
}

/**
 * @deprecated O backend agora exige a senha do operador no corpo do estorno
 * (ver `reverseEntryWithPassword`) — chamar esta função (sem senha) contra o
 * backend atualizado deve falhar (provavelmente 400/422 por faltar o campo
 * `password`). Mantida apenas para o HistoryPage compilar, que ainda chama
 * `reverseEntry(id)` com um único argumento e não coleta a senha do operador
 * (Módulo 5/6). Remova esta função e troque a chamada por
 * `reverseEntryWithPassword` quando a página ganhar esse campo.
 */
export async function reverseEntry(historyId: number) {
  const res = await api.post(`/history/${historyId}/reverse`)
  return res.data
}
