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
  /**
   * Busca server-side por operador, usuário, operação, revenda, tipo, marca,
   * modelo ou identificador. Precisa ser feita no backend: com paginação, um
   * filtro no cliente varreria apenas a página carregada e passaria a impressão
   * falsa de ter pesquisado o histórico inteiro.
   */
  search?: string
  limit?: number
  offset?: number
}

/**
 * `GET /api/history` paginado: aceita `search`/`limit`/`offset` e devolve
 * `{ items, total }`. A busca é resolvida no servidor — ver `ListHistoryParams`.
 */
export async function listHistoryPaginated(params?: ListHistoryParams): Promise<Paginated<HistoryEntry>> {
  const res = await api.get('/history', { params })
  return res.data as Paginated<HistoryEntry>
}

/**
 * Estorna um lançamento do histórico. O backend exige a senha do operador no
 * corpo: é uma operação destrutiva, e só o role não basta como autorização.
 */
export async function reverseEntryWithPassword(historyId: number, password: string) {
  const res = await api.post(`/history/${historyId}/reverse`, { password })
  return res.data
}

