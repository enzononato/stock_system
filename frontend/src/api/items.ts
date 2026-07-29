import api, { type Paginated } from './client'

export interface Item {
  id: number
  tipo?: string
  brand?: string
  model?: string
  identificador?: string
  nota_fiscal?: string
  status?: string
  assigned_to?: string
  cpf?: string
  revenda?: string
  dominio?: string
  host?: string
  endereco_fisico?: string
  cpu?: string
  ram?: string
  storage?: string
  sistema?: string
  licenca?: string
  anydesk?: string
  setor?: string
  ip?: string
  mac?: string
  fornecedor?: string
  potencia_nominal?: string
  autonomia_estimada?: string
  ip_snmp?: string
  codigo_patrimonial?: string
  responsavel?: string
  local_instalacao?: string
  poe?: string
  quantidade_portas?: string
  date_registered?: string
  date_issued?: string
  peripheral_count?: number
}

export interface ListItemsParams {
  tipo?: string
  status?: string
  revenda?: string
  search?: string
  limit?: number
  offset?: number
}

/**
 * `GET /api/items` paginado: aceita `tipo`/`status`/`revenda`/`search` e
 * `limit`/`offset`, devolvendo `{ items, total }`. Os filtros são aplicados no
 * SQL, não no cliente.
 *
 * Atenção ao omitir `limit`: o backend assume DEFAULT_PAGE_SIZE (50). Telas que
 * usam a lista para preencher um select precisam passar um limite explícito,
 * ou truncam em silêncio.
 */
export async function listItemsPaginated(params?: ListItemsParams): Promise<Paginated<Item>> {
  const res = await api.get('/items', { params })
  return res.data as Paginated<Item>
}

export async function getItem(id: number) {
  const res = await api.get(`/items/${id}`)
  return res.data as Item
}

export async function createItem(data: Record<string, unknown>) {
  const res = await api.post('/items', data)
  return res.data as { detail: string; id: number }
}

export async function updateItem(id: number, data: Record<string, unknown>) {
  const res = await api.put(`/items/${id}`, data)
  return res.data
}

export async function removeItem(id: number, reason: string, attachment?: File) {
  const form = new FormData()
  form.append('reason', reason)
  if (attachment) form.append('attachment', attachment)
  const res = await api.delete(`/items/${id}`, { data: form })
  return res.data
}
