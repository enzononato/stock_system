import api from './client'

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

export async function listItems(params?: {
  tipo?: string
  status?: string
  revenda?: string
  search?: string
}) {
  const res = await api.get('/items', { params })
  return res.data as Item[]
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
