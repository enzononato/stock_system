import api, { downloadAuthenticated, type Paginated } from './client'

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
  colaborador_id?: number | null
  local_id?: number | null
  local_nome?: string | null
  colaborador_nome?: string | null
  colaborador_setor?: string | null
  colaborador_cargo?: string | null
  superior_nome?: string | null
  /** Compra (F4): valor como texto (`"4599.90"`), datas só-dia `AAAA-MM-DD`. */
  valor_compra?: string | null
  data_compra?: string | null
  garantia_ate?: string | null
  nf_anexo_key?: string | null
}

export interface ListItemsParams {
  tipo?: string
  status?: string
  revenda?: string
  search?: string
  /** Só equipamentos com este colaborador. */
  colaborador_id?: number
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

/** Registro de apagamento de dados (LGPD) exigido na baixa de equipamentos com armazenamento. */
export interface ApagamentoInput {
  dados_apagados: 'sim' | 'nao_aplica'
  metodo_apagamento?: string
  apagado_por?: string
  /** dd/mm/aaaa */
  apagado_em?: string
  comprovante?: File | null
}

export async function removeItem(id: number, reason: string, attachment?: File, apagamento?: ApagamentoInput) {
  const form = new FormData()
  form.append('reason', reason)
  if (attachment) form.append('attachment', attachment)
  if (apagamento) {
    form.append('dados_apagados', apagamento.dados_apagados)
    if (apagamento.metodo_apagamento) form.append('metodo_apagamento', apagamento.metodo_apagamento)
    if (apagamento.apagado_por) form.append('apagado_por', apagamento.apagado_por)
    if (apagamento.apagado_em) form.append('apagado_em', apagamento.apagado_em)
    if (apagamento.comprovante) form.append('comprovante_apagamento', apagamento.comprovante)
  }
  const res = await api.delete(`/items/${id}`, { data: form })
  return res.data
}

/** `POST /items/{id}/nota-fiscal` (multipart, campo `arquivo`; PDF/PNG/JPG até 10 MB). */
export async function uploadNotaFiscal(
  id: number,
  arquivo: File
): Promise<{ detail: string; nf_anexo_key: string }> {
  const form = new FormData()
  form.append('arquivo', arquivo)
  const res = await api.post(`/items/${id}/nota-fiscal`, form)
  return res.data
}

/** Baixa a nota fiscal anexada (autenticado); o nome usa a extensão da chave do storage. */
export async function baixarNotaFiscal(item: Pick<Item, 'id' | 'nf_anexo_key'>): Promise<void> {
  const extensao = item.nf_anexo_key?.split('.').pop() ?? 'pdf'
  await downloadAuthenticated(`/items/${item.id}/nota-fiscal`, `nota_fiscal_item_${item.id}.${extensao}`)
}

export interface TransferenciaInput {
  para_unidade_id: number
  para_local_id: number | null
  motivo: string
}

export async function transferirItem(
  id: number,
  data: TransferenciaInput
): Promise<{ detail: string; id: number }> {
  const res = await api.post(`/items/${id}/transferir`, data)
  return res.data
}

export interface EventoLinhaDoTempo {
  ocorrido_em: string
  origem: string
  tipo: string
  titulo: string
  descricao?: string | null
  usuario?: string | null
  referencia_id?: number | null
}

export interface LinhaDoTempo {
  item_id: number
  eventos: EventoLinhaDoTempo[]
  fontes_com_erro: string[]
}

/** `GET /items/{id}/linha-do-tempo` — eventos em ordem cronológica (Gestor e Técnico). */
export async function getLinhaDoTempo(id: number): Promise<LinhaDoTempo> {
  const res = await api.get(`/items/${id}/linha-do-tempo`)
  return res.data as LinhaDoTempo
}
