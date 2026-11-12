import api from './client'

/** Local físico dentro de uma unidade (sala, almoxarifado, andar...). */
export interface Local {
  id: number
  unidade_id: number
  unidade_nome?: string
  nome: string
  descricao?: string | null
  ativo: boolean
}

export interface ListLocaisParams {
  unidade_id?: number
  include_inactive?: boolean
}

/** Criação: o nome é único por unidade (o backend responde 400 com `detail` se repetir). */
export interface LocalCreateInput {
  unidade_id: number
  nome: string
  descricao?: string
}

/** Edição: chave ausente = sem mudança. */
export interface LocalUpdateInput {
  nome?: string
  descricao?: string
  ativo?: boolean
}

/** `GET /api/locais` — sem `include_inactive`, só os locais ativos. */
export async function listLocais(params: ListLocaisParams = {}): Promise<Local[]> {
  const res = await api.get('/locais', { params })
  return res.data as Local[]
}

/** `POST /api/locais`. */
export async function createLocal(data: LocalCreateInput): Promise<{ id: number }> {
  const res = await api.post('/locais', data)
  return res.data as { id: number }
}

/** `PUT /api/locais/{id}` — aceita qualquer combinação de `nome`, `descricao` e `ativo`. */
export async function updateLocal(id: number, data: LocalUpdateInput): Promise<unknown> {
  const res = await api.put(`/locais/${id}`, data)
  return res.data
}
