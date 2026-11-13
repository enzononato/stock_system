import api from './client'

/** O que a etiqueta imprime: nenhum dado pessoal. `url` (o conteúdo do QR) vem sempre do backend. */
export interface Etiqueta {
  id: number
  tipo?: string | null
  brand?: string | null
  model?: string | null
  identificador?: string | null
  revenda?: string | null
  url: string
}

export interface EtiquetasResposta {
  base_url: string
  etiquetas: Etiqueta[]
  /** Ids pedidos que não saíram: baixados ou inexistentes. */
  ignorados: number[]
}

/** `GET /api/etiquetas/ids` — todos os itens ativos da unidade (e dos filtros), em ordem crescente. */
export async function listIdsParaEtiquetas(params: {
  revenda: string
  tipo?: string
  status?: string
  search?: string
}): Promise<{ ids: number[]; total: number }> {
  const res = await api.get('/etiquetas/ids', { params })
  return res.data as { ids: number[]; total: number }
}

/** `POST /api/etiquetas` — até 2000 itens por vez. */
export async function gerarEtiquetas(itemIds: number[]): Promise<EtiquetasResposta> {
  const res = await api.post('/etiquetas', { item_ids: itemIds })
  return res.data as EtiquetasResposta
}
