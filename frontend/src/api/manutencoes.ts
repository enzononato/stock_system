import api, { downloadAuthenticated } from './client'

export type StatusManutencao =
  | 'aberta'
  | 'enviada'
  | 'em_reparo'
  | 'retornada'
  | 'sem_conserto'
  | 'cancelada'

/** Espelho de `ManutencaoResponse`. Valores monetários vêm como texto (`"150.00"`), datas como ISO. */
export interface Manutencao {
  id: number
  item_id: number
  status: StatusManutencao
  defeito: string
  assistencia?: string | null
  aberta_por: string
  aberta_em: string
  enviada_em?: string | null
  em_reparo_em?: string | null
  /** Só-dia: `AAAA-MM-DD` (converter por texto, nunca por `new Date`). */
  previsao_retorno?: string | null
  retornada_em?: string | null
  cancelada_em?: string | null
  encerrada_por?: string | null
  custo?: string | null
  anexo_key?: string | null
  observacoes?: string | null
  item_tipo?: string | null
  item_brand?: string | null
  item_model?: string | null
  item_identificador?: string | null
  item_revenda?: string | null
  item_status?: string | null
  custo_acumulado_item: string
}

export interface ManutencaoLista {
  items: Manutencao[]
  total: number
  /** Soma do custo de TODAS as linhas filtradas (não só da página). */
  custo_total: string
}

export interface ListManutencoesParams {
  status?: StatusManutencao | 'em_andamento'
  item_id?: number
  revenda?: string
  tipo?: string
  /** `AAAA-MM-DD`, inclusivo. */
  de?: string
  ate?: string
  busca?: string
  limit?: number
  offset?: number
}

export async function listManutencoes(params?: ListManutencoesParams): Promise<ManutencaoLista> {
  const res = await api.get('/manutencoes', { params })
  return res.data as ManutencaoLista
}

export async function getManutencao(id: number): Promise<Manutencao> {
  const res = await api.get(`/manutencoes/${id}`)
  return res.data as Manutencao
}

export async function abrirManutencao(data: {
  item_id: number
  defeito: string
  assistencia?: string
  observacoes?: string
}): Promise<{ detail: string; id: number }> {
  const res = await api.post('/manutencoes', data)
  return res.data
}

export async function enviarManutencao(
  id: number,
  data: { assistencia: string; previsao_retorno?: string }
): Promise<{ detail: string }> {
  const res = await api.post(`/manutencoes/${id}/enviar`, data)
  return res.data
}

export async function marcarEmReparo(
  id: number,
  data: { previsao_retorno?: string }
): Promise<{ detail: string }> {
  const res = await api.post(`/manutencoes/${id}/em-reparo`, data)
  return res.data
}

export interface RetornoInput {
  resultado: 'retornada' | 'sem_conserto'
  custo?: number | null
  observacoes?: string
  anexo?: File | null
}

/** Multipart: o `custo` vai com ponto e duas casas (`150.00`); campos vazios não são enviados. */
export async function registrarRetorno(
  id: number,
  data: RetornoInput
): Promise<{ detail: string; sugerir_baixa: boolean }> {
  const form = new FormData()
  form.append('resultado', data.resultado)
  if (data.custo !== null && data.custo !== undefined) form.append('custo', data.custo.toFixed(2))
  if (data.observacoes) form.append('observacoes', data.observacoes)
  if (data.anexo) form.append('anexo', data.anexo)
  const res = await api.post(`/manutencoes/${id}/retorno`, form)
  return res.data
}

export async function cancelarManutencao(id: number, data: { motivo: string }): Promise<{ detail: string }> {
  const res = await api.post(`/manutencoes/${id}/cancelar`, data)
  return res.data
}

export async function baixarAnexoManutencao(m: Pick<Manutencao, 'id' | 'anexo_key'>): Promise<void> {
  const extensao = m.anexo_key?.split('.').pop() ?? 'pdf'
  await downloadAuthenticated(`/manutencoes/${m.id}/anexo`, `manutencao_${m.id}.${extensao}`)
}
