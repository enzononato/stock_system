import api, { type Paginated } from './client'

export type TermoTipo = 'entrega' | 'devolucao' | 'uso'

export type TermoStatus =
  | 'agendado'
  | 'enviado'
  | 'assinado'
  | 'recusado'
  | 'expirado'
  | 'cancelado'
  | 'erro'
  | 'manual'

/** Solicitação de assinatura de um termo (`GET /api/items/{id}/termos`). */
export interface TermoSolicitacao {
  /** `null` no termo "legado": empréstimo confirmado antes da integração, sem solicitação gravada. */
  id: number | null
  item_id: number
  item_descricao?: string | null
  colaborador_id?: number | null
  colaborador_nome?: string | null
  tipo: TermoTipo
  history_id?: number | null
  modelo_versao?: string | null
  status: TermoStatus
  destinatario_email?: string | null
  solicitado_por: string
  solicitado_em: string
  agendado_para?: string | null
  enviado_em?: string | null
  assinado_em?: string | null
  cancelado_em?: string | null
  motivo?: string | null
  tem_documento: boolean
  tem_assinado: boolean
  legado: boolean
}

export interface TermosDoItem {
  revsign_ativo: boolean
  vigente: TermoSolicitacao | null
  historico: Paginated<TermoSolicitacao>
}

/** Solicitações ainda em aberto: uma nova solicitação do mesmo equipamento cancela estas. */
export const STATUS_TERMO_PENDENTES: TermoStatus[] = ['agendado', 'enviado', 'manual', 'erro']

export const TIPO_TERMO_ROTULO: Record<TermoTipo, string> = {
  entrega: 'Termo de entrega',
  devolucao: 'Termo de devolução',
  uso: 'Termo de uso',
}

export const STATUS_TERMO_ROTULO: Record<TermoStatus, string> = {
  agendado: 'Agendado',
  enviado: 'Enviado',
  assinado: 'Assinado',
  recusado: 'Recusado',
  expirado: 'Expirado',
  cancelado: 'Cancelado',
  erro: 'Erro',
  manual: 'Assinatura manual',
}

interface PageParams {
  limit: number
  offset: number
}

export async function getTermosDoItem(itemId: number, params: PageParams): Promise<TermosDoItem> {
  const res = await api.get(`/items/${itemId}/termos`, { params })
  return res.data as TermosDoItem
}

/**
 * `POST /api/items/{id}/termos`. `agendado_para` é um horário de Brasília no formato
 * `AAAA-MM-DDTHH:MM` (o valor de um `input type="datetime-local"`).
 */
export async function solicitarTermo(itemId: number, data: { agendado_para?: string }): Promise<TermoSolicitacao> {
  const res = await api.post(`/items/${itemId}/termos`, data)
  return res.data as TermoSolicitacao
}

export async function getTermosDoColaborador(
  colaboradorId: number,
  params: PageParams
): Promise<Paginated<TermoSolicitacao>> {
  const res = await api.get(`/colaboradores/${colaboradorId}/termos`, { params })
  return res.data as Paginated<TermoSolicitacao>
}

export async function cancelarTermo(id: number, motivo?: string): Promise<TermoSolicitacao> {
  const res = await api.post(`/termos/${id}/cancelar`, motivo ? { motivo } : {})
  return res.data as TermoSolicitacao
}

export async function reenviarTermo(id: number): Promise<TermoSolicitacao> {
  const res = await api.post(`/termos/${id}/reenviar`)
  return res.data as TermoSolicitacao
}

export async function anexarTermoAssinado(id: number, pdf: File): Promise<TermoSolicitacao> {
  const form = new FormData()
  form.append('signed_pdf', pdf)
  const res = await api.post(`/termos/${id}/assinado`, form)
  return res.data as TermoSolicitacao
}

export async function getArquivosTermo(id: number): Promise<{ documento_url: string | null; assinado_url: string | null }> {
  const res = await api.get(`/termos/${id}/arquivos`)
  return res.data as { documento_url: string | null; assinado_url: string | null }
}
