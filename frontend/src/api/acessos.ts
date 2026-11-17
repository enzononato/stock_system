import api, { type Paginated } from './client'

/**
 * Módulo de Acessos (F8). O sistema nunca pede nem guarda senha: nenhum tipo
 * ou função deste arquivo carrega esse campo, só o login.
 */

export type TipoEvento = 'admissao' | 'desligamento' | 'mudanca_funcao'
export type StatusEvento = 'aberto' | 'concluido' | 'cancelado'
export type AcaoTarefa = 'criar' | 'desativar'
export type StatusTarefa = 'pendente' | 'feito' | 'nao_se_aplica'
export type StatusAcesso = 'ativo' | 'desativado'

export interface Sistema {
  id: number
  nome: string
  descricao: string | null
  ativo: boolean
}

export interface Pacote {
  cargo: string
  sistemas: Sistema[]
}

export interface Acesso {
  id: number
  colaborador_id: number
  colaborador_nome: string
  sistema_id: number
  sistema_nome: string
  login: string
  status: StatusAcesso
  item_id: number | null
  item_descricao: string | null
  criado_por: string
  criado_em: string
  desativado_por: string | null
  desativado_em: string | null
}

export interface EquipamentoResumo {
  id: number
  tipo: string | null
  brand: string | null
  model: string | null
  identificador: string | null
  status: string
  revenda: string | null
}

export interface Tarefa {
  id: number
  evento_id: number
  sistema_id: number
  sistema_nome: string
  acao: AcaoTarefa
  status: StatusTarefa
  login: string | null
  acesso_id: number | null
  feito_por: string | null
  feito_em: string | null
  observacao: string | null
}

export interface EventoResumo {
  id: number
  colaborador_id: number
  colaborador_nome: string
  tipo: TipoEvento
  cargo_anterior: string | null
  cargo_novo: string | null
  status: StatusEvento
  observacao: string | null
  aberto_por: string
  aberto_em: string
  concluido_por: string | null
  concluido_em: string | null
  cancelado_por: string | null
  cancelado_em: string | null
  motivo_cancelamento: string | null
  total_tarefas: number
  tarefas_pendentes: number
}

export interface EventoDetalhe extends EventoResumo {
  colaborador_situacao: string
  tarefas: Tarefa[]
  equipamentos: EquipamentoResumo[]
  pode_concluir: boolean
  pendencias: string[]
}

// ── Catálogo e pacotes ───────────────────────────────────────────────────────

export async function listSistemas(params: { include_inactive?: boolean } = {}): Promise<Sistema[]> {
  const res = await api.get('/acessos/sistemas', { params })
  return res.data as Sistema[]
}

export async function createSistema(data: { nome: string; descricao: string | null }): Promise<Sistema> {
  const res = await api.post('/acessos/sistemas', data)
  return res.data as Sistema
}

export async function updateSistema(
  id: number,
  data: { nome?: string; descricao?: string | null; ativo?: boolean }
): Promise<Sistema> {
  const res = await api.put(`/acessos/sistemas/${id}`, data)
  return res.data as Sistema
}

export async function listPacotes(): Promise<Pacote[]> {
  const res = await api.get('/acessos/pacotes')
  return res.data as Pacote[]
}

export async function savePacote(data: { cargo: string; sistema_ids: number[] }): Promise<Pacote> {
  const res = await api.put('/acessos/pacotes', data)
  return res.data as Pacote
}

// ── Ficha de acessos ─────────────────────────────────────────────────────────

export async function listAcessos(params: {
  colaborador_id?: number
  item_id?: number
  status?: StatusAcesso
}): Promise<Acesso[]> {
  const res = await api.get('/acessos/ficha', { params })
  return res.data as Acesso[]
}

export async function createAcesso(data: {
  colaborador_id: number
  sistema_id: number
  login: string
  item_id: number | null
}): Promise<Acesso> {
  const res = await api.post('/acessos/ficha', data)
  return res.data as Acesso
}

export async function updateAcesso(
  id: number,
  data: { login?: string; item_id?: number | null }
): Promise<Acesso> {
  const res = await api.put(`/acessos/ficha/${id}`, data)
  return res.data as Acesso
}

export async function desativarAcesso(id: number): Promise<Acesso> {
  const res = await api.post(`/acessos/ficha/${id}/desativar`)
  return res.data as Acesso
}

export async function listEquipamentosDoColaborador(colaboradorId: number): Promise<EquipamentoResumo[]> {
  const res = await api.get(`/acessos/colaboradores/${colaboradorId}/equipamentos`)
  return res.data as EquipamentoResumo[]
}

// ── Eventos e tarefas ────────────────────────────────────────────────────────

export async function listEventos(params: {
  status?: StatusEvento
  tipo?: TipoEvento
  colaborador_id?: number
  busca?: string
  limit?: number
  offset?: number
}): Promise<Paginated<EventoResumo>> {
  const res = await api.get('/acessos/eventos', { params })
  return res.data as Paginated<EventoResumo>
}

export async function getEvento(id: number): Promise<EventoDetalhe> {
  const res = await api.get(`/acessos/eventos/${id}`)
  return res.data as EventoDetalhe
}

export async function createEvento(data: {
  colaborador_id: number
  tipo: TipoEvento
  cargo_anterior: string | null
  cargo_novo: string | null
  observacao: string | null
}): Promise<EventoDetalhe> {
  const res = await api.post('/acessos/eventos', data)
  return res.data as EventoDetalhe
}

export async function resolverTarefa(
  id: number,
  data: {
    status: 'feito' | 'nao_se_aplica' | 'manter'
    login?: string
    item_id?: number | null
    observacao?: string
  }
): Promise<EventoDetalhe> {
  const res = await api.post(`/acessos/tarefas/${id}/resolver`, data)
  return res.data as EventoDetalhe
}

export async function concluirEvento(id: number): Promise<EventoDetalhe> {
  const res = await api.post(`/acessos/eventos/${id}/concluir`)
  return res.data as EventoDetalhe
}

export async function cancelarEvento(
  id: number,
  data: { motivo_cancelamento: string }
): Promise<EventoDetalhe> {
  const res = await api.post(`/acessos/eventos/${id}/cancelar`, data)
  return res.data as EventoDetalhe
}
