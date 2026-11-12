import api, { type Paginated } from './client'

export type SituacaoColaborador = 'ativo' | 'afastado' | 'desligado'

/**
 * Colaborador vindo da base de RH (`origem='rh'`, sincronizado) ou cadastrado à
 * mão (`origem='manual'`). Para `origem='rh'`, só e-mail, telefone e superior
 * são editáveis no sistema; `email_manual`/`telefone_manual` indicam que o
 * valor foi digitado aqui e não será sobrescrito pela sincronização.
 */
export interface Colaborador {
  id: number
  cpf: string
  nome: string
  matricula?: string | null
  cargo?: string | null
  setor?: string | null
  unidade_rh?: string | null
  unidade_id?: number | null
  unidade_nome?: string | null
  situacao: SituacaoColaborador
  email?: string | null
  telefone?: string | null
  email_manual: boolean
  telefone_manual: boolean
  superior_id?: number | null
  superior_nome?: string | null
  origem: 'rh' | 'manual'
  sincronizado_em?: string | null
}

export interface ListColaboradoresParams {
  busca?: string
  situacao?: string
  unidade_id?: number
  limit?: number
  offset?: number
}

/** Payload de criação (cadastro manual) e edição (`PUT`). */
export interface ColaboradorInput {
  cpf?: string
  nome?: string
  matricula?: string | null
  cargo?: string | null
  setor?: string | null
  unidade_id?: number | null
  situacao?: SituacaoColaborador
  email?: string | null
  telefone?: string | null
  superior_id?: number | null
}

export interface ResumoSincronizacao {
  total_base: number
  criados: number
  atualizados: number
  desligados: number
}

export interface SyncStatus {
  ultima: {
    iniciado_em: string
    concluido_em: string | null
    status: 'ok' | 'erro' | 'em_andamento'
    total_base: number | null
    criados: number | null
    atualizados: number | null
    desligados: number | null
    erro: string | null
  } | null
  configurado: boolean
}

/** `GET /api/colaboradores` — filtros e paginação aplicados no servidor. */
export async function listColaboradores(
  params: ListColaboradoresParams = {}
): Promise<Paginated<Colaborador>> {
  const res = await api.get('/colaboradores', { params })
  return res.data as Paginated<Colaborador>
}

export async function getColaborador(id: number): Promise<Colaborador> {
  const res = await api.get(`/colaboradores/${id}`)
  return res.data as Colaborador
}

export async function getColaboradorPorCpf(cpf: string): Promise<Colaborador> {
  const res = await api.get(`/colaboradores/cpf/${cpf.replace(/\D/g, '')}`)
  return res.data as Colaborador
}

/** `POST /api/colaboradores` — cadastro manual; 409 se o CPF já existe. */
export async function createColaborador(data: ColaboradorInput): Promise<Colaborador> {
  const res = await api.post('/colaboradores', data)
  return res.data as Colaborador
}

/** `PUT /api/colaboradores/{id}` — campos da base de RH só valem para `origem='manual'`. */
export async function updateColaborador(id: number, data: ColaboradorInput): Promise<Colaborador> {
  const res = await api.put(`/colaboradores/${id}`, data)
  return res.data as Colaborador
}

/** `POST /api/colaboradores/sincronizar` — Gestor; 503 se a base de RH estiver indisponível. */
export async function sincronizarColaboradores(): Promise<ResumoSincronizacao> {
  const res = await api.post('/colaboradores/sincronizar')
  return res.data as ResumoSincronizacao
}

export async function getSyncStatus(): Promise<SyncStatus> {
  const res = await api.get('/colaboradores/sincronizacao')
  return res.data as SyncStatus
}

export async function listCargos(): Promise<string[]> {
  const res = await api.get('/colaboradores/cargos')
  return res.data as string[]
}
