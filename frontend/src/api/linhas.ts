import api, { downloadAuthenticated, type Paginated } from './client'
import type { SituacaoColaborador } from './colaboradores'

export type StatusLinha = 'ativa' | 'suspensa' | 'cancelada'

export type TipoEventoLinha =
  | 'ativacao'
  | 'troca_aparelho'
  | 'troca_usuario'
  | 'portabilidade'
  | 'suspensao'
  | 'reativacao'
  | 'cancelamento'
  | 'edicao'

export const OPERADORAS = ['Vivo', 'Claro', 'TIM', 'Algar', 'Outra'] as const

/** Linha telefônica. `custo_mensal` chega como texto (`"49.90"`) para não perder centavos. */
export interface Linha {
  id: number
  numero: string
  operadora: string
  plano: string | null
  custo_mensal: string
  status: StatusLinha
  unidade_id: number
  unidade_nome: string
  item_id: number | null
  item_tipo: string | null
  item_brand: string | null
  item_model: string | null
  item_identificador: string | null
  item_ativo: boolean | null
  colaborador_id: number | null
  colaborador_nome: string | null
  colaborador_situacao: SituacaoColaborador | null
  iccid: string | null
  observacoes: string | null
  criada_em: string
  atualizada_em: string
}

export interface LinhaEvento {
  id: number
  linha_id: number
  linha_numero: string
  tipo: TipoEventoLinha
  campo: string | null
  de_valor: string | null
  para_valor: string | null
  item_id: number | null
  item_id_anterior: number | null
  ocorrido_em: string
  por_usuario: string | null
  observacao: string | null
}

/** Celular ou tablet que ainda pode receber uma linha. */
export interface AparelhoOpcao {
  id: number
  tipo: string
  brand: string | null
  model: string | null
  identificador: string | null
  codigo_patrimonial: string | null
  revenda: string | null
  status: string | null
  responsavel_nome: string | null
}

export interface CustoUnidade {
  unidade_id: number
  unidade_nome: string
  quantidade: number
  custo_total: string
}

export interface RelatorioCusto {
  unidades: CustoUnidade[]
  total_linhas: number
  custo_total: string
  gerado_em: string
}

/** Criação e edição. Na edição, chave ausente = sem mudança e `null` limpa (plano, aparelho, colaborador, ICCID, observações). */
export interface LinhaInput {
  numero?: string
  operadora?: string
  plano?: string | null
  custo_mensal?: string
  unidade_id?: number
  item_id?: number | null
  colaborador_id?: number | null
  iccid?: string | null
  observacoes?: string | null
  motivo?: string | null
}

export interface ListLinhasParams {
  busca?: string
  status?: StatusLinha
  unidade_id?: number
  operadora?: string
  item_id?: number
  colaborador_id?: number
  limit?: number
  offset?: number
}

/** `GET /api/linhas` — filtros e paginação no servidor. */
export async function listLinhas(params: ListLinhasParams = {}): Promise<Paginated<Linha>> {
  const res = await api.get('/linhas', { params })
  return res.data as Paginated<Linha>
}

export async function getLinha(id: number): Promise<Linha> {
  const res = await api.get(`/linhas/${id}`)
  return res.data as Linha
}

export async function createLinha(data: LinhaInput): Promise<Linha> {
  const res = await api.post('/linhas', data)
  return res.data as Linha
}

export async function updateLinha(id: number, data: LinhaInput): Promise<Linha> {
  const res = await api.put(`/linhas/${id}`, data)
  return res.data as Linha
}

/** `POST /api/linhas/{id}/status` — suspender (`suspensa`), reativar (`ativa`) ou cancelar (`cancelada`). */
export async function alterarStatusLinha(id: number, status: StatusLinha, observacao?: string): Promise<Linha> {
  const res = await api.post(`/linhas/${id}/status`, { status, observacao })
  return res.data as Linha
}

export async function listEventosLinha(id: number): Promise<LinhaEvento[]> {
  const res = await api.get(`/linhas/${id}/eventos`)
  return res.data as LinhaEvento[]
}

/** Eventos de todas as linhas que passaram pelo aparelho (inclusive as que saíram dele). */
export async function listEventosDoAparelho(itemId: number): Promise<LinhaEvento[]> {
  const res = await api.get('/linhas/eventos', { params: { item_id: itemId } })
  return res.data as LinhaEvento[]
}

export async function listAparelhos(params: {
  busca?: string
  linha_id?: number
  limit?: number
}): Promise<AparelhoOpcao[]> {
  const res = await api.get('/linhas/aparelhos', { params })
  return res.data as AparelhoOpcao[]
}

export async function getRelatorioCusto(): Promise<RelatorioCusto> {
  const res = await api.get('/linhas/relatorio/custo')
  return res.data as RelatorioCusto
}

function dataLocalCompacta(): string {
  const d = new Date()
  const dois = (n: number) => String(n).padStart(2, '0')
  return `${d.getFullYear()}${dois(d.getMonth() + 1)}${dois(d.getDate())}`
}

/** Baixa o relatório de custo em Excel ou PDF (`custo_linhas_AAAAMMDD.<formato>`). */
export async function exportarRelatorioCusto(formato: 'xlsx' | 'pdf'): Promise<void> {
  await downloadAuthenticated(
    `/linhas/relatorio/custo/export?formato=${formato}`,
    `custo_linhas_${dataLocalCompacta()}.${formato}`
  )
}
