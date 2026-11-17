import api, { downloadAuthenticated, type Paginated } from './client'

export type StatusInventario = 'aberto' | 'encerrado' | 'cancelado'
export type ResultadoLeitura =
  | 'ok'
  | 'local_errado'
  | 'usuario_diferente'
  | 'outra_unidade'
  | 'nao_cadastrado'
export type CategoriaRelatorio = ResultadoLeitura | 'nao_encontrado'

export interface ResumoInventario {
  esperados: number
  lidos: number
  encontrados: number
  ok: number
  local_errado: number
  usuario_diferente: number
  outra_unidade: number
  nao_cadastrado: number
  nao_encontrados: number
}

export interface Inventario {
  id: number
  unidade_id: number
  unidade_nome: string
  status: StatusInventario
  iniciado_por: string
  iniciado_em: string
  encerrado_por?: string | null
  encerrado_em?: string | null
  observacao?: string | null
  motivo_cancelamento?: string | null
  total_leituras: number
}

export interface InventarioDetalhe extends Inventario {
  resumo: ResumoInventario
}

export interface Leitura {
  id: number
  inventario_id: number
  item_id?: number | null
  codigo_lido: string
  lido_por: string
  lido_em: string
  resultado: ResultadoLeitura
  detalhe?: string | null
  observacao?: string | null
  atualizado_em?: string | null
  local_id_encontrado?: number | null
  local_encontrado_nome?: string | null
  usuario_conferido: boolean
  colaborador_id_encontrado?: number | null
  colaborador_encontrado_nome?: string | null
  local_id_esperado?: number | null
  local_esperado_nome?: string | null
  usuario_esperado?: string | null
  item_tipo?: string | null
  item_brand?: string | null
  item_model?: string | null
  item_identificador?: string | null
  item_revenda?: string | null
  /** Só vem preenchido na resposta do POST: a mesma leitura repetida não conta de novo. */
  duplicada?: boolean
}

export interface RelatorioLinha {
  categoria: CategoriaRelatorio
  leitura_id?: number | null
  item_id?: number | null
  tipo?: string | null
  brand?: string | null
  model?: string | null
  identificador?: string | null
  status_item?: string | null
  unidade_item?: string | null
  local_esperado?: string | null
  local_encontrado?: string | null
  usuario_esperado?: string | null
  usuario_encontrado?: string | null
  codigo_lido?: string | null
  lido_por?: string | null
  lido_em?: string | null
  detalhe?: string | null
  observacao?: string | null
}

export interface RelatorioInventario {
  inventario: InventarioDetalhe
  linhas: RelatorioLinha[]
}

export interface ListInventariosParams {
  unidade_id?: number
  status?: StatusInventario
  limit?: number
  offset?: number
}

export interface ListLeiturasParams {
  resultado?: ResultadoLeitura
  limit?: number
  offset?: number
}

export interface AtualizarLeituraInput {
  local_id_encontrado?: number | null
  colaborador_id_encontrado?: number | null
  usuario_conferido?: boolean
  observacao?: string | null
}

/** `GET /api/inventarios`. */
export async function listInventarios(params: ListInventariosParams = {}): Promise<Paginated<Inventario>> {
  const res = await api.get('/inventarios', { params })
  return res.data as Paginated<Inventario>
}

/** `POST /api/inventarios` — um inventário aberto por unidade (409 se já houver). */
export async function createInventario(data: {
  unidade_id: number
  observacao?: string
}): Promise<InventarioDetalhe> {
  const res = await api.post('/inventarios', data)
  return res.data as InventarioDetalhe
}

export async function getInventario(id: number): Promise<InventarioDetalhe> {
  const res = await api.get(`/inventarios/${id}`)
  return res.data as InventarioDetalhe
}

export async function encerrarInventario(id: number): Promise<InventarioDetalhe> {
  const res = await api.post(`/inventarios/${id}/encerrar`)
  return res.data as InventarioDetalhe
}

/** Só o Gestor cancela. */
export async function cancelarInventario(id: number, motivo: string): Promise<InventarioDetalhe> {
  const res = await api.post(`/inventarios/${id}/cancelar`, { motivo })
  return res.data as InventarioDetalhe
}

export async function listLeituras(id: number, params: ListLeiturasParams = {}): Promise<Paginated<Leitura>> {
  const res = await api.get(`/inventarios/${id}/leituras`, { params })
  return res.data as Paginated<Leitura>
}

/** `POST .../leituras`: 201 para leitura nova, 200 com `duplicada: true` para a repetida. */
export async function registrarLeitura(
  id: number,
  data: { codigo: string; local_id_encontrado?: number | null }
): Promise<Leitura> {
  const res = await api.post(`/inventarios/${id}/leituras`, data)
  return res.data as Leitura
}

export async function atualizarLeitura(
  id: number,
  leituraId: number,
  data: AtualizarLeituraInput
): Promise<Leitura> {
  const res = await api.patch(`/inventarios/${id}/leituras/${leituraId}`, data)
  return res.data as Leitura
}

export async function excluirLeitura(id: number, leituraId: number): Promise<void> {
  await api.delete(`/inventarios/${id}/leituras/${leituraId}`)
}

export async function getRelatorio(id: number): Promise<RelatorioInventario> {
  const res = await api.get(`/inventarios/${id}/relatorio`)
  return res.data as RelatorioInventario
}

/** Baixa o relatório em Excel ou PDF (com o Bearer token, por `downloadAuthenticated`). */
export function exportarRelatorio(id: number, formato: 'xlsx' | 'pdf'): Promise<void> {
  return downloadAuthenticated(`/inventarios/${id}/relatorio/exportar?formato=${formato}`, `inventario_${id}.${formato}`)
}
