import api, { type Paginated } from '@/api/client'
import type { Item } from '@/api/items'

/** Chaves aceitas por `GET /api/items?sort=` (lista branca do backend). */
export const ESTOQUE_SORTS = [
  'id',
  'tipo',
  'equipamento',
  'status',
  'colaborador',
  'setor',
  'cargo',
  'unidade',
  'cadastro',
] as const

export type EstoqueSort = (typeof ESTOQUE_SORTS)[number]
export type SortDir = 'asc' | 'desc'

export interface EstoqueParams {
  tipo?: string
  status?: string
  /** Nome original da unidade (nunca o texto em CAIXA ALTA exibido). */
  revenda?: string
  search?: string
  colaborador_id?: number
  sort: EstoqueSort
  dir: SortDir
  limit: number
  offset: number
}

export async function listEstoque(params: EstoqueParams): Promise<Paginated<Item>> {
  const res = await api.get('/items', { params })
  return res.data as Paginated<Item>
}

/** Clique numa coluna nova começa em `asc`; na coluna ativa inverte a direção. */
export function proximaOrdenacao(
  atual: { sort: EstoqueSort; dir: SortDir },
  coluna: EstoqueSort
): { sort: EstoqueSort; dir: SortDir } {
  if (atual.sort === coluna) return { sort: coluna, dir: atual.dir === 'asc' ? 'desc' : 'asc' }
  return { sort: coluna, dir: 'asc' }
}
