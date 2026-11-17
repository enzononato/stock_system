import type { CategoriaRelatorio, StatusInventario } from '@/api/inventarios'
import type { StatusTone } from '@/lib/status'

/** Singular, igual a `ROTULOS_CATEGORIA` do backend (uma leitura/linha). */
export const ROTULO_RESULTADO: Record<CategoriaRelatorio, string> = {
  ok: 'Conferido',
  local_errado: 'Local errado',
  usuario_diferente: 'Usuário diferente',
  outra_unidade: 'De outra unidade',
  nao_cadastrado: 'Não cadastrado',
  nao_encontrado: 'Não encontrado',
}

/** Plural, para abas e indicadores. */
export const ROTULO_CATEGORIA: Record<CategoriaRelatorio, string> = {
  ok: 'Conferidos',
  local_errado: 'Local errado',
  usuario_diferente: 'Usuário diferente',
  outra_unidade: 'De outra unidade',
  nao_cadastrado: 'Não cadastrados',
  nao_encontrado: 'Não encontrados',
}

export const TOM_RESULTADO: Record<CategoriaRelatorio, StatusTone> = {
  ok: 'available',
  local_errado: 'pending',
  usuario_diferente: 'return',
  outra_unidade: 'maintenance',
  nao_cadastrado: 'danger',
  nao_encontrado: 'danger',
}

/** Mesma ordem do relatório do backend. */
export const ORDEM_CATEGORIAS: CategoriaRelatorio[] = [
  'nao_encontrado',
  'nao_cadastrado',
  'outra_unidade',
  'usuario_diferente',
  'local_errado',
  'ok',
]

/** Estes rótulos já mapeiam para pending/retired em `statusTone`. */
export const ROTULO_STATUS_INVENTARIO: Record<StatusInventario, string> = {
  aberto: 'Aberto',
  encerrado: 'Encerrado',
  cancelado: 'Cancelado',
}

/** Com o inventário aberto, "não encontrado" ainda é só "não lido". */
export function rotuloCategoria(c: CategoriaRelatorio, status: StatusInventario): string {
  return c === 'nao_encontrado' && status === 'aberto' ? 'Ainda não lidos' : ROTULO_CATEGORIA[c]
}
