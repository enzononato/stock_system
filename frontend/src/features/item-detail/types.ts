import type { ComponentType } from 'react'
import type { Item } from '@/api/items'

/**
 * Uma aba da página de detalhe do equipamento. Cada fase acrescenta a sua
 * criando UM arquivo em `features/item-detail/tabs/` com um `export default`
 * deste tipo — o registro (`registry.ts`) descobre os arquivos sozinho.
 *
 * Ordens reservadas: dados 10, linha do tempo 20, termos 30, manutenções 40,
 * linha telefônica 50, acessos 60, histórico 90.
 */
export interface ItemDetailTab {
  /** Identificador estável, usado em `?aba=<id>`. */
  id: string
  label: string
  order: number
  /** Quando informado, só esses papéis veem a aba. */
  roles?: string[]
  Component: ComponentType<{ item: Item }>
}
