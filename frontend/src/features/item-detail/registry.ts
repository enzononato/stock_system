import { buildTabRegistry, type TabModules } from '@/lib/tabRegistry'
import type { ItemDetailTab } from './types'

/**
 * Filtra por papel e ordena por `order` (ver `buildTabRegistry`). Separada do
 * `import.meta.glob` para ser testável: módulos sem `default` válido são
 * ignorados e ids repetidos geram aviso em desenvolvimento.
 */
export function buildItemDetailTabs(modules: TabModules, role: string | undefined): ItemDetailTab[] {
  return buildTabRegistry<ItemDetailTab>(modules, role, 'item-detail')
}

// Cada arquivo em ./tabs/ é uma aba; nenhum outro arquivo precisa ser editado
// ao acrescentar uma. Arquivos de teste ficam de fora.
const modules = import.meta.glob<{ default: ItemDetailTab }>(['./tabs/*.tsx', '!./tabs/*.test.tsx'], {
  eager: true,
})

export function getItemDetailTabs(role: string | undefined): ItemDetailTab[] {
  return buildItemDetailTabs(modules, role)
}
