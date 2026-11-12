import type { ItemDetailTab } from './types'

type TabModules = Record<string, { default: ItemDetailTab }>

/** Filtra por papel e ordena por `order`. Separada do `import.meta.glob` para ser testável. */
export function buildItemDetailTabs(modules: TabModules, role: string | undefined): ItemDetailTab[] {
  return Object.values(modules)
    .map((m) => m.default)
    .filter((tab) => !tab.roles || (role !== undefined && tab.roles.includes(role)))
    .sort((a, b) => a.order - b.order)
}

// Cada arquivo em ./tabs/ é uma aba; nenhum outro arquivo precisa ser editado
// ao acrescentar uma.
const modules = import.meta.glob<{ default: ItemDetailTab }>('./tabs/*.tsx', { eager: true })

export function getItemDetailTabs(role: string | undefined): ItemDetailTab[] {
  return buildItemDetailTabs(modules, role)
}
