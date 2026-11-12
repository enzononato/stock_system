import type { ItemDetailTab } from './types'

type TabModules = Record<string, { default?: unknown } | undefined>

function isItemDetailTab(v: unknown): v is ItemDetailTab {
  if (!v || typeof v !== 'object') return false
  const t = v as Partial<ItemDetailTab>
  return (
    typeof t.id === 'string' &&
    t.id !== '' &&
    typeof t.label === 'string' &&
    typeof t.order === 'number' &&
    t.Component != null
  )
}

/**
 * Filtra por papel e ordena por `order`. Separada do `import.meta.glob` para
 * ser testável. Módulos sem um `default` válido (helpers, arquivos soltos em
 * `tabs/`) são ignorados; ids repetidos geram aviso em desenvolvimento e vale
 * o primeiro (na ordem dos caminhos).
 */
export function buildItemDetailTabs(modules: TabModules, role: string | undefined): ItemDetailTab[] {
  const seen = new Set<string>()
  const tabs: ItemDetailTab[] = []
  for (const path of Object.keys(modules).sort()) {
    const candidate = modules[path]?.default
    if (!isItemDetailTab(candidate)) continue
    if (seen.has(candidate.id)) {
      if (import.meta.env.DEV) console.warn(`[item-detail] id de aba duplicado "${candidate.id}" em ${path}; ignorado.`)
      continue
    }
    seen.add(candidate.id)
    tabs.push(candidate)
  }
  return tabs
    .filter((tab) => !tab.roles || (role !== undefined && tab.roles.includes(role)))
    .sort((a, b) => a.order - b.order)
}

// Cada arquivo em ./tabs/ é uma aba; nenhum outro arquivo precisa ser editado
// ao acrescentar uma. Arquivos de teste ficam de fora.
const modules = import.meta.glob<{ default: ItemDetailTab }>(['./tabs/*.tsx', '!./tabs/*.test.tsx'], {
  eager: true,
})

export function getItemDetailTabs(role: string | undefined): ItemDetailTab[] {
  return buildItemDetailTabs(modules, role)
}
