import type { ComponentType } from 'react'

/** Forma mínima comum às abas plugáveis das páginas de detalhe. */
export interface RegistryTab {
  id: string
  label: string
  order: number
  roles?: string[]
  Component: ComponentType<any>
}

export type TabModules = Record<string, { default?: unknown } | undefined>

function isRegistryTab(v: unknown): v is RegistryTab {
  if (!v || typeof v !== 'object') return false
  const t = v as Partial<RegistryTab>
  return (
    typeof t.id === 'string' &&
    t.id !== '' &&
    typeof t.label === 'string' &&
    typeof t.order === 'number' &&
    t.Component != null
  )
}

/**
 * Núcleo dos registros de abas (equipamento, colaborador...): filtra por papel
 * e ordena por `order`. Separado do `import.meta.glob` para ser testável.
 * Módulos sem um `default` válido (helpers, arquivos soltos em `tabs/`) são
 * ignorados; ids repetidos geram aviso em desenvolvimento (prefixado por
 * `scope`) e vale o primeiro, na ordem dos caminhos.
 */
export function buildTabRegistry<T extends RegistryTab>(
  modules: TabModules,
  role: string | undefined,
  scope: string
): T[] {
  const seen = new Set<string>()
  const tabs: T[] = []
  for (const path of Object.keys(modules).sort()) {
    const candidate = modules[path]?.default
    if (!isRegistryTab(candidate)) continue
    if (seen.has(candidate.id)) {
      if (import.meta.env.DEV) console.warn(`[${scope}] id de aba duplicado "${candidate.id}" em ${path}; ignorado.`)
      continue
    }
    seen.add(candidate.id)
    tabs.push(candidate as T)
  }
  return tabs
    .filter((tab) => !tab.roles || (role !== undefined && tab.roles.includes(role)))
    .sort((a, b) => a.order - b.order)
}
