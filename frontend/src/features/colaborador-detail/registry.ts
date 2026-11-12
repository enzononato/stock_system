import { buildTabRegistry, type TabModules } from '@/lib/tabRegistry'
import type { ColaboradorDetailTab } from './types'

/**
 * Filtra por papel e ordena por `order` (ver `buildTabRegistry`). Separada do
 * `import.meta.glob` para ser testável: módulos sem `default` válido são
 * ignorados e ids repetidos geram aviso em desenvolvimento.
 */
export function buildColaboradorDetailTabs(modules: TabModules, role: string | undefined): ColaboradorDetailTab[] {
  return buildTabRegistry<ColaboradorDetailTab>(modules, role, 'colaborador-detail')
}

// Cada arquivo em ./tabs/ é uma aba; nenhum outro arquivo precisa ser editado
// ao acrescentar uma. Arquivos de teste ficam de fora.
const modules = import.meta.glob<{ default: ColaboradorDetailTab }>(['./tabs/*.tsx', '!./tabs/*.test.tsx'], {
  eager: true,
})

export function getColaboradorDetailTabs(role: string | undefined): ColaboradorDetailTab[] {
  return buildColaboradorDetailTabs(modules, role)
}
