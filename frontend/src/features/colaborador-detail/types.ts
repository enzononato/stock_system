import type { ComponentType } from 'react'
import type { Colaborador } from '@/api/colaboradores'

/**
 * Uma aba da página de detalhe do colaborador. Cada fase acrescenta a sua
 * criando UM arquivo em `features/colaborador-detail/tabs/` com um
 * `export default` deste tipo — o registro (`registry.ts`) descobre os
 * arquivos sozinho.
 *
 * Ordens reservadas: dados 10, equipamentos 20, termos 30 (F3), linhas 40 (F5),
 * acessos 50 (F8).
 */
export interface ColaboradorDetailTab {
  /** Identificador estável, usado em `?aba=<id>`. */
  id: string
  label: string
  order: number
  /** Quando informado, só esses papéis veem a aba. */
  roles?: string[]
  Component: ComponentType<{ colaborador: Colaborador }>
}
