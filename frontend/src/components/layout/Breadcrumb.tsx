import { useLocation } from 'react-router-dom'
import { ChevronRight } from 'lucide-react'
import { getFlatNavEntries } from './navigation'

/**
 * Sub-rotas que não têm uma entrada própria na estrutura de navegação (ex.:
 * edição de um equipamento específico por id) e por isso precisam de um
 * rótulo manual aqui — o restante do sistema é 100% derivado de
 * `navigation.ts`.
 */
const EXTRA_CRUMBS: { test: (pathname: string) => boolean; groupTitle: string; label: string }[] = [
  { test: (pathname) => pathname.startsWith('/edit/'), groupTitle: 'Gestão de Itens', label: 'Editar Equipamento' },
]

/**
 * Trilha "Grupo › Página" exibida na TopBar, derivada da MESMA estrutura de
 * navegação usada pela sidebar e pela paleta de comandos (Task 8, Step 3).
 * O título da página em si continua vivendo no `PageHeader` de cada tela —
 * isto aqui é só o contexto de localização, não repete o título.
 */
export function BreadcrumbTrail() {
  const location = useLocation()
  const pathname = location.pathname

  const extra = EXTRA_CRUMBS.find((crumb) => crumb.test(pathname))
  const match = !extra
    ? getFlatNavEntries().find((item) => (item.to === '/' ? pathname === '/' : pathname.startsWith(item.to)))
    : undefined

  const groupTitle = extra?.groupTitle ?? match?.groupTitle ?? 'Sistema'
  const label = extra?.label ?? match?.label ?? 'Visão Geral'

  return (
    <nav aria-label="Trilha de navegação" className="flex items-center gap-1.5 text-body-sm min-w-0">
      <span className="text-muted-foreground truncate">{groupTitle}</span>
      <ChevronRight size={14} className="text-muted-foreground shrink-0" aria-hidden="true" />
      <span className="font-semibold text-foreground tracking-tight truncate">{label}</span>
    </nav>
  )
}
