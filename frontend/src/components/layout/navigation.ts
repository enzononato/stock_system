import type { LucideIcon } from 'lucide-react'
import {
  LayoutDashboard, Package, PackagePlus, Cpu, Link2,
  ArrowRightLeft, Undo2, Trash2, History, BarChart2, LineChart,
  FileText, Users, Building2,
} from 'lucide-react'

export interface NavItem {
  to: string
  label: string
  icon: LucideIcon
  roles?: string[]
}

export interface NavGroup {
  title: string
  items: NavItem[]
}

/**
 * Estrutura ÚNICA de navegação do app. Sidebar.tsx e CommandPalette.tsx (e a
 * trilha de breadcrumbs em Breadcrumb.tsx) consomem esta mesma lista — nunca
 * duplicá-la, senão um destino novo (ou uma mudança de papel) precisaria ser
 * editado em dois lugares e um dos dois acabaria ficando desatualizado.
 *
 * Diferente da referência (`origin/redesign-frontend`), aqui `/register`,
 * `/peripherals`, `/link`, `/loan`, `/return` e `/terms` continuam TODOS
 * visíveis na sidebar — a referência esconde `/remove`, `/terms` e `/link`
 * dela e só deixa esses alcançáveis via Ctrl+K. Optamos por não escondê-los
 * (ver relatório da Task 8, seção "Remoções necessárias para ficar
 * idêntico"), então a paleta de comandos oferece exatamente os mesmos itens
 * que já aparecem na sidebar.
 */
export const navGroups: NavGroup[] = [
  {
    title: 'Visão Geral',
    items: [
      // Task 7: Dashboard assume "/" (nova home); o estoque, que ficava
      // aqui, passou para "/stock".
      { to: '/', label: 'Dashboard', icon: LayoutDashboard },
      { to: '/stock', label: 'Estoque', icon: Package },
      { to: '/charts', label: 'Indicadores & Gráficos', icon: LineChart },
    ],
  },
  {
    title: 'Gestão de Itens',
    items: [
      { to: '/register', label: 'Cadastrar Equipamento', icon: PackagePlus, roles: ['Gestor', 'Técnico'] },
      { to: '/peripherals', label: 'Periféricos', icon: Cpu, roles: ['Gestor', 'Técnico'] },
      { to: '/link', label: 'Vincular Periférico', icon: Link2, roles: ['Gestor', 'Técnico'] },
      { to: '/loan', label: 'Emprestar', icon: ArrowRightLeft, roles: ['Gestor', 'Técnico'] },
      { to: '/return', label: 'Devolver', icon: Undo2, roles: ['Gestor', 'Técnico'] },
      { to: '/terms', label: 'Termos de Resp.', icon: FileText, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Relatórios & Auditoria',
    items: [
      { to: '/history', label: 'Histórico de Ações', icon: History, roles: ['Gestor', 'Técnico'] },
      { to: '/report', label: 'Relatórios BI', icon: BarChart2, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Administração',
    items: [
      { to: '/remove', label: 'Remover / Estorno', icon: Trash2, roles: ['Gestor'] },
      { to: '/unidades', label: 'Unidades de Revenda', icon: Building2, roles: ['Gestor'] },
      { to: '/users', label: 'Gestão de Usuários', icon: Users, roles: ['Gestor'] },
    ],
  },
]

/** Um item é visível quando não exige papel algum, ou quando o papel do usuário está na lista. */
function isVisible(item: NavItem, role: string | undefined): boolean {
  return !item.roles || (!!role && item.roles.includes(role))
}

/** Grupos de navegação já filtrados pelo papel do usuário, com grupos vazios removidos. */
export function getVisibleNavGroups(role: string | undefined): NavGroup[] {
  return navGroups
    .map((group) => ({ ...group, items: group.items.filter((item) => isVisible(item, role)) }))
    .filter((group) => group.items.length > 0)
}

/** Todos os itens visíveis para o papel, em lista única (usado pela paleta de comandos). */
export function getVisibleNavItems(role: string | undefined): NavItem[] {
  return getVisibleNavGroups(role).flatMap((group) => group.items)
}

export interface FlatNavEntry extends NavItem {
  groupTitle: string
}

/** Todos os itens da navegação (sem filtro de papel), cada um com o título do grupo a que pertence — usado pelos breadcrumbs. */
export function getFlatNavEntries(): FlatNavEntry[] {
  return navGroups.flatMap((group) => group.items.map((item) => ({ ...item, groupTitle: group.title })))
}
