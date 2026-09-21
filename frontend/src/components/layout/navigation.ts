import type { ComponentType } from 'react'
import type { LucideIcon } from 'lucide-react'
import { PackagePlus, Link2, LineChart } from 'lucide-react'
import {
  DashboardIcon,
  EstoqueIcon,
  PerifericosIcon,
  EmprestimosIcon,
  DevolucoesIcon,
  TermosIcon,
  MovimentacoesIcon,
  RelatoriosIcon,
  LixeiraIcon,
  EmpresasIcon,
  UsuariosIcon,
} from '@/components/animated-icons/stock-system-animated-sidebar-icons'

/**
 * Props mínimas que qualquer ícone de navegação precisa aceitar. Os ícones do
 * lucide-react atendem este contrato de sobra (aceitam qualquer atributo SVG);
 * os ícones animados da sidebar (API mais restrita: size/className/
 * strokeWidth/title/aria-hidden) também o atendem. `NavIcon` é a união dos
 * dois — não é um cast, é o tipo real de cada ícone usado em `navGroups`.
 */
export interface NavIconProps {
  size?: number | string
  className?: string
  strokeWidth?: number
  title?: string
  'aria-hidden'?: boolean | 'true' | 'false'
}

export type NavIcon = LucideIcon | ComponentType<NavIconProps>

export interface NavItem {
  to: string
  label: string
  icon: NavIcon
  roles?: string[]
  /**
   * Esconde o item da sidebar (igual à referência `origin/redesign-frontend`),
   * mas mantém o destino disponível na paleta de comandos (Ctrl+K) e nos
   * breadcrumbs — só `getVisibleNavGroups` (consumido pela Sidebar) respeita
   * esta flag; `getVisibleNavItems` (paleta) e `getFlatNavEntries`
   * (breadcrumbs) continuam enxergando o item normalmente.
   */
  hiddenFromSidebar?: boolean
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
 * Igual à referência (`origin/redesign-frontend`): `/remove`, `/terms` e
 * `/link` ficam marcados com `hiddenFromSidebar` e não aparecem no menu — só
 * são alcançáveis via paleta de comandos (Ctrl+K), que ignora essa flag. As
 * rotas continuam registradas em `App.tsx` normalmente.
 */
export const navGroups: NavGroup[] = [
  {
    title: 'Visão Geral',
    items: [
      // Task 7: Dashboard assume "/" (nova home); o estoque, que ficava
      // aqui, passou para "/stock".
      { to: '/', label: 'Dashboard', icon: DashboardIcon },
      { to: '/stock', label: 'Estoque', icon: EstoqueIcon },
      { to: '/charts', label: 'Indicadores & Gráficos', icon: LineChart },
    ],
  },
  {
    title: 'Gestão de Itens',
    items: [
      { to: '/register', label: 'Cadastrar Equipamento', icon: PackagePlus, roles: ['Gestor', 'Técnico'] },
      { to: '/peripherals', label: 'Periféricos', icon: PerifericosIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/link', label: 'Vincular Periférico', icon: Link2, roles: ['Gestor', 'Técnico'], hiddenFromSidebar: true },
      { to: '/loan', label: 'Emprestar', icon: EmprestimosIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/return', label: 'Devolver', icon: DevolucoesIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/terms', label: 'Termos de Resp.', icon: TermosIcon, roles: ['Gestor', 'Técnico'], hiddenFromSidebar: true },
    ],
  },
  {
    title: 'Relatórios & Auditoria',
    items: [
      { to: '/history', label: 'Histórico de Ações', icon: MovimentacoesIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/report', label: 'Relatórios BI', icon: RelatoriosIcon, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Administração',
    items: [
      { to: '/remove', label: 'Remover / Estorno', icon: LixeiraIcon, roles: ['Gestor'], hiddenFromSidebar: true },
      { to: '/unidades', label: 'Unidades de Revenda', icon: EmpresasIcon, roles: ['Gestor'] },
      { to: '/users', label: 'Gestão de Usuários', icon: UsuariosIcon, roles: ['Gestor'] },
    ],
  },
]

/** Um item é visível quando não exige papel algum, ou quando o papel do usuário está na lista. */
function isVisible(item: NavItem, role: string | undefined): boolean {
  return !item.roles || (!!role && item.roles.includes(role))
}

/**
 * Grupos de navegação já filtrados pelo papel do usuário, com grupos vazios
 * removidos — usado pela Sidebar. Também remove os itens com
 * `hiddenFromSidebar` (ver comentário no tipo `NavItem`): a paleta de
 * comandos usa `getVisibleNavItems`, não esta função, então continua
 * enxergando esses destinos.
 */
export function getVisibleNavGroups(role: string | undefined): NavGroup[] {
  return navGroups
    .map((group) => ({
      ...group,
      items: group.items.filter((item) => isVisible(item, role) && !item.hiddenFromSidebar),
    }))
    .filter((group) => group.items.length > 0)
}

/**
 * Todos os itens visíveis para o papel, em lista única (usado pela paleta de
 * comandos) — filtra só por papel, ignorando `hiddenFromSidebar` de propósito:
 * um item escondido da sidebar precisa continuar alcançável por Ctrl+K.
 */
export function getVisibleNavItems(role: string | undefined): NavItem[] {
  return navGroups.flatMap((group) => group.items).filter((item) => isVisible(item, role))
}

export interface FlatNavEntry extends NavItem {
  groupTitle: string
}

/** Todos os itens da navegação (sem filtro de papel), cada um com o título do grupo a que pertence — usado pelos breadcrumbs. */
export function getFlatNavEntries(): FlatNavEntry[] {
  return navGroups.flatMap((group) => group.items.map((item) => ({ ...item, groupTitle: group.title })))
}
