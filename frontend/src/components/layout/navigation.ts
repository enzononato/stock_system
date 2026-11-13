import type { ComponentType } from 'react'
import { Contact, MapPin, ScrollText, type LucideIcon } from 'lucide-react'
import { Smartphone } from 'lucide-react'
import {
  DashboardIcon,
  EstoqueIcon,
  PatrimoniosIcon,
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
 * strokeWidth/title/aria-hidden) também o atendem. Atenção: os animados
 * renderizam `role="img" aria-label={title}` e IGNORAM `aria-hidden`; quem os
 * usa ao lado de um rótulo (Sidebar, CommandPalette) os envolve num
 * `<span aria-hidden="true">` para o nome acessível não duplicar. `NavIcon` é a união dos
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
}

export interface NavGroup {
  title: string
  items: NavItem[]
}

/**
 * Estrutura ÚNICA de navegação do app. Sidebar.tsx e CommandPalette.tsx
 * consomem esta mesma lista — nunca duplicá-la, senão um destino novo (ou uma
 * mudança de papel) precisaria ser editado em dois lugares e um dos dois
 * acabaria ficando desatualizado.
 *
 * Grupos, rótulos, ícones e ordem iguais aos da branch `origin/FT_STC`, a
 * referência visual escolhida pelo dono: todas as telas aparecem no menu.
 */
export const navGroups: NavGroup[] = [
  {
    title: 'Principal',
    items: [
      { to: '/', label: 'Dashboard', icon: DashboardIcon },
    ],
  },
  {
    title: 'Inventário',
    items: [
      { to: '/stock', label: 'Estoque', icon: EstoqueIcon },
      { to: '/register', label: 'Cadastro', icon: PatrimoniosIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/history', label: 'Histórico', icon: MovimentacoesIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/locais', label: 'Locais', icon: MapPin, roles: ['Gestor', 'Técnico'] },
      { to: '/linhas', label: 'Linhas', icon: Smartphone },
    ],
  },
  {
    title: 'Operações',
    items: [
      { to: '/loan', label: 'Empréstimos', icon: EmprestimosIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/return', label: 'Devoluções', icon: DevolucoesIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/peripherals', label: 'Periféricos', icon: PerifericosIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/link', label: 'Vincular Periférico', icon: MovimentacoesIcon, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Gestão',
    items: [
      { to: '/users', label: 'Usuários', icon: UsuariosIcon, roles: ['Gestor'] },
      { to: '/unidades', label: 'Unidades', icon: EmpresasIcon, roles: ['Gestor'] },
      { to: '/colaboradores', label: 'Colaboradores', icon: Contact, roles: ['Gestor', 'Técnico'] },
      { to: '/terms', label: 'Termos de Resp.', icon: TermosIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/report', label: 'Relatórios', icon: RelatoriosIcon, roles: ['Gestor', 'Técnico'] },
      { to: '/charts', label: 'Indicadores', icon: DashboardIcon },
      { to: '/remove', label: 'Baixa de Ativos', icon: LixeiraIcon, roles: ['Gestor'] },
      { to: '/auditoria', label: 'Auditoria', icon: ScrollText, roles: ['Gestor'] },
    ],
  },
]

/** Um item é visível quando não exige papel algum, ou quando o papel do usuário está na lista. */
function isVisible(item: NavItem, role: string | undefined): boolean {
  return !item.roles || (!!role && item.roles.includes(role))
}

/** Grupos de navegação já filtrados pelo papel do usuário, com grupos vazios removidos — usado pela Sidebar. */
export function getVisibleNavGroups(role: string | undefined): NavGroup[] {
  return navGroups
    .map((group) => ({ ...group, items: group.items.filter((item) => isVisible(item, role)) }))
    .filter((group) => group.items.length > 0)
}

/** Todos os itens visíveis para o papel, em lista única — usado pela paleta de comandos (Ctrl+K). */
export function getVisibleNavItems(role: string | undefined): NavItem[] {
  return navGroups.flatMap((group) => group.items).filter((item) => isVisible(item, role))
}
