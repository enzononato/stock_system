import { useCallback, useRef } from 'react'
import { NavLink } from 'react-router-dom'
import { cn } from '@/lib/utils'
import { useAuth } from '@/contexts/AuthContext'
import {
  Package, PackagePlus, Cpu, Link2,
  ArrowRightLeft, Undo2, Trash2, History, BarChart2, LineChart,
  FileText, Users, Building2, Boxes, PanelLeftClose, PanelLeftOpen
} from 'lucide-react'

interface NavItem {
  to: string
  label: string
  icon: React.ReactNode
  roles?: string[]
}

interface NavGroup {
  title: string
  items: NavItem[]
}

const navGroups: NavGroup[] = [
  {
    title: 'Visão Geral',
    items: [
      { to: '/', label: 'Estoque', icon: <Package size={17} /> },
      { to: '/charts', label: 'Dashboard & Gráficos', icon: <LineChart size={17} /> },
    ],
  },
  {
    title: 'Gestão de Itens',
    items: [
      { to: '/register', label: 'Cadastrar Equipamento', icon: <PackagePlus size={17} />, roles: ['Gestor', 'Técnico'] },
      { to: '/peripherals', label: 'Periféricos', icon: <Cpu size={17} />, roles: ['Gestor', 'Técnico'] },
      { to: '/link', label: 'Vincular Periférico', icon: <Link2 size={17} />, roles: ['Gestor', 'Técnico'] },
      { to: '/loan', label: 'Emprestar', icon: <ArrowRightLeft size={17} />, roles: ['Gestor', 'Técnico'] },
      { to: '/return', label: 'Devolver', icon: <Undo2 size={17} />, roles: ['Gestor', 'Técnico'] },
      { to: '/terms', label: 'Termos de Resp.', icon: <FileText size={17} />, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Relatórios & Auditoria',
    items: [
      { to: '/history', label: 'Histórico de Ações', icon: <History size={17} />, roles: ['Gestor', 'Técnico'] },
      { to: '/report', label: 'Relatórios BI', icon: <BarChart2 size={17} />, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Administração',
    items: [
      { to: '/remove', label: 'Remover / Estorno', icon: <Trash2 size={17} />, roles: ['Gestor'] },
      { to: '/unidades', label: 'Unidades de Revenda', icon: <Building2 size={17} />, roles: ['Gestor'] },
      { to: '/users', label: 'Gestão de Usuários', icon: <Users size={17} />, roles: ['Gestor'] },
    ],
  },
]

interface SidebarProps {
  collapsed: boolean
  onToggle: () => void
}

const PROXIMITY_RADIUS = 70

export default function Sidebar({ collapsed, onToggle }: SidebarProps) {
  const { user } = useAuth()
  const navRef = useRef<HTMLElement>(null)

  // Proximity effect adapted from React Bits' LineSidebar: nav items shift
  // toward the brand color as the cursor gets vertically close to them.
  const handlePointerMove = useCallback((e: React.PointerEvent<HTMLElement>) => {
    const nav = navRef.current
    if (!nav) return
    const items = nav.querySelectorAll<HTMLElement>('[data-nav-item]')
    const rect = nav.getBoundingClientRect()
    const pointerY = e.clientY - rect.top
    items.forEach((el) => {
      const center = el.offsetTop + el.offsetHeight / 2
      const distance = Math.abs(pointerY - center)
      const proximity = Math.max(0, 1 - distance / PROXIMITY_RADIUS)
      el.style.setProperty('--proximity', proximity.toFixed(3))
    })
  }, [])

  const handlePointerLeave = useCallback(() => {
    const nav = navRef.current
    if (!nav) return
    nav.querySelectorAll<HTMLElement>('[data-nav-item]').forEach((el) => el.style.setProperty('--proximity', '0'))
  }, [])

  return (
    <aside
      className={cn(
        'hidden lg:flex flex-col min-h-screen bg-sidebar text-sidebar-foreground border-r border-sidebar-border flex-shrink-0 select-none transition-[width] duration-200',
        collapsed ? 'w-[76px]' : 'w-64'
      )}
    >
      {/* Brand Header */}
      <div
        className={cn(
          'flex items-center h-[64px] border-b border-sidebar-border/80',
          collapsed ? 'justify-center px-0' : 'gap-3 px-4'
        )}
      >
        <div className="h-8 w-8 shrink-0 rounded-md bg-primary text-primary-foreground flex items-center justify-center">
          <Boxes size={18} className="stroke-[2.2]" />
        </div>
        {!collapsed && (
          <div className="min-w-0">
            <p className="text-sm font-semibold tracking-tight text-foreground truncate">Revalle</p>
            <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Controle de Estoque TI</p>
          </div>
        )}
      </div>

      {/* Navigation Groups */}
      <nav
        ref={navRef}
        onPointerMove={handlePointerMove}
        onPointerLeave={handlePointerLeave}
        className="flex-1 overflow-y-auto py-5 px-3 space-y-6"
      >
        {navGroups.map((group) => {
          const visibleItems = group.items.filter(
            (item) => !item.roles || (user && item.roles.includes(user.role))
          )
          if (visibleItems.length === 0) return null

          return (
            <div key={group.title} className="space-y-1">
              <p
                className={cn(
                  'px-3 text-[10px] font-semibold uppercase tracking-[0.12em] text-muted-foreground/70',
                  collapsed && 'sr-only'
                )}
              >
                {group.title}
              </p>
              <div className="space-y-0.5">
                {visibleItems.map((item) => (
                  <NavLink
                    key={item.to}
                    data-nav-item
                    to={item.to}
                    end={item.to === '/'}
                    title={collapsed ? item.label : undefined}
                    style={{
                      transform: 'translateX(calc(var(--proximity, 0) * 3px))',
                    }}
                    className={({ isActive }) =>
                      cn(
                        'group flex items-center gap-3 rounded-md px-3 h-9 text-[13px] font-medium transition-[background-color,color] duration-150',
                        collapsed && 'justify-center px-0',
                        isActive
                          ? 'bg-sidebar-primary/15 text-sidebar-primary'
                          : 'text-sidebar-foreground/70 hover:bg-sidebar-accent hover:text-sidebar-accent-foreground'
                      )
                    }
                  >
                    {({ isActive }) => (
                      <>
                        <span className={isActive ? 'text-sidebar-primary' : 'text-sidebar-foreground/60'}>
                          {item.icon}
                        </span>
                        {!collapsed && <span className="truncate">{item.label}</span>}
                        {isActive && !collapsed && (
                          <span className="ml-auto h-1.5 w-1.5 rounded-full bg-sidebar-primary" />
                        )}
                      </>
                    )}
                  </NavLink>
                ))}
              </div>
            </div>
          )
        })}
      </nav>

      {/* Collapse toggle */}
      <div className="border-t border-sidebar-border/80 p-2">
        <button
          type="button"
          onClick={onToggle}
          aria-label={collapsed ? 'Expandir menu' : 'Recolher menu'}
          className={cn(
            'flex items-center gap-3 w-full rounded-md px-3 h-9 text-[13px] font-medium text-sidebar-foreground/60 hover:bg-sidebar-accent hover:text-sidebar-accent-foreground transition-colors',
            collapsed && 'justify-center px-0'
          )}
        >
          {collapsed ? <PanelLeftOpen size={17} /> : <PanelLeftClose size={17} />}
          {!collapsed && <span>Recolher menu</span>}
        </button>
      </div>

      {/* User Footer Summary */}
      {user && (
        <div className={cn('border-t border-sidebar-border/80 p-3', collapsed && 'px-2')}>
          <div
            className={cn(
              'flex items-center gap-3 rounded-md p-2 hover:bg-sidebar-accent transition-colors',
              collapsed && 'justify-center'
            )}
          >
            <div className="h-8 w-8 shrink-0 rounded-full bg-gradient-to-br from-sky-400 to-blue-700 flex items-center justify-center text-[11px] font-bold text-white">
              {user.username.substring(0, 2).toUpperCase()}
            </div>
            {!collapsed && (
              <div className="min-w-0 flex-1">
                <p className="text-xs font-medium text-foreground truncate">{user.username}</p>
                <p className="text-[11px] text-muted-foreground truncate">{user.role}</p>
              </div>
            )}
          </div>
        </div>
      )}
    </aside>
  )
}
