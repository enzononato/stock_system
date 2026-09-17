import { useState, useEffect } from 'react'
import { NavLink, Link } from 'react-router-dom'
import { cn } from '@/lib/utils'
import { useAuth } from '@/contexts/AuthContext'
import { ShieldCheck, ChevronRight } from 'lucide-react'
import {
  EstoqueIcon,
  DashboardIcon,
  PatrimoniosIcon,
  PerifericosIcon,
  MovimentacoesIcon,
  EmprestimosIcon,
  DevolucoesIcon,
  RelatoriosIcon,
  TermosIcon,
  LixeiraIcon,
  EmpresasIcon,
  UsuariosIcon,
} from '@/components/animated-icons/stock-system-animated-sidebar-icons'

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
    title: 'Principal',
    items: [
      { to: '/', label: 'Dashboard', icon: <DashboardIcon size={18} /> },
    ],
  },
  {
    title: 'Inventário',
    items: [
      { to: '/stock', label: 'Estoque', icon: <EstoqueIcon size={18} /> },
      { to: '/register', label: 'Cadastro', icon: <PatrimoniosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/history', label: 'Histórico', icon: <MovimentacoesIcon size={18} />, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Operações',
    items: [
      { to: '/loan', label: 'Empréstimos', icon: <EmprestimosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/return', label: 'Devoluções', icon: <DevolucoesIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/peripherals', label: 'Periféricos', icon: <PerifericosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/link', label: 'Vincular Periférico', icon: <MovimentacoesIcon size={18} />, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Gestão',
    items: [
      { to: '/users', label: 'Usuários', icon: <UsuariosIcon size={18} />, roles: ['Gestor'] },
      { to: '/unidades', label: 'Unidades', icon: <EmpresasIcon size={18} />, roles: ['Gestor'] },
      { to: '/terms', label: 'Termos de Resp.', icon: <TermosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/report', label: 'Relatórios', icon: <RelatoriosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/charts', label: 'Indicadores', icon: <DashboardIcon size={18} /> },
      { to: '/remove', label: 'Baixa de Ativos', icon: <LixeiraIcon size={18} />, roles: ['Gestor'] },
    ],
  },
]

const STORAGE_KEY = 'revalle_sidebar_groups'

export default function Sidebar() {
  const { user } = useAuth()

  const [openGroups, setOpenGroups] = useState<Record<string, boolean>>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY)
      if (saved) return JSON.parse(saved)
    } catch {
      // Ignora erro de storage
    }
    return {
      'Principal': true,
      'Inventário': true,
      'Operações': true,
      'Gestão': true,
    }
  })

  function toggleGroup(title: string) {
    setOpenGroups((prev) => {
      const updated = { ...prev, [title]: !prev[title] }
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(updated))
      } catch {
        // Ignora erro de storage
      }
      return updated
    })
  }

  return (
    <aside className="flex flex-col w-64 h-full bg-sidebar border-r border-sidebar-border select-none shrink-0">
      {/* Brand Header */}
      <Link
        to="/"
        className="flex flex-col items-start justify-center py-5 px-4 border-b border-border shrink-0 gap-2.5 transition-colors hover:bg-surface-alt group"
        title="Início - Revalle Controle de Patrimônio"
      >
        <div className="w-full flex items-center justify-start overflow-hidden">
          <img
            src="/logo-revalle.png"
            alt="Revalle"
            className="h-14 w-auto max-w-[205px] object-contain filter brightness-0 dark:brightness-100 drop-shadow-md transition-transform duration-200 group-hover:scale-[1.02]"
          />
        </div>
        <div className="flex items-center gap-1.5 px-2 py-0.5 rounded bg-surface-alt border border-border">
          <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
          <span className="text-[10px] font-semibold text-foreground tracking-widest uppercase font-mono">
            Controle de Patrimônio
          </span>
        </div>
      </Link>

      {/* Navigation Groups */}
      <nav className="flex-1 overflow-y-auto py-4 px-2.5 space-y-4">
        {navGroups.map((group) => {
          const visibleItems = group.items.filter(
            (item) => !item.roles || (user && item.roles.includes(user.role))
          )
          if (visibleItems.length === 0) return null
          const isOpen = openGroups[group.title] ?? true

          return (
            <div key={group.title} className="space-y-1">
              <button
                type="button"
                onClick={() => toggleGroup(group.title)}
                className="flex w-full items-center justify-between px-2.5 py-1 text-caption text-muted-foreground hover:text-foreground transition-colors cursor-pointer group/hdr"
              >
                <span className="font-semibold tracking-wider">{group.title}</span>
                <ChevronRight
                  className={cn(
                    'size-3.5 text-muted-foreground transition-transform duration-micro group-hover/hdr:text-foreground',
                    isOpen && 'rotate-90'
                  )}
                  aria-hidden="true"
                />
              </button>

              {isOpen && (
                <div className="space-y-0.5 animate-accordion-down">
                  {visibleItems.map((item) => (
                    <NavLink
                      key={item.to}
                      to={item.to}
                      end={item.to === '/'}
                      className={({ isActive }) =>
                        cn(
                          'flex items-center gap-2.5 px-2.5 py-1.5 rounded text-body-sm transition-colors duration-micro',
                          isActive
                            ? 'bg-surface-alt text-foreground font-semibold border-l-2 border-foreground'
                            : 'text-muted-foreground hover:text-foreground hover:bg-surface-alt'
                        )
                      }
                    >
                      {item.icon}
                      <span className="truncate">{item.label}</span>
                    </NavLink>
                  ))}
                </div>
              )}
            </div>
          )
        })}
      </nav>

      {/* User Footer Summary */}
      {user && (
        <div className="border-t border-border p-3 shrink-0">
          <div className="flex items-center gap-2.5">
            <div className="h-8 w-8 rounded bg-surface-alt text-foreground border border-border-strong flex items-center justify-center text-xs font-semibold uppercase shrink-0">
              {user.username.substring(0, 2)}
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-body-sm font-medium truncate">{user.username}</p>
              <div className="flex items-center gap-1 text-[11px] text-muted-foreground font-mono">
                <ShieldCheck size={11} />
                <span>{user.role}</span>
              </div>
            </div>
          </div>
        </div>
      )}
    </aside>
  )
}
