import { NavLink, Link } from 'react-router-dom'
import { cn } from '@/lib/utils'
import { useAuth } from '@/contexts/AuthContext'
import { ShieldCheck } from 'lucide-react'
import {
  EstoqueIcon,
  DashboardIcon,
  PatrimoniosIcon,
  PerifericosIcon,
  MovimentacoesIcon,
  EmprestimosIcon,
  DevolucoesIcon,
  RelatoriosIcon,
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
    title: 'Visão Geral',
    items: [
      { to: '/', label: 'Estoque', icon: <EstoqueIcon size={18} /> },
      { to: '/charts', label: 'Dashboard & Gráficos', icon: <DashboardIcon size={18} /> },
    ],
  },
  {
    title: 'Gestão de Itens',
    items: [
      { to: '/register', label: 'Cadastrar Equipamento', icon: <PatrimoniosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/peripherals', label: 'Periféricos', icon: <PerifericosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/link', label: 'Vincular Periférico', icon: <MovimentacoesIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/loan', label: 'Emprestar', icon: <EmprestimosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/return', label: 'Devolver', icon: <DevolucoesIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/terms', label: 'Termos de Resp.', icon: <RelatoriosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Relatórios & Auditoria',
    items: [
      { to: '/history', label: 'Histórico de Ações', icon: <MovimentacoesIcon size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/report', label: 'Relatórios BI', icon: <RelatoriosIcon size={18} />, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Administração',
    items: [
      { to: '/remove', label: 'Remover / Estorno', icon: <LixeiraIcon size={18} />, roles: ['Gestor'] },
      { to: '/unidades', label: 'Unidades de Revenda', icon: <EmpresasIcon size={18} />, roles: ['Gestor'] },
      { to: '/users', label: 'Gestão de Usuários', icon: <UsuariosIcon size={18} />, roles: ['Gestor'] },
    ],
  },
]

export default function Sidebar() {
  const { user } = useAuth()

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
      <nav className="flex-1 overflow-y-auto py-4 px-2.5 space-y-5">
        {navGroups.map((group) => {
          const visibleItems = group.items.filter(
            (item) => !item.roles || (user && item.roles.includes(user.role))
          )
          if (visibleItems.length === 0) return null

          return (
            <div key={group.title} className="space-y-0.5">
              <h3 className="text-caption text-muted-foreground px-2 py-1">
                {group.title}
              </h3>
              <div className="space-y-0.5">
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
