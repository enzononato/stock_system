import { NavLink } from 'react-router-dom'
import { cn } from '@/lib/utils'
import { useAuth } from '@/contexts/AuthContext'
import {
  Package, PackagePlus, Cpu, Link2,
  ArrowRightLeft, Undo2, Trash2, History, BarChart2, LineChart,
  FileText, Users, Building2, Boxes, ShieldCheck
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
      { to: '/', label: 'Estoque', icon: <Package size={18} /> },
      { to: '/charts', label: 'Dashboard & Gráficos', icon: <LineChart size={18} /> },
    ],
  },
  {
    title: 'Gestão de Itens',
    items: [
      { to: '/register', label: 'Cadastrar Equipamento', icon: <PackagePlus size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/peripherals', label: 'Periféricos', icon: <Cpu size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/link', label: 'Vincular Periférico', icon: <Link2 size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/loan', label: 'Emprestar', icon: <ArrowRightLeft size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/return', label: 'Devolver', icon: <Undo2 size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/terms', label: 'Termos de Resp.', icon: <FileText size={18} />, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Relatórios & Auditoria',
    items: [
      { to: '/history', label: 'Histórico de Ações', icon: <History size={18} />, roles: ['Gestor', 'Técnico'] },
      { to: '/report', label: 'Relatórios BI', icon: <BarChart2 size={18} />, roles: ['Gestor', 'Técnico'] },
    ],
  },
  {
    title: 'Administração',
    items: [
      { to: '/remove', label: 'Remover / Estorno', icon: <Trash2 size={18} />, roles: ['Gestor'] },
      { to: '/unidades', label: 'Unidades de Revenda', icon: <Building2 size={18} />, roles: ['Gestor'] },
      { to: '/users', label: 'Gestão de Usuários', icon: <Users size={18} />, roles: ['Gestor'] },
    ],
  },
]

export default function Sidebar() {
  const { user } = useAuth()

  return (
    <aside className="flex flex-col w-64 h-full bg-sidebar border-r border-sidebar-border select-none shrink-0">
      {/* Brand Header */}
      <div className="flex items-center gap-2.5 h-14 px-3.5 border-b border-border shrink-0">
        <div className="h-8 w-8 rounded-md bg-primary flex items-center justify-center text-primary-foreground shrink-0">
          <Boxes size={16} />
        </div>
        <div className="min-w-0">
          <p className="text-sm font-semibold tracking-tight truncate">Revalle</p>
          <p className="text-[11px] text-muted-foreground truncate">Controle de Estoque TI</p>
        </div>
      </div>

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
