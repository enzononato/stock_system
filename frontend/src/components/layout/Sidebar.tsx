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
    <aside className="flex flex-col w-64 min-h-screen bg-slate-950 text-slate-200 border-r border-slate-800/80 py-5 px-3.5 flex-shrink-0 justify-between select-none">
      <div className="space-y-6">
        {/* Brand Header */}
        <div className="flex items-center gap-3 px-2 py-1">
          <div className="h-10 w-10 rounded-xl bg-gradient-to-tr from-indigo-600 via-indigo-500 to-violet-500 flex items-center justify-center text-white shadow-lg shadow-indigo-500/30 ring-1 ring-white/20">
            <Boxes size={22} className="stroke-[2.2]" />
          </div>
          <div>
            <div className="flex items-center gap-1.5">
              <span className="text-base font-bold tracking-tight text-white font-heading">Revalle</span>
              <span className="text-[10px] font-extrabold uppercase px-1.5 py-0.5 rounded bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
                PRO
              </span>
            </div>
            <p className="text-xs text-slate-400 font-medium">Controle de Estoque TI</p>
          </div>
        </div>

        {/* Navigation Groups */}
        <nav className="space-y-5">
          {navGroups.map((group) => {
            const visibleItems = group.items.filter(
              (item) => !item.roles || (user && item.roles.includes(user.role))
            )
            if (visibleItems.length === 0) return null

            return (
              <div key={group.title} className="space-y-1">
                <h3 className="px-3 text-[11px] font-semibold text-slate-400 uppercase tracking-wider">
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
                          'group relative flex items-center gap-3 rounded-xl px-3 py-2 text-xs font-semibold transition-all duration-200',
                          isActive
                            ? 'bg-indigo-600/90 text-white shadow-md shadow-indigo-600/20 font-bold backdrop-blur-sm'
                            : 'text-slate-400 hover:bg-slate-900/80 hover:text-slate-100'
                        )
                      }
                    >
                      {({ isActive }) => (
                        <>
                          <span
                            className={cn(
                              'transition-transform duration-200 group-hover:scale-110',
                              isActive ? 'text-white' : 'text-slate-400 group-hover:text-indigo-400'
                            )}
                          >
                            {item.icon}
                          </span>
                          <span className="truncate">{item.label}</span>
                          {isActive && (
                            <span className="absolute left-0 top-1/2 -translate-y-1/2 h-5 w-1 rounded-r-full bg-white shadow-glow-indigo" />
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
      </div>

      {/* User Footer Summary */}
      {user && (
        <div className="pt-4 border-t border-slate-800/80 px-2">
          <div className="flex items-center gap-3 p-2.5 rounded-xl bg-slate-900/70 border border-slate-800/60">
            <div className="h-8 w-8 rounded-lg bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 flex items-center justify-center text-xs font-bold uppercase">
              {user.username.substring(0, 2)}
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-xs font-semibold text-white truncate">{user.username}</p>
              <div className="flex items-center gap-1 text-[10px] text-indigo-400 font-medium">
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

