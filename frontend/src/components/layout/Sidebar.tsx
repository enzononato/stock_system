import { NavLink } from 'react-router-dom'
import { cn } from '@/lib/utils'
import { useAuth } from '@/contexts/AuthContext'
import {
  LayoutDashboard, Package, PackagePlus, Pencil, Cpu, Link2,
  ArrowRightLeft, Undo2, Trash2, History, BarChart2, LineChart,
  FileText, Users,
} from 'lucide-react'

interface NavItem {
  to: string
  label: string
  icon: React.ReactNode
  roles?: string[]
}

const navItems: NavItem[] = [
  { to: '/', label: 'Estoque', icon: <Package size={18} /> },
  { to: '/register', label: 'Cadastrar', icon: <PackagePlus size={18} />, roles: ['Gestor', 'Técnico'] },
  { to: '/peripherals', label: 'Periféricos', icon: <Cpu size={18} />, roles: ['Gestor', 'Técnico'] },
  { to: '/link', label: 'Vincular', icon: <Link2 size={18} />, roles: ['Gestor', 'Técnico'] },
  { to: '/loan', label: 'Emprestar', icon: <ArrowRightLeft size={18} />, roles: ['Gestor', 'Técnico'] },
  { to: '/return', label: 'Devolver', icon: <Undo2 size={18} />, roles: ['Gestor', 'Técnico'] },
  { to: '/remove', label: 'Remover', icon: <Trash2 size={18} />, roles: ['Gestor'] },
  { to: '/history', label: 'Histórico', icon: <History size={18} />, roles: ['Gestor', 'Técnico'] },
  { to: '/report', label: 'Relatório', icon: <BarChart2 size={18} />, roles: ['Gestor', 'Técnico'] },
  { to: '/charts', label: 'Gráficos', icon: <LineChart size={18} /> },
  { to: '/terms', label: 'Termos', icon: <FileText size={18} />, roles: ['Gestor', 'Técnico'] },
  { to: '/users', label: 'Usuários', icon: <Users size={18} />, roles: ['Gestor'] },
]

export default function Sidebar() {
  const { user } = useAuth()

  const visible = navItems.filter(
    (item) => !item.roles || (user && item.roles.includes(user.role))
  )

  return (
    <aside className="flex flex-col w-56 min-h-screen bg-slate-900 text-slate-100 py-6 px-3 gap-1 flex-shrink-0">
      {/* Logo */}
      <div className="mb-6 px-3">
        <span className="text-lg font-bold tracking-tight">Revalle</span>
        <p className="text-xs text-slate-400">Controle de Estoque</p>
      </div>

      <nav className="flex flex-col gap-0.5">
        {visible.map((item) => (
          <NavLink
            key={item.to}
            to={item.to}
            end={item.to === '/'}
            className={({ isActive }) =>
              cn(
                'flex items-center gap-3 rounded-md px-3 py-2 text-sm font-medium transition-colors',
                isActive
                  ? 'bg-slate-700 text-white'
                  : 'text-slate-400 hover:bg-slate-800 hover:text-white'
              )
            }
          >
            {item.icon}
            {item.label}
          </NavLink>
        ))}
      </nav>
    </aside>
  )
}
