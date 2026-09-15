import { NavLink } from 'react-router-dom'
import { cn } from '@/lib/utils'
import { useAuth } from '@/contexts/AuthContext'
import { Boxes, ShieldCheck } from 'lucide-react'
import { getVisibleNavGroups } from './navigation'

interface SidebarProps {
  className?: string
  /** Chamado ao clicar em um item de navegação — usado pela gaveta mobile para fechar após navegar. */
  onNavigate?: () => void
}

export default function Sidebar({ className, onNavigate }: SidebarProps) {
  const { user } = useAuth()
  const visibleGroups = getVisibleNavGroups(user?.role)

  return (
    <aside className={cn('flex flex-col w-64 h-full bg-sidebar border-r border-sidebar-border select-none shrink-0', className)}>
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
        {visibleGroups.map((group) => (
          <div key={group.title} className="space-y-0.5">
            <h3 className="text-caption text-muted-foreground px-2 py-1">
              {group.title}
            </h3>
            <div className="space-y-0.5">
              {group.items.map((item) => (
                <NavLink
                  key={item.to}
                  to={item.to}
                  end={item.to === '/'}
                  onClick={onNavigate}
                  className={({ isActive }) =>
                    cn(
                      'flex items-center gap-2.5 px-2.5 py-1.5 rounded text-body-sm transition-colors duration-micro',
                      isActive
                        ? 'bg-surface-alt text-foreground font-semibold border-l-2 border-foreground'
                        : 'text-muted-foreground hover:text-foreground hover:bg-surface-alt'
                    )
                  }
                >
                  <item.icon size={18} />
                  <span className="truncate">{item.label}</span>
                </NavLink>
              ))}
            </div>
          </div>
        ))}
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
