import { useAuth } from '@/contexts/AuthContext'
import { useNavigate } from 'react-router-dom'
import { LogOut, Moon, Sun } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { useTheme } from '@/lib/theme'

export default function TopBar() {
  const { user, logout } = useAuth()
  const navigate = useNavigate()
  const { theme, toggle } = useTheme()

  async function handleLogout() {
    await logout()
    navigate('/login')
  }

  return (
    <header className="h-14 bg-surface border-b border-border px-4 sm:px-6 flex items-center justify-end gap-3 shrink-0 z-30">
      {user && (
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            <span className="text-body-sm text-foreground">{user.username}</span>
            <span className="text-[11px] text-muted-foreground font-mono uppercase">
              {user.role}
            </span>
          </div>

          <Button
            variant="ghost"
            size="icon"
            onClick={toggle}
            aria-label={theme === 'dark' ? 'Ativar modo claro' : 'Ativar modo escuro'}
            title={theme === 'dark' ? 'Modo claro' : 'Modo escuro'}
            className="text-muted-foreground hover:text-foreground"
          >
            {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
          </Button>

          <Button
            variant="ghost"
            size="icon"
            onClick={handleLogout}
            title="Sair do sistema"
            className="text-muted-foreground hover:text-destructive"
          >
            <LogOut size={16} />
          </Button>
        </div>
      )}
    </header>
  )
}
