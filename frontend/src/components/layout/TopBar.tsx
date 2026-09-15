import { useAuth } from '@/contexts/AuthContext'
import { useNavigate } from 'react-router-dom'
import { LogOut, Menu, Moon, Sun } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { useTheme } from '@/lib/theme'
import { BreadcrumbTrail } from './Breadcrumb'
import { CommandPaletteTrigger } from './CommandPalette'

interface TopBarProps {
  /** Abre a gaveta de navegação mobile (abaixo do breakpoint `lg`). */
  onOpenMobileNav: () => void
  /** Abre a paleta de comandos (Ctrl+K). */
  onOpenCommandPalette: () => void
}

export default function TopBar({ onOpenMobileNav, onOpenCommandPalette }: TopBarProps) {
  const { user, logout } = useAuth()
  const navigate = useNavigate()
  const { theme, toggle } = useTheme()

  async function handleLogout() {
    await logout()
    navigate('/login')
  }

  return (
    <header className="h-14 bg-surface border-b border-border px-4 sm:px-6 flex items-center justify-between gap-3 shrink-0 z-30">
      <div className="flex items-center gap-3 min-w-0">
        {/* Abre a gaveta de navegação — só existe abaixo de lg, onde a sidebar fixa some (Task 8, Step 1) */}
        <Button
          variant="ghost"
          size="icon"
          onClick={onOpenMobileNav}
          aria-label="Abrir menu de navegação"
          className="lg:hidden text-muted-foreground hover:text-foreground shrink-0"
        >
          <Menu size={18} />
        </Button>

        <BreadcrumbTrail />
      </div>

      {user && (
        <div className="flex items-center gap-2 sm:gap-3 shrink-0">
          <div className="hidden sm:flex items-center gap-2">
            <span className="text-body-sm text-foreground">{user.username}</span>
            <span className="text-[11px] text-muted-foreground font-mono uppercase">
              {user.role}
            </span>
          </div>

          <CommandPaletteTrigger onClick={onOpenCommandPalette} />

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
