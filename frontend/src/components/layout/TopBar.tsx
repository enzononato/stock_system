import { useState } from 'react'
import { useAuth } from '@/contexts/AuthContext'
import { useNavigate } from 'react-router-dom'
import { LogOut, Menu, Moon, Search, Sun } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { useTheme } from '@/lib/theme'

interface TopBarProps {
  /** Abre a gaveta de navegação mobile (abaixo do breakpoint `lg`). */
  onOpenMobileNav: () => void
  /** Abre a paleta de comandos (também acessível por Ctrl+K); é a única forma de busca no toque. */
  onOpenCommandPalette: () => void
}

/**
 * Cabeçalho igual ao da origin/FT_STC: sem campo de busca nem breadcrumbs,
 * só o perfil (nome e papel), a troca de tema e o botão de sair, alinhados à
 * direita. É o ÚNICO lugar que mostra o perfil — a Sidebar não repete.
 */
export default function TopBar({ onOpenMobileNav, onOpenCommandPalette }: TopBarProps) {
  const { user, logout } = useAuth()
  const navigate = useNavigate()
  const { theme, toggle } = useTheme()
  const [saindo, setSaindo] = useState(false)

  async function handleLogout() {
    if (saindo) return
    setSaindo(true)
    try {
      await logout()
      navigate('/login')
    } finally {
      setSaindo(false)
    }
  }

  return (
    <header className="h-14 bg-surface border-b border-border px-4 sm:px-6 flex items-center justify-between lg:justify-end gap-3 shrink-0 z-30">
      {/* Abre a gaveta de navegação — só existe abaixo de lg, onde a sidebar fixa some */}
      <Button
        variant="ghost"
        size="icon"
        onClick={onOpenMobileNav}
        aria-label="Abrir menu de navegação"
        className="lg:hidden text-muted-foreground hover:text-foreground shrink-0"
      >
        <Menu size={18} />
      </Button>

      {user && (
        <div className="flex min-w-0 items-center gap-3">
          <div className="flex min-w-0 items-center gap-2">
            <span className="truncate max-w-[40vw] text-body-sm text-foreground">{user.email ?? user.username}</span>
            <span className="hidden sm:inline text-[11px] text-muted-foreground font-mono uppercase">
              {user.role}
            </span>
          </div>

          <Button
            variant="ghost"
            size="icon"
            onClick={onOpenCommandPalette}
            aria-label="Buscar tela (Ctrl+K)"
            title="Buscar tela (Ctrl+K)"
            className="text-muted-foreground hover:text-foreground"
          >
            <Search size={16} />
          </Button>

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
            loading={saindo}
            aria-label="Sair do sistema"
            title="Sair do sistema"
            className="text-muted-foreground hover:text-destructive"
          >
            {!saindo && <LogOut size={16} />}
          </Button>
        </div>
      )}
    </header>
  )
}
