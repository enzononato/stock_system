import { useEffect, useState } from 'react'
import { Outlet, Navigate, useLocation } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import Sidebar from './Sidebar'
import TopBar from './TopBar'
import { CommandPalette } from './CommandPalette'
import { Loader2 } from 'lucide-react'

export default function AppLayout() {
  const { user, isLoading } = useAuth()
  const location = useLocation()

  // Gaveta de navegação mobile (abaixo de `lg`) — acima de `lg` a sidebar
  // continua fixa e este estado não tem efeito nenhum (Task 8, Step 1).
  const [mobileNavOpen, setMobileNavOpen] = useState(false)
  const [commandPaletteOpen, setCommandPaletteOpen] = useState(false)

  // 1ª forma de fechar a gaveta: navegar para outra tela.
  useEffect(() => {
    setMobileNavOpen(false)
  }, [location.pathname])

  // 2ª forma de fechar a gaveta: Escape. (O clique fora é tratado pelo
  // próprio overlay escurecido, mais abaixo.)
  useEffect(() => {
    if (!mobileNavOpen) return
    function handleKeyDown(e: KeyboardEvent) {
      if (e.key === 'Escape') setMobileNavOpen(false)
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [mobileNavOpen])

  // Com a gaveta aberta, o conteúdo atrás não deve rolar.
  useEffect(() => {
    if (!mobileNavOpen) return
    const overflowOriginal = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    return () => {
      document.body.style.overflow = overflowOriginal
    }
  }, [mobileNavOpen])

  if (isLoading) {
    return (
      <div className="flex h-screen items-center justify-center bg-background">
        <div className="flex flex-col items-center gap-4">
          <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
          <p className="text-body-sm text-muted-foreground">Carregando sistema...</p>
        </div>
      </div>
    )
  }

  if (!user) return <Navigate to="/login" replace />

  return (
    <div className="flex h-screen bg-background text-foreground overflow-hidden font-sans">
      {/* Sidebar fixa — só a partir de `lg` */}
      <Sidebar className="hidden lg:flex" />

      {/* Gaveta mobile — some a partir de `lg`, sobre um overlay escurecido.
          O container precisa ser `flex`: um <div> em bloco comum ocuparia
          100% da largura mesmo só querendo abrigar a sidebar de w-64,
          ficando por cima do overlay (mesmo z-index) e roubando o clique
          que deveria fechar a gaveta na área escurecida ao lado dela. Como
          item de flex (sem flex-grow), o wrapper encolhe para a largura do
          conteúdo (a própria Sidebar), deixando o resto da área para o
          overlay clicável. */}
      {mobileNavOpen && (
        <div className="fixed inset-0 z-40 flex lg:hidden">
          <div
            className="absolute inset-0 bg-black/50"
            aria-hidden="true"
            onClick={() => setMobileNavOpen(false)}
          />
          <div className="relative h-full">
            <Sidebar onNavigate={() => setMobileNavOpen(false)} />
          </div>
        </div>
      )}

      <div className="flex flex-col flex-1 min-w-0">
        <TopBar
          onOpenMobileNav={() => setMobileNavOpen(true)}
          onOpenCommandPalette={() => setCommandPaletteOpen(true)}
        />
        <main className="flex-1 overflow-y-auto p-4 sm:p-6 lg:p-8">
          <div className="page-container-dense">
            <Outlet />
          </div>
        </main>
      </div>

      <CommandPalette open={commandPaletteOpen} onOpenChange={setCommandPaletteOpen} />
    </div>
  )
}
