import { useEffect, useState } from 'react'
import { Outlet, Navigate, useLocation } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import Sidebar from './Sidebar'
import TopBar from './TopBar'
import { Loader2, X } from 'lucide-react'

export default function AppLayout() {
  const { user, isLoading } = useAuth()
  const [collapsed, setCollapsed] = useState(false)
  const [mobileNavOpen, setMobileNavOpen] = useState(false)
  const location = useLocation()

  // Close the mobile drawer automatically on route change.
  useEffect(() => setMobileNavOpen(false), [location.pathname])

  if (isLoading) {
    return (
      <div className="flex h-screen items-center justify-center bg-background">
        <div className="flex flex-col items-center gap-4 p-8 rounded-lg surface shadow-2xl">
          <Loader2 className="h-8 w-8 animate-spin text-primary" />
          <p className="text-sm font-medium text-muted-foreground">Carregando sistema...</p>
        </div>
      </div>
    )
  }

  if (!user) return <Navigate to="/login" replace />

  return (
    <div className="flex h-screen bg-background overflow-hidden font-sans">
      {/* Desktop sidebar (lg and up) */}
      <Sidebar collapsed={collapsed} onToggle={() => setCollapsed((c) => !c)} />

      {/* Mobile sidebar drawer (below lg) */}
      {mobileNavOpen && (
        <div className="fixed inset-0 z-40 lg:hidden">
          <div className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={() => setMobileNavOpen(false)} />
          <div className="absolute left-0 top-0 h-full w-72 shadow-2xl animate-in slide-in-from-left duration-200">
            <div className="relative h-full">
              <button
                onClick={() => setMobileNavOpen(false)}
                className="absolute -right-11 top-3 h-9 w-9 rounded-full bg-card border border-border flex items-center justify-center text-muted-foreground"
                aria-label="Fechar menu"
              >
                <X size={16} />
              </button>
              <Sidebar collapsed={false} onToggle={() => setMobileNavOpen(false)} forceVisible />
            </div>
          </div>
        </div>
      )}

      <div className="flex flex-col flex-1 min-w-0 overflow-y-auto">
        <TopBar onOpenMobileNav={() => setMobileNavOpen(true)} />
        <main className="p-4 sm:p-6 md:p-8 space-y-6 max-w-[1440px] w-full mx-auto">
          {/* Key on pathname triggers a fresh fade/slide on every route change */}
          <div key={location.pathname} className="animate-in fade-in slide-in-from-bottom-1 duration-200">
            <Outlet />
          </div>
        </main>
      </div>
    </div>
  )
}
