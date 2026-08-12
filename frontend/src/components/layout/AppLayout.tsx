import { Outlet, Navigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import Sidebar from './Sidebar'
import TopBar from './TopBar'
import { Loader2 } from 'lucide-react'

export default function AppLayout() {
  const { user, isLoading } = useAuth()

  if (isLoading) {
    return (
      <div className="flex h-screen items-center justify-center bg-slate-950">
        <div className="flex flex-col items-center gap-4 p-8 rounded-2xl bg-slate-900/60 border border-slate-800 backdrop-blur-md shadow-2xl">
          <div className="relative">
            <Loader2 className="h-10 w-10 animate-spin text-indigo-500" />
            <span className="absolute inset-0 rounded-full bg-indigo-500/20 animate-ping -z-10" />
          </div>
          <p className="text-sm font-semibold text-slate-300">Carregando sistema...</p>
        </div>
      </div>
    )
  }

  if (!user) return <Navigate to="/login" replace />

  return (
    <div className="flex h-screen bg-slate-50/70 overflow-hidden font-sans">
      <Sidebar />
      <div className="flex flex-col flex-1 min-w-0 overflow-y-auto bg-[radial-gradient(ellipse_at_top,_var(--tw-gradient-stops))] from-indigo-50/40 via-slate-50/60 to-slate-100/80">
        <TopBar />
        <main className="p-6 md:p-8 space-y-6 max-w-7xl w-full mx-auto animate-fade-in">
          <Outlet />
        </main>
      </div>
    </div>
  )
}


