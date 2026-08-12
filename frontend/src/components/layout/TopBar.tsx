import { useAuth } from '@/contexts/AuthContext'
import { useNavigate, useLocation } from 'react-router-dom'
import { LogOut, User, Sparkles } from 'lucide-react'
import { Button } from '@/components/ui/button'

const ROLE_BADGES: Record<string, string> = {
  Gestor: 'bg-purple-50 text-purple-700 border-purple-200/80 shadow-sm shadow-purple-500/10',
  'Técnico': 'bg-indigo-50 text-indigo-700 border-indigo-200/80 shadow-sm shadow-indigo-500/10',
  'Jovem Aprendiz': 'bg-emerald-50 text-emerald-700 border-emerald-200/80 shadow-sm shadow-emerald-500/10',
}

const PAGE_TITLES: Record<string, string> = {
  '/': 'Visão Geral do Estoque',
  '/register': 'Cadastrar Novo Equipamento',
  '/peripherals': 'Gestão de Periféricos',
  '/link': 'Vincular Periférico a Equipamento',
  '/loan': 'Realizar Empréstimo',
  '/return': 'Confirmar Devolução',
  '/remove': 'Remover / Estornar Itens',
  '/history': 'Histórico & Auditoria de Movimentações',
  '/report': 'Relatórios Gerenciais BI',
  '/charts': 'Dashboard & Análise de Gráficos',
  '/terms': 'Termos de Responsabilidade',
  '/users': 'Gestão de Usuários do Sistema',
  '/unidades': 'Unidades Operacionais & Revendas',
}

export default function TopBar({ title }: { title?: string }) {
  const { user, logout } = useAuth()
  const navigate = useNavigate()
  const location = useLocation()

  const activeTitle = title || PAGE_TITLES[location.pathname] || 'Controle de Estoque'

  async function handleLogout() {
    await logout()
    navigate('/login')
  }

  return (
    <header className="sticky top-0 z-30 flex items-center justify-between px-6 py-3.5 h-16 bg-white/85 backdrop-blur-md border-b border-slate-200/80 shadow-sm transition-all">
      {/* Title & Status */}
      <div className="flex items-center gap-3">
        <div>
          <h1 className="text-base font-bold tracking-tight text-slate-800 font-heading">
            {activeTitle}
          </h1>
          <div className="flex items-center gap-2 text-xs text-slate-400">
            <span className="inline-flex items-center gap-1">
              <span className="h-1.5 w-1.5 rounded-full bg-emerald-500 animate-pulse" />
              <span className="text-[11px] font-medium text-slate-500">Sistema Conectado</span>
            </span>
          </div>
        </div>
      </div>

      {/* User Actions */}
      {user && (
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2.5 px-3 py-1.5 rounded-full bg-white/90 border border-slate-200/80 shadow-sm">
            <div className="h-6 w-6 rounded-full bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold text-xs">
              <User size={13} />
            </div>
            <span className="text-xs font-semibold text-slate-700">{user.username}</span>
            <span
              className={`text-[11px] font-bold px-2 py-0.5 rounded-full border ${ROLE_BADGES[user.role] || 'bg-slate-100 text-slate-700'}`}
            >
              {user.role}
            </span>
          </div>

          <Button
            variant="ghost"
            size="sm"
            onClick={handleLogout}
            title="Sair do sistema"
            className="h-9 w-9 p-0 rounded-full text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors"
          >
            <LogOut size={17} />
          </Button>
        </div>
      )}
    </header>
  )
}

