import { useAuth } from '@/contexts/AuthContext'
import { useNavigate } from 'react-router-dom'
import { LogOut, User } from 'lucide-react'
import { Button } from '@/components/ui/button'

const ROLE_COLORS: Record<string, string> = {
  Gestor: 'bg-purple-100 text-purple-800',
  'Técnico': 'bg-blue-100 text-blue-800',
  'Jovem Aprendiz': 'bg-green-100 text-green-800',
}

export default function TopBar({ title }: { title?: string }) {
  const { user, logout } = useAuth()
  const navigate = useNavigate()

  async function handleLogout() {
    await logout()
    navigate('/login')
  }

  return (
    <header className="flex items-center justify-between px-6 py-3 bg-white border-b border-slate-200 h-14">
      <h1 className="text-lg font-semibold text-slate-800">{title || 'Controle de Estoque'}</h1>
      {user && (
        <div className="flex items-center gap-3">
          <span
            className={`text-xs font-medium px-2 py-0.5 rounded-full ${ROLE_COLORS[user.role] || 'bg-gray-100 text-gray-700'}`}
          >
            {user.role}
          </span>
          <div className="flex items-center gap-1.5 text-sm text-slate-600">
            <User size={14} />
            <span>{user.username}</span>
          </div>
          <Button variant="ghost" size="sm" onClick={handleLogout} className="text-slate-500 hover:text-red-600">
            <LogOut size={16} />
          </Button>
        </div>
      )}
    </header>
  )
}
