import { useAuth } from '@/contexts/AuthContext'
import { useNavigate, useLocation, Link } from 'react-router-dom'
import { LogOut, Menu, ChevronRight } from 'lucide-react'
import { Button } from '@/components/ui/button'

const ROLE_BADGE: Record<string, string> = {
  Gestor: 'bg-purple/10 text-purple-light border-purple/25',
  'Técnico': 'bg-info/10 text-info border-info/25',
  'Jovem Aprendiz': 'bg-success/10 text-success border-success/25',
}

// path -> [breadcrumb trail, title]. Trail is what shows before the current
// page (e.g. "Estoque / Cadastrar Equipamento").
const PAGE_META: Record<string, { trail: string[]; title: string }> = {
  '/': { trail: [], title: 'Visão Geral do Estoque' },
  '/register': { trail: ['Estoque'], title: 'Cadastrar Novo Equipamento' },
  '/peripherals': { trail: ['Estoque'], title: 'Gestão de Periféricos' },
  '/link': { trail: ['Estoque', 'Periféricos'], title: 'Vincular Periférico' },
  '/loan': { trail: ['Estoque'], title: 'Realizar Empréstimo' },
  '/return': { trail: ['Estoque'], title: 'Confirmar Devolução' },
  '/remove': { trail: ['Administração'], title: 'Remover / Estornar Itens' },
  '/history': { trail: ['Relatórios'], title: 'Histórico & Auditoria' },
  '/report': { trail: ['Relatórios'], title: 'Relatórios Gerenciais BI' },
  '/charts': { trail: ['Estoque'], title: 'Dashboard & Análise de Gráficos' },
  '/terms': { trail: ['Estoque'], title: 'Termos de Responsabilidade' },
  '/users': { trail: ['Administração'], title: 'Gestão de Usuários' },
  '/unidades': { trail: ['Administração'], title: 'Unidades Operacionais & Revendas' },
}

export default function TopBar({ title, onOpenMobileNav }: { title?: string; onOpenMobileNav?: () => void }) {
  const { user, logout } = useAuth()
  const navigate = useNavigate()
  const location = useLocation()

  const meta = PAGE_META[location.pathname]
  const activeTitle = title || meta?.title || 'Controle de Estoque'
  const trail = meta?.trail ?? []

  async function handleLogout() {
    await logout()
    navigate('/login')
  }

  return (
    <header className="sticky top-0 z-30 flex items-center justify-between gap-3 px-4 sm:px-6 py-3 h-16 bg-card/90 backdrop-blur-md border-b border-border">
      <div className="flex items-center gap-3 min-w-0">
        {/* Mobile nav toggle — sidebar is hidden below lg, this is the only entry point */}
        {onOpenMobileNav && (
          <button
            onClick={onOpenMobileNav}
            className="lg:hidden h-9 w-9 shrink-0 rounded-md flex items-center justify-center text-muted-foreground hover:text-foreground hover:bg-accent transition-colors"
            aria-label="Abrir menu"
          >
            <Menu size={18} />
          </button>
        )}
        <div className="min-w-0">
          {trail.length > 0 && (
            <div className="hidden sm:flex items-center gap-1 text-[11px] text-muted-foreground mb-0.5">
              {trail.map((t, i) => (
                <span key={i} className="flex items-center gap-1">
                  {i > 0 && <ChevronRight size={10} />}
                  {i === 0 && trail.length > 0 && <Link to="/" className="hover:text-foreground transition-colors">{t}</Link>}
                  {i > 0 && <span>{t}</span>}
                </span>
              ))}
              <ChevronRight size={10} />
              <span className="text-foreground/70">{activeTitle}</span>
            </div>
          )}
          <h1 className="text-h3 text-foreground truncate">{activeTitle}</h1>
          {trail.length === 0 && (
            <div className="flex items-center gap-1.5">
              <span className="h-1.5 w-1.5 rounded-full bg-success animate-pulse" />
              <span className="text-[11px] font-medium text-muted-foreground">Sistema Conectado</span>
            </div>
          )}
        </div>
      </div>

      {user && (
        <div className="flex items-center gap-2 sm:gap-3 shrink-0">
          <div className="hidden sm:flex items-center gap-2 px-2.5 py-1.5 rounded-full bg-secondary/60 border border-border">
            <span className="text-xs font-medium text-foreground">{user.username}</span>
            <span className={`text-[10px] font-semibold px-2 py-0.5 rounded-full border ${ROLE_BADGE[user.role] || 'bg-secondary text-muted-foreground border-border'}`}>
              {user.role}
            </span>
          </div>
          <Button
            variant="ghost"
            size="icon"
            onClick={handleLogout}
            title="Sair do sistema"
            className="rounded-full text-muted-foreground hover:text-destructive hover:bg-destructive/10"
          >
            <LogOut size={16} />
          </Button>
        </div>
      )}
    </header>
  )
}
