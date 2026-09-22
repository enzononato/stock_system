import { useState } from 'react'
import { Link, NavLink } from 'react-router-dom'
import { cn } from '@/lib/utils'
import { useAuth } from '@/contexts/AuthContext'
import { ShieldCheck, ChevronRight } from 'lucide-react'
import { getVisibleNavGroups } from './navigation'

interface SidebarProps {
  className?: string
  /** Chamado ao clicar em um item de navegação — usado pela gaveta mobile para fechar após navegar. */
  onNavigate?: () => void
}

// Chave de persistência do estado de colapso dos grupos — um objeto
// `{ [título do grupo]: aberto }`. Grupo ausente do objeto salvo (primeira
// visita, ou grupo novo adicionado depois) é tratado como aberto por padrão.
const STORAGE_KEY = 'sidebar-groups-abertos'

function lerGruposAbertosSalvos(): Record<string, boolean> {
  try {
    const salvo = localStorage.getItem(STORAGE_KEY)
    if (salvo) return JSON.parse(salvo) as Record<string, boolean>
  } catch {
    // Storage indisponível (modo privado, quota etc.) — segue com tudo aberto.
  }
  return {}
}

export default function Sidebar({ className, onNavigate }: SidebarProps) {
  const { user } = useAuth()
  const visibleGroups = getVisibleNavGroups(user?.role)
  const [gruposAbertos, setGruposAbertos] = useState<Record<string, boolean>>(lerGruposAbertosSalvos)

  function alternarGrupo(titulo: string) {
    setGruposAbertos((prev) => {
      const aberto = prev[titulo] ?? true
      const atualizado = { ...prev, [titulo]: !aberto }
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(atualizado))
      } catch {
        // Ignora erro de storage — o colapso ainda funciona nesta sessão.
      }
      return atualizado
    })
  }

  return (
    <aside className={cn('flex flex-col w-64 h-full bg-sidebar border-r border-sidebar-border select-none shrink-0', className)}>
      {/* Bloco de marca — com o logo da referência (origin/FT_STC), mas sem a
          cor de status crua dela (uma classe Tailwind de verde fixo): o
          indicador usa o token `--status-available` (mesmo usado pelo
          StatusBadge), preservando a paleta monocromática com cor só em
          status/ação destrutiva. */}
      <Link
        to="/"
        onClick={onNavigate}
        className="flex flex-col items-start justify-center py-4 px-3.5 border-b border-border shrink-0 gap-2 transition-colors hover:bg-surface-alt"
        title="Início — Revalle Controle de Patrimônio"
      >
        <img
          src="/logo-revalle.jpg"
          alt="Revalle"
          className="h-8 w-auto max-w-[160px] object-contain rounded"
        />
        <div className="flex items-center gap-1.5 px-2 py-0.5 rounded bg-surface-alt border border-border">
          <span className="w-1.5 h-1.5 rounded-full bg-[var(--status-available)]" />
          <span className="text-[10px] font-semibold text-foreground tracking-widest uppercase font-mono">
            Controle de Estoque TI
          </span>
        </div>
      </Link>

      {/* Navigation Groups — colapsáveis, com o estado persistido em
          localStorage (T7). A LISTA em si vem inteira de `navigation.ts`
          (`getVisibleNavGroups`): nunca duplicar os destinos aqui, senão a
          paleta de comandos (Ctrl+K) e os breadcrumbs, que consomem a mesma
          fonte, ficariam divergentes da Sidebar. */}
      <nav className="flex-1 overflow-y-auto py-4 px-2.5 space-y-1">
        {visibleGroups.map((group) => {
          const aberto = gruposAbertos[group.title] ?? true
          return (
            <div key={group.title} className="space-y-0.5">
              <button
                type="button"
                onClick={() => alternarGrupo(group.title)}
                className="flex w-full items-center justify-between px-2 py-1 text-caption text-muted-foreground hover:text-foreground transition-colors rounded"
                aria-expanded={aberto}
              >
                <span>{group.title}</span>
                <ChevronRight
                  size={13}
                  className={cn('transition-transform duration-micro', aberto && 'rotate-90')}
                  aria-hidden="true"
                />
              </button>

              {aberto && (
                <div className="space-y-0.5">
                  {group.items.map((item) => (
                    <NavLink
                      key={item.to}
                      to={item.to}
                      end={item.to === '/'}
                      onClick={onNavigate}
                      className={({ isActive }) =>
                        cn(
                          // `group`: permite que os ícones animados (ver
                          // animated-sidebar-icons.css) disparem a animação a
                          // partir do hover/focus do item inteiro, não só do SVG.
                          'group flex items-center gap-2.5 px-2.5 py-1.5 rounded text-body-sm transition-colors duration-micro',
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
              )}
            </div>
          )
        })}
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
