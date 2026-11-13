import { useState } from 'react'
import { Link, NavLink } from 'react-router-dom'
import { cn } from '@/lib/utils'
import { useAuth } from '@/contexts/AuthContext'
import { ChevronRight } from 'lucide-react'
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
      {/* Bloco de marca igual ao da origin/FT_STC: o logo é a palavra
          "revalle" branca com fundo transparente — `brightness-0` a deixa
          preta no tema claro e `dark:brightness-100` devolve o branco no
          escuro. Só o ponto de status troca o verde cru dela pelo token
          `--status-available` (mesmo usado pelo StatusBadge). */}
      <Link
        to="/"
        onClick={onNavigate}
        className="flex flex-col items-start justify-center py-5 px-4 border-b border-border shrink-0 gap-2.5 transition-colors hover:bg-surface-alt group"
        title="Início - Revalle Controle de Patrimônio"
      >
        <div className="w-full flex items-center justify-start overflow-hidden">
          <img
            src="/logo-revalle.png"
            alt="Revalle"
            className="h-14 w-auto max-w-[205px] object-contain filter brightness-0 dark:brightness-100 drop-shadow-md transition-transform duration-200 group-hover:scale-[1.02]"
          />
        </div>
        <div className="flex items-center gap-1.5 px-2 py-0.5 rounded-full bg-brand-soft">
          <span className="w-1.5 h-1.5 rounded-full bg-[var(--status-available)] motion-safe:animate-pulse" />
          <span className="text-[10px] font-semibold text-brand-soft-foreground tracking-widest uppercase font-mono">
            Controle de Patrimônio
          </span>
        </div>
      </Link>

      {/* Navigation Groups — colapsáveis, com o estado persistido em
          localStorage (T7). A LISTA em si vem inteira de `navigation.ts`
          (`getVisibleNavGroups`): nunca duplicar os destinos aqui, senão a
          paleta de comandos (Ctrl+K), que consome a mesma fonte, ficaria
          divergente da Sidebar. */}
      <nav className="flex-1 overflow-y-auto py-4 px-2.5 space-y-4">
        {visibleGroups.map((group) => {
          const aberto = gruposAbertos[group.title] ?? true
          return (
            <div key={group.title} className="space-y-1">
              <button
                type="button"
                onClick={() => alternarGrupo(group.title)}
                className="flex w-full items-center justify-between px-2.5 py-1 text-caption text-brand-text hover:text-foreground transition-colors cursor-pointer group/hdr"
                aria-expanded={aberto}
              >
                <span className="font-semibold tracking-wider">{group.title}</span>
                <ChevronRight
                  className={cn(
                    'size-3.5 text-brand-text transition-transform duration-micro group-hover/hdr:text-foreground',
                    aberto && 'rotate-90'
                  )}
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
                            ? 'bg-lime text-lime-foreground font-semibold'
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
      {/* Sem bloco de usuário no rodapé: o perfil (nome e papel) aparece só
          no TopBar, para não repetir a mesma informação em dois lugares. */}
    </aside>
  )
}
