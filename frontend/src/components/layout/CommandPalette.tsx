import { useEffect, useMemo, useRef, useState, type KeyboardEvent } from 'react'
import { createPortal } from 'react-dom'
import { useNavigate } from 'react-router-dom'
import { Search } from 'lucide-react'
import { useAuth } from '@/contexts/AuthContext'
import { cn } from '@/lib/utils'
import { getVisibleNavItems } from './navigation'

interface CommandPaletteProps {
  open: boolean
  onOpenChange: (open: boolean) => void
}

/**
 * Paleta de comandos (Ctrl+K / ⌘K), implementada inteiramente à mão — a
 * biblioteca `cmdk` não está instalada e a Task 8 proíbe adicionar
 * dependências novas. Alimenta a lista de destinos a partir da MESMA
 * estrutura de navegação da sidebar (`navigation.ts`), já filtrada pelo
 * papel do usuário logado, para nunca oferecer uma tela que ele não pode
 * abrir de fato.
 */
export function CommandPalette({ open, onOpenChange }: CommandPaletteProps) {
  const { user } = useAuth()
  const navigate = useNavigate()
  const [query, setQuery] = useState('')
  const [activeIndex, setActiveIndex] = useState(0)
  const inputRef = useRef<HTMLInputElement>(null)

  const allItems = useMemo(() => getVisibleNavItems(user?.role), [user?.role])

  const results = useMemo(() => {
    const termo = query.trim().toLowerCase()
    if (!termo) return allItems
    return allItems.filter((item) => item.label.toLowerCase().includes(termo))
  }, [allItems, query])

  // Atalho global Ctrl+K / ⌘K: alterna a paleta e evita que o navegador abra
  // a própria busca (padrão do Ctrl+K nativo do Chrome/Edge/Firefox).
  useEffect(() => {
    function handleKeyDown(e: KeyboardEvent | globalThis.KeyboardEvent) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault()
        onOpenChange(!open)
      }
    }
    window.addEventListener('keydown', handleKeyDown as EventListener)
    return () => window.removeEventListener('keydown', handleKeyDown as EventListener)
  }, [open, onOpenChange])

  // Escape fecha mesmo se o foco não estiver no campo de busca (ex.: usuário
  // passou o mouse sobre um item da lista antes de apertar Escape).
  useEffect(() => {
    if (!open) return
    function handleEscape(e: globalThis.KeyboardEvent) {
      if (e.key === 'Escape') onOpenChange(false)
    }
    window.addEventListener('keydown', handleEscape)
    return () => window.removeEventListener('keydown', handleEscape)
  }, [open, onOpenChange])

  // Toda vez que a paleta abre, reseta a busca/seleção e foca o campo.
  useEffect(() => {
    if (!open) return
    setQuery('')
    setActiveIndex(0)
    const id = requestAnimationFrame(() => inputRef.current?.focus())
    return () => cancelAnimationFrame(id)
  }, [open])

  // Mantém o índice ativo dentro dos limites quando o filtro muda o tamanho da lista.
  useEffect(() => {
    setActiveIndex((i) => Math.min(i, Math.max(results.length - 1, 0)))
  }, [results.length])

  function handleSelect(to: string) {
    onOpenChange(false)
    navigate(to)
  }

  function handleInputKeyDown(e: KeyboardEvent<HTMLInputElement>) {
    if (e.key === 'ArrowDown') {
      e.preventDefault()
      setActiveIndex((i) => (results.length === 0 ? 0 : (i + 1) % results.length))
    } else if (e.key === 'ArrowUp') {
      e.preventDefault()
      setActiveIndex((i) => (results.length === 0 ? 0 : (i - 1 + results.length) % results.length))
    } else if (e.key === 'Enter') {
      e.preventDefault()
      const item = results[activeIndex]
      if (item) handleSelect(item.to)
    }
  }

  if (!open) return null

  return createPortal(
    <div className="fixed inset-0 z-50 flex items-start justify-center pt-24 p-4 bg-black/50 select-none">
      {/* Overlay — fecha ao clicar fora do painel */}
      <div className="absolute inset-0" aria-hidden="true" onClick={() => onOpenChange(false)} />

      <div
        role="dialog"
        aria-modal="true"
        aria-label="Paleta de comandos"
        className="relative z-10 w-full max-w-lg surface-panel shadow-overlay overflow-hidden"
      >
        <div className="flex items-center gap-2.5 border-b border-border px-3.5 h-11 shrink-0">
          <Search size={16} className="text-muted-foreground shrink-0" aria-hidden="true" />
          <input
            ref={inputRef}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={handleInputKeyDown}
            placeholder="Buscar uma tela do sistema..."
            aria-label="Buscar uma tela do sistema"
            className="flex-1 bg-transparent text-body-sm text-foreground placeholder:text-muted-foreground outline-none"
          />
          <kbd className="hidden sm:inline-flex text-[10px] font-mono text-muted-foreground border border-border rounded-sm px-1.5 py-0.5 shrink-0">
            Esc
          </kbd>
        </div>

        <ul role="listbox" aria-label="Destinos de navegação" className="max-h-80 overflow-y-auto py-1.5">
          {results.length === 0 && (
            <li className="px-3.5 py-6 text-center text-body-sm text-muted-foreground">
              Nenhum destino encontrado.
            </li>
          )}
          {results.map((item, index) => (
            <li key={item.to}>
              <button
                type="button"
                role="option"
                aria-selected={index === activeIndex}
                onMouseEnter={() => setActiveIndex(index)}
                onClick={() => handleSelect(item.to)}
                className={cn(
                  'flex w-full items-center gap-2.5 px-3.5 py-2 text-body-sm text-left transition-colors duration-micro',
                  index === activeIndex
                    ? 'bg-surface-alt text-foreground'
                    : 'text-muted-foreground hover:bg-surface-alt hover:text-foreground'
                )}
              >
                <item.icon size={16} className="shrink-0" aria-hidden="true" />
                <span className="truncate">{item.label}</span>
              </button>
            </li>
          ))}
        </ul>
      </div>
    </div>,
    document.body
  )
}

export function CommandPaletteTrigger({ onClick }: { onClick: () => void }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="hidden sm:flex items-center gap-1.5 h-8 px-2.5 rounded border border-border bg-surface hover:bg-surface-alt hover:border-border-strong text-muted-foreground hover:text-foreground transition-colors duration-micro text-xs"
      title="Busca rápida (Ctrl+K)"
    >
      <Search size={14} className="shrink-0" aria-hidden="true" />
      <span>Buscar</span>
      <kbd className="ml-1 rounded-sm border border-border bg-surface px-1 py-0.5 text-[10px] font-mono text-muted-foreground">
        Ctrl K
      </kbd>
    </button>
  )
}
