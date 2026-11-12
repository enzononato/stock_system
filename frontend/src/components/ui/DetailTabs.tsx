import { useRef, type KeyboardEvent, type ReactNode } from 'react'
import { cn } from '@/lib/utils'

interface DetailTabsProps {
  tabs: { id: string; label: string }[]
  /** Aba ativa; `undefined` quando não há nenhuma aba disponível. */
  activeId: string | undefined
  onSelect: (id: string) => void
  /** Nome acessível do `tablist`. */
  label: string
  /** Conteúdo do painel da aba ativa. */
  children: ReactNode
}

/**
 * Abas acessíveis das páginas de detalhe: `tablist`/`tab`/`tabpanel`, roving
 * tabindex e navegação por setas, Home e End. O estado (qual aba está ativa)
 * fica com quem usa — normalmente em `?aba=` na URL.
 */
export function DetailTabs({ tabs, activeId, onSelect, label, children }: DetailTabsProps) {
  const tabRefs = useRef<Record<string, HTMLButtonElement | null>>({})

  function select(tabId: string) {
    onSelect(tabId)
    tabRefs.current[tabId]?.focus()
  }

  function onKeyDown(e: KeyboardEvent<HTMLButtonElement>, index: number) {
    let target: number | null = null
    if (e.key === 'ArrowRight') target = (index + 1) % tabs.length
    else if (e.key === 'ArrowLeft') target = (index - 1 + tabs.length) % tabs.length
    else if (e.key === 'Home') target = 0
    else if (e.key === 'End') target = tabs.length - 1
    if (target === null) return
    e.preventDefault()
    select(tabs[target].id)
  }

  return (
    <div>
      <div role="tablist" aria-label={label} className="flex gap-1 overflow-x-auto border-b border-border">
        {tabs.map((tab, index) => {
          const selected = tab.id === activeId
          return (
            <button
              key={tab.id}
              ref={(el) => {
                tabRefs.current[tab.id] = el
              }}
              type="button"
              role="tab"
              id={`aba-${tab.id}`}
              aria-selected={selected}
              aria-controls={`painel-${tab.id}`}
              tabIndex={selected ? 0 : -1}
              onClick={() => select(tab.id)}
              onKeyDown={(e) => onKeyDown(e, index)}
              className={cn(
                '-mb-px whitespace-nowrap border-b-2 px-4 py-2 text-sm font-medium transition-colors',
                'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2',
                selected
                  ? 'border-brand-text text-brand-text'
                  : 'border-transparent text-muted-foreground hover:text-foreground'
              )}
            >
              {tab.label}
            </button>
          )
        })}
      </div>

      {activeId && (
        <div
          role="tabpanel"
          id={`painel-${activeId}`}
          aria-labelledby={`aba-${activeId}`}
          tabIndex={0}
          className="pt-5 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring"
        >
          {children}
        </div>
      )}
    </div>
  )
}
