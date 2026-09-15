import { useState, useRef, useEffect } from 'react'
import { Check, ChevronDown, Search, X } from 'lucide-react'
import { cn } from '@/lib/utils'

export interface SearchableOption {
  value: string
  label: string
  subtitle?: string
}

interface SearchableSelectProps {
  options: SearchableOption[]
  value: string
  onValueChange: (val: string) => void
  placeholder?: string
  searchPlaceholder?: string
  className?: string
  disabled?: boolean
  required?: boolean
  /**
   * Controle externo da busca (mesma convenção de `pagination.search` /
   * `pagination.onSearchChange` do `DataTable`): quando `onSearchChange` é
   * passado, o campo de busca do dropdown deixa de ser interno — passa a
   * exibir `search` e delegar toda digitação para `onSearchChange`, e o
   * componente PARA de filtrar `options` localmente (assume que quem chamou
   * já mandou a lista filtrada, tipicamente por busca no servidor). Sem essa
   * prop, o componente se comporta exatamente como antes: busca interna,
   * não controlada, filtrando a lista local.
   */
  search?: string
  onSearchChange?: (termo: string) => void
}

export function SearchableSelect({
  options,
  value,
  onValueChange,
  placeholder = 'Selecione uma opção...',
  searchPlaceholder = 'Digitar para buscar...',
  className,
  disabled = false,
  search: searchControlado,
  onSearchChange,
}: SearchableSelectProps) {
  const [isOpen, setIsOpen] = useState(false)
  const [searchInterno, setSearchInterno] = useState('')
  const containerRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  // Modo controlado: só existe quando quem chama passa `onSearchChange`
  // (mesmo espírito do `DataTable`, que só liga a busca externa quando
  // `pagination.onSearchChange` é fornecido).
  const isControlado = onSearchChange !== undefined
  const search = isControlado ? (searchControlado ?? '') : searchInterno
  function handleSearchChange(termo: string) {
    if (isControlado) {
      onSearchChange?.(termo)
    } else {
      setSearchInterno(termo)
    }
  }

  const selectedOption = options.find((o) => o.value === value)

  // Fecha o dropdown se clicar fora do componente
  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [])

  // Foca no input de busca automaticamente ao abrir. Só limpa a busca ao
  // fechar no modo não controlado — no modo controlado, quem dona o estado é
  // quem chama, e o componente não decide por ela quando limpar.
  useEffect(() => {
    if (isOpen) {
      setTimeout(() => inputRef.current?.focus(), 50)
    } else if (!isControlado) {
      setSearchInterno('')
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen])

  // No modo controlado, `options` já chega filtrada (busca no servidor) —
  // filtrar de novo aqui seria redundante e, pior, esconderia resultados que
  // não batem com o texto ainda não aplicado (debounce em andamento).
  const filtered = isControlado
    ? options
    : options.filter((o) => {
        if (!search.trim()) return true
        const term = search.toLowerCase().trim()
        return (
          o.label.toLowerCase().includes(term) ||
          (o.subtitle && o.subtitle.toLowerCase().includes(term)) ||
          o.value.toLowerCase().includes(term)
        )
      })

  return (
    <div ref={containerRef} className={cn('relative w-full select-none', className)}>
      {/* Trigger Button */}
      <button
        type="button"
        disabled={disabled}
        onClick={() => !disabled && setIsOpen(!isOpen)}
        className={cn(
          'flex h-10 w-full items-center justify-between rounded-md border border-input bg-surface px-3 py-2 text-sm text-foreground transition-colors duration-micro focus-visible:ring-1 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-50',
          isOpen && 'ring-1 ring-ring border-border-strong'
        )}
      >
        <span className="truncate font-medium text-xs sm:text-sm">
          {selectedOption ? (
            <span className="flex items-center gap-1.5 truncate">
              <span className="font-semibold text-foreground">{selectedOption.label}</span>
              {selectedOption.subtitle && (
                <span className="text-muted-foreground font-normal text-xs truncate">
                  ({selectedOption.subtitle})
                </span>
              )}
            </span>
          ) : (
            <span className="text-muted-foreground font-normal">{placeholder}</span>
          )}
        </span>
        <ChevronDown size={16} className={cn('text-muted-foreground transition-transform duration-micro', isOpen && 'rotate-180 text-foreground')} />
      </button>

      {/* Dropdown Popover */}
      {isOpen && (
        <div className="absolute left-0 top-full z-50 mt-1.5 w-full surface-panel shadow-overlay p-2 space-y-2 max-h-80 flex flex-col">
          {/* Live Search Input */}
          <div className="relative flex-shrink-0">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none" />
            <input
              ref={inputRef}
              type="text"
              value={search}
              onChange={(e) => handleSearchChange(e.target.value)}
              placeholder={searchPlaceholder}
              className="w-full h-9 pl-9 pr-8 text-xs font-medium bg-surface-alt rounded-md border border-border focus-visible:ring-1 focus-visible:ring-ring text-foreground placeholder:text-muted-foreground"
            />
            {search && (
              <button
                type="button"
                onClick={() => handleSearchChange('')}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground p-0.5 rounded-full hover:bg-surface-alt"
              >
                <X size={12} />
              </button>
            )}
          </div>

          {/* Options Scrollable Container */}
          <div className="flex-1 overflow-y-auto max-h-60 space-y-0.5 pr-1">
            {filtered.length === 0 ? (
              <div className="py-6 text-center text-xs text-muted-foreground font-medium">
                Nenhum equipamento encontrado com "{search}"
              </div>
            ) : (
              filtered.map((opt) => {
                const isSelected = opt.value === value
                return (
                  <button
                    key={opt.value}
                    type="button"
                    onClick={() => {
                      onValueChange(opt.value)
                      setIsOpen(false)
                    }}
                    className={cn(
                      'w-full flex items-center justify-between rounded-md px-3 py-2 text-left text-xs font-semibold transition-colors duration-micro',
                      isSelected
                        ? 'bg-primary text-primary-foreground'
                        : 'text-foreground hover:bg-surface-alt'
                    )}
                  >
                    <div className="truncate pr-2">
                      <p className="truncate font-semibold">{opt.label}</p>
                      {opt.subtitle && (
                        <p className={cn('truncate text-[10px] font-normal', isSelected ? 'text-border-strong' : 'text-muted-foreground')}>
                          {opt.subtitle}
                        </p>
                      )}
                    </div>
                    {isSelected && <Check size={14} className="flex-shrink-0 text-primary-foreground" />}
                  </button>
                )
              })
            )}
          </div>
        </div>
      )}
    </div>
  )
}
