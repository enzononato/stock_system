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
}

export function SearchableSelect({
  options,
  value,
  onValueChange,
  placeholder = 'Selecione uma opção...',
  searchPlaceholder = 'Digitar para buscar...',
  className,
  disabled = false,
}: SearchableSelectProps) {
  const [isOpen, setIsOpen] = useState(false)
  const [search, setSearch] = useState('')
  const containerRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

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

  // Foca no input de busca automaticamente ao abrir
  useEffect(() => {
    if (isOpen) {
      setTimeout(() => inputRef.current?.focus(), 50)
    } else {
      setSearch('')
    }
  }, [isOpen])

  const filtered = options.filter((o) => {
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
          'flex h-10 w-full items-center justify-between rounded-xl border border-slate-200 bg-white/90 px-3 py-2 text-sm text-slate-800 shadow-sm transition-all focus:outline-none focus:ring-2 focus:ring-indigo-500/30 disabled:cursor-not-allowed disabled:opacity-50',
          isOpen && 'ring-2 ring-indigo-500/30 border-indigo-400'
        )}
      >
        <span className="truncate font-medium text-xs sm:text-sm">
          {selectedOption ? (
            <span className="flex items-center gap-1.5 truncate">
              <span className="font-semibold text-slate-900">{selectedOption.label}</span>
              {selectedOption.subtitle && (
                <span className="text-slate-400 font-normal text-xs truncate">
                  ({selectedOption.subtitle})
                </span>
              )}
            </span>
          ) : (
            <span className="text-slate-400 font-normal">{placeholder}</span>
          )}
        </span>
        <ChevronDown size={16} className={cn('text-slate-400 transition-transform duration-200', isOpen && 'rotate-180 text-indigo-600')} />
      </button>

      {/* Dropdown Popover */}
      {isOpen && (
        <div className="absolute left-0 top-full z-50 mt-1.5 w-full rounded-2xl border border-slate-200/90 bg-white/95 backdrop-blur-xl shadow-2xl p-2 animate-scale-in space-y-2 max-h-80 flex flex-col">
          {/* Live Search Input */}
          <div className="relative flex-shrink-0">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <input
              ref={inputRef}
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder={searchPlaceholder}
              className="w-full h-9 pl-9 pr-8 text-xs font-medium bg-slate-50 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-indigo-500/30 text-slate-800 placeholder:text-slate-400"
            />
            {search && (
              <button
                type="button"
                onClick={() => setSearch('')}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-0.5 rounded-full hover:bg-slate-200/60"
              >
                <X size={12} />
              </button>
            )}
          </div>

          {/* Options Scrollable Container */}
          <div className="flex-1 overflow-y-auto max-h-60 space-y-0.5 pr-1">
            {filtered.length === 0 ? (
              <div className="py-6 text-center text-xs text-slate-400 font-medium">
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
                      'w-full flex items-center justify-between rounded-xl px-3 py-2 text-left text-xs font-semibold transition-colors duration-150',
                      isSelected
                        ? 'bg-indigo-600 text-white shadow-sm shadow-indigo-500/20'
                        : 'text-slate-700 hover:bg-indigo-50/60 hover:text-indigo-900'
                    )}
                  >
                    <div className="truncate pr-2">
                      <p className="truncate font-bold">{opt.label}</p>
                      {opt.subtitle && (
                        <p className={cn('truncate text-[10px] font-normal', isSelected ? 'text-indigo-100' : 'text-slate-400')}>
                          {opt.subtitle}
                        </p>
                      )}
                    </div>
                    {isSelected && <Check size={14} className="flex-shrink-0 text-white" />}
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
