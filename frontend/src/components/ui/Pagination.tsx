import * as React from 'react'
import { ChevronLeft, ChevronRight } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { cn } from '@/lib/utils'

interface PaginationProps {
  /** Página atual, começando em 1. */
  page: number
  pageSize: number
  total: number
  onPageChange: (page: number) => void
  onPageSizeChange?: (size: number) => void
  pageSizeOptions?: number[]
  isFetching?: boolean
}

export function Pagination({
  page,
  pageSize,
  total,
  onPageChange,
  onPageSizeChange,
  pageSizeOptions = [25, 50, 100],
  isFetching,
}: PaginationProps) {
  const lastPage = Math.max(1, Math.ceil(total / pageSize))
  // Página fora da faixa (total encolheu) nunca produz "126–120": mostra a última válida.
  const atual = Math.min(Math.max(1, page), lastPage)
  const from = total === 0 ? 0 : (atual - 1) * pageSize + 1
  const to = Math.min(atual * pageSize, total)
  const [goTo, setGoTo] = React.useState('')

  const confirmGoTo = () => {
    const texto = goTo.trim()
    // Só dígitos: "1e2", "2.5", "-1" e "0x3" não mudam nada; número fora da faixa é corrigido para [1, última].
    if (!/^\d+$/.test(texto)) return
    onPageChange(Math.min(lastPage, Math.max(1, Number(texto))))
    setGoTo('')
  }

  return (
    <nav
      aria-label="Paginação"
      aria-busy={isFetching ? true : undefined}
      className={cn(
        'flex flex-wrap items-center justify-between gap-3 text-sm text-muted-foreground transition-opacity duration-micro motion-reduce:transition-none',
        isFetching && 'opacity-70'
      )}
    >
      <span aria-live="polite">{`Mostrando ${from}–${to} de ${total}`}</span>

      <div className="flex flex-wrap items-center gap-3">
        {onPageSizeChange && (
          <label className="flex items-center gap-2">
            <span>Itens por página</span>
            <select
              aria-label="Itens por página"
              value={pageSize}
              onChange={(e) => onPageSizeChange(Number(e.target.value))}
              className="h-8 rounded border border-input bg-surface-alt px-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring"
            >
              {pageSizeOptions.map((o) => (
                <option key={o} value={o}>
                  {o}
                </option>
              ))}
            </select>
          </label>
        )}

        <div className="flex items-center gap-1">
          <Button
            type="button"
            variant="outline"
            size="icon"
            className="h-8 w-8"
            aria-label="Anterior"
            disabled={atual <= 1}
            onClick={() => onPageChange(atual - 1)}
          >
            <ChevronLeft aria-hidden="true" />
          </Button>
          <span className="px-2 tabular-nums">
            {atual} / {lastPage}
          </span>
          <Button
            type="button"
            variant="outline"
            size="icon"
            className="h-8 w-8"
            aria-label="Próxima"
            disabled={atual >= lastPage}
            onClick={() => onPageChange(atual + 1)}
          >
            <ChevronRight aria-hidden="true" />
          </Button>
        </div>

        <div className="flex items-center gap-1">
          <input
            type="text"
            inputMode="numeric"
            aria-label="Ir para a página"
            placeholder="Ir para"
            value={goTo}
            onChange={(e) => setGoTo(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault()
                confirmGoTo()
              }
            }}
            className="h-8 w-20 rounded border border-input bg-surface-alt px-2 text-sm text-foreground placeholder:text-muted-foreground focus-visible:ring-1 focus-visible:ring-ring"
          />
          <Button type="button" variant="outline" size="sm" aria-label="Ir" onClick={confirmGoTo}>
            Ir
          </Button>
        </div>
      </div>
    </nav>
  )
}
