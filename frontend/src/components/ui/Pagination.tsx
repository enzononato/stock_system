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
  const from = total === 0 ? 0 : (page - 1) * pageSize + 1
  const to = Math.min(page * pageSize, total)
  const [goTo, setGoTo] = React.useState('')

  const confirmGoTo = () => {
    const n = Number(goTo.trim())
    // Texto inválido não muda nada; número fora da faixa é corrigido para [1, última].
    if (goTo.trim() === '' || !Number.isFinite(n)) return
    onPageChange(Math.min(lastPage, Math.max(1, Math.trunc(n))))
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
            disabled={page <= 1}
            onClick={() => onPageChange(page - 1)}
          >
            <ChevronLeft aria-hidden="true" />
          </Button>
          <span className="px-2 tabular-nums">
            {page} / {lastPage}
          </span>
          <Button
            type="button"
            variant="outline"
            size="icon"
            className="h-8 w-8"
            aria-label="Próxima"
            disabled={page >= lastPage}
            onClick={() => onPageChange(page + 1)}
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
