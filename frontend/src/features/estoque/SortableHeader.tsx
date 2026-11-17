import { ArrowDown, ArrowUp, ArrowUpDown } from 'lucide-react'
import { cn } from '@/lib/utils'
import type { EstoqueSort, SortDir } from './api'

interface SortableHeaderProps {
  label: string
  coluna: EstoqueSort
  sort: EstoqueSort
  dir: SortDir
  onSort: (coluna: EstoqueSort) => void
  className?: string
}

/** `<th>` com botão de ordenação; `aria-sort` reflete a ordenação ativa. */
export function SortableHeader({ label, coluna, sort, dir, onSort, className }: SortableHeaderProps) {
  const ativa = sort === coluna
  const Icone = ativa ? (dir === 'asc' ? ArrowUp : ArrowDown) : ArrowUpDown
  return (
    <th
      scope="col"
      aria-sort={ativa ? (dir === 'asc' ? 'ascending' : 'descending') : 'none'}
      className={cn('px-4 py-3 text-caption font-semibold', className)}
    >
      <button
        type="button"
        onClick={() => onSort(coluna)}
        className="inline-flex items-center gap-1.5 rounded text-left font-semibold hover:underline focus-visible:ring-1 focus-visible:ring-ring"
      >
        {label}
        <Icone aria-hidden="true" className={cn('size-3.5 shrink-0', !ativa && 'opacity-60')} />
      </button>
    </th>
  )
}
