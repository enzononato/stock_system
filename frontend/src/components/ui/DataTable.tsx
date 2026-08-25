import { useState } from 'react'
import {
  useReactTable,
  getCoreRowModel,
  getFilteredRowModel,
  getSortedRowModel,
  flexRender,
  type ColumnDef,
  type SortingState,
} from '@tanstack/react-table'
import { Input } from './input'
import { Button } from './button'
import { cn } from '@/lib/utils'
import { ArrowUpDown, ChevronLeft, ChevronRight, Search, X } from 'lucide-react'

export interface DataTablePaginationProps {
  total: number
  pageIndex: number
  pageSize: number
  onPageChange: (index: number) => void
  search?: string
  onSearchChange?: (termo: string) => void
}

interface DataTableProps<TData> {
  data: TData[]
  columns: ColumnDef<TData, unknown>[]
  searchPlaceholder?: string
  className?: string
  pagination?: DataTablePaginationProps
  onRowClick?: (row: TData) => void
}

export function DataTable<TData>({
  data,
  columns,
  searchPlaceholder = 'Buscar...',
  className,
  pagination,
  onRowClick,
}: DataTableProps<TData>) {
  const [globalFilter, setGlobalFilter] = useState('')
  const [sorting, setSorting] = useState<SortingState>([])

  const table = useReactTable({
    data,
    columns,
    state: { globalFilter, sorting },
    onGlobalFilterChange: setGlobalFilter,
    onSortingChange: setSorting,
    getCoreRowModel: getCoreRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getSortedRowModel: getSortedRowModel(),
  })

  const rows = table.getRowModel().rows

  const pageSize = pagination?.pageSize || 1
  const pageCount = pagination ? Math.max(1, Math.ceil(pagination.total / pageSize)) : 0
  const currentPage = (pagination?.pageIndex ?? 0) + 1
  const rangeStart = pagination && pagination.total > 0 ? pagination.pageIndex * pageSize + 1 : 0
  const rangeEnd = pagination ? Math.min(pagination.total, (pagination.pageIndex + 1) * pageSize) : 0

  const currentValue = pagination ? (pagination.search ?? '') : globalFilter
  const handleSearchInput = (val: string) => {
    if (pagination?.onSearchChange) {
      pagination.onSearchChange(val)
    } else {
      setGlobalFilter(val)
    }
  }

  const showSearchInput = pagination ? Boolean(pagination.onSearchChange) : true

  return (
    <div className={cn('surface rounded-lg p-4 space-y-4', className)}>
      {/* Search Header Bar */}
      {showSearchInput && (
        <div className="flex items-center justify-between gap-4">
          <div className="relative max-w-md w-full">
            <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none" />
            <Input
              placeholder={searchPlaceholder}
              value={currentValue}
              onChange={(e) => handleSearchInput(e.target.value)}
              className="pl-10 pr-9 bg-secondary/50 border-border focus-visible:ring-primary/30 rounded-md"
            />
            {currentValue && (
              <button
                onClick={() => handleSearchInput('')}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground p-0.5 rounded-full hover:bg-accent"
              >
                <X size={14} />
              </button>
            )}
          </div>
        </div>
      )}

      {/* Table Container */}
      <div className="rounded-lg border border-border overflow-hidden bg-card">
        {rows.length === 0 ? (
          // Rendered outside the horizontally-scrollable table: with many
          // columns the table can be wider than the viewport, which pushed
          // this centered content off to one side. A full-width block here
          // centers correctly regardless of column count.
          <div className="flex flex-col items-center justify-center gap-2 px-4 py-16 text-center text-muted-foreground">
            <div className="h-10 w-10 rounded-full bg-secondary flex items-center justify-center text-muted-foreground">
              <Search size={20} />
            </div>
            <p className="text-sm font-medium text-foreground">Nenhum resultado encontrado.</p>
            <p className="text-xs text-muted-foreground">Tente ajustar seus termos de busca ou filtros.</p>
          </div>
        ) : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-secondary/40 border-b border-border">
              {table.getHeaderGroups().map((headerGroup) => (
                <tr key={headerGroup.id}>
                  {headerGroup.headers.map((header) => (
                    <th
                      key={header.id}
                      className="px-4 py-3.5 text-left text-[10px] font-semibold uppercase tracking-wider text-muted-foreground whitespace-nowrap select-none"
                      onClick={header.column.getToggleSortingHandler()}
                      style={{ cursor: header.column.getCanSort() ? 'pointer' : 'default' }}
                    >
                      <div className="flex items-center gap-1.5 hover:text-primary transition-colors">
                        {flexRender(header.column.columnDef.header, header.getContext())}
                        {header.column.getCanSort() && <ArrowUpDown size={12} className="opacity-50" />}
                      </div>
                    </th>
                  ))}
                </tr>
              ))}
            </thead>
            <tbody className="divide-y divide-border/60">
              {rows.map((row) => (
                <tr
                  key={row.id}
                  onClick={(e) => {
                    const target = e.target as HTMLElement
                    if (target.closest('button, a, input, select, [role="button"]')) {
                      return
                    }
                    onRowClick?.(row.original)
                  }}
                  className={cn(
                    'hover:bg-accent/60 transition-colors duration-150 group',
                    onRowClick && 'cursor-pointer'
                  )}
                >
                  {row.getVisibleCells().map((cell) => (
                    <td key={cell.id} className="px-4 py-2.5 whitespace-nowrap text-foreground/90 text-xs">
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        )}
      </div>

      {/* Pagination Footer */}
      {pagination ? (
        <div className="grid grid-cols-[1fr_auto_1fr] items-center gap-4 pt-1">
          <p className="text-xs font-medium text-muted-foreground whitespace-nowrap">
            {pagination.total === 0 ? '0 registros' : `${rangeStart}–${rangeEnd} de ${pagination.total} registros`}
          </p>
          <div className="flex items-center gap-2 justify-self-center">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => pagination.onPageChange(pagination.pageIndex - 1)}
              disabled={pagination.pageIndex <= 0}
              className="rounded-lg text-xs"
            >
              <ChevronLeft size={14} className="mr-1" />Anterior
            </Button>
            <span className="text-xs font-semibold text-foreground px-2 py-1 rounded-md bg-secondary whitespace-nowrap">
              Página {currentPage} de {pageCount}
            </span>
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => pagination.onPageChange(pagination.pageIndex + 1)}
              disabled={pagination.pageIndex + 1 >= pageCount}
              className="rounded-lg text-xs"
            >
              Próxima<ChevronRight size={14} className="ml-1" />
            </Button>
          </div>
          <div aria-hidden />
        </div>
      ) : (
        <div className="flex items-center justify-between pt-1">
          <p className="text-xs font-medium text-muted-foreground">
            {table.getFilteredRowModel().rows.length} de {data.length} registros
          </p>
        </div>
      )}
    </div>
  )
}

