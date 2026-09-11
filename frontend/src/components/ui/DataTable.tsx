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
    <div className={cn('surface-panel p-5 space-y-4', className)}>
      {/* Search Header Bar */}
      {showSearchInput && (
        <div className="flex items-center justify-between gap-4">
          <div className="relative max-w-md w-full">
            <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none" />
            <Input
              placeholder={searchPlaceholder}
              value={currentValue}
              onChange={(e) => handleSearchInput(e.target.value)}
              className="pl-10 pr-9"
            />
            {currentValue && (
              <button
                onClick={() => handleSearchInput('')}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground p-0.5 rounded-sm hover:bg-surface-alt"
              >
                <X size={14} />
              </button>
            )}
          </div>
        </div>
      )}

      {/* Table */}
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-surface-alt border-b border-border">
            {table.getHeaderGroups().map((headerGroup) => (
              <tr key={headerGroup.id}>
                {headerGroup.headers.map((header) => (
                  <th
                    key={header.id}
                    className="text-caption text-muted-foreground px-4 py-2.5 text-left whitespace-nowrap select-none"
                    onClick={header.column.getToggleSortingHandler()}
                    style={{ cursor: header.column.getCanSort() ? 'pointer' : 'default' }}
                  >
                    <div className="flex items-center gap-1.5 hover:text-foreground transition-colors">
                      {flexRender(header.column.columnDef.header, header.getContext())}
                      {header.column.getCanSort() && <ArrowUpDown size={12} className="opacity-50" />}
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={columns.length} className="text-body-sm text-muted-foreground py-10 text-center">
                  <p>Nenhum resultado encontrado.</p>
                  <p className="mt-1">Tente ajustar seus termos de busca ou filtros.</p>
                </td>
              </tr>
            ) : (
              rows.map((row) => (
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
                    'border-b border-border last:border-0 hover:bg-surface-alt transition-colors duration-micro',
                    onRowClick && 'cursor-pointer'
                  )}
                >
                  {row.getVisibleCells().map((cell) => (
                    <td key={cell.id} className="px-4 py-2.5 text-body-sm whitespace-nowrap">
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination Footer */}
      {pagination ? (
        <div className="flex items-center justify-between gap-4 pt-1">
          <p className="text-caption text-muted-foreground whitespace-nowrap">
            {pagination.total === 0 ? (
              <>
                <span className="num">0</span> registros
              </>
            ) : (
              <>
                <span className="num">{rangeStart}</span>–<span className="num">{rangeEnd}</span> de{' '}
                <span className="num">{pagination.total}</span> registros
              </>
            )}
          </p>
          <div className="flex items-center gap-2">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => pagination.onPageChange(pagination.pageIndex - 1)}
              disabled={pagination.pageIndex <= 0}
            >
              <ChevronLeft size={14} className="mr-1" />Anterior
            </Button>
            <span className="text-caption text-muted-foreground px-2 py-1 rounded-md bg-surface-alt whitespace-nowrap">
              Página <span className="num">{currentPage}</span> de <span className="num">{pageCount}</span>
            </span>
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => pagination.onPageChange(pagination.pageIndex + 1)}
              disabled={pagination.pageIndex + 1 >= pageCount}
            >
              Próxima<ChevronRight size={14} className="ml-1" />
            </Button>
          </div>
        </div>
      ) : (
        <div className="flex items-center justify-between pt-1">
          <p className="text-caption text-muted-foreground">
            <span className="num">{table.getFilteredRowModel().rows.length}</span> de{' '}
            <span className="num">{data.length}</span> registros
          </p>
        </div>
      )}
    </div>
  )
}

