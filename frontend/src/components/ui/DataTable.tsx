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
import { ArrowUpDown, ChevronLeft, ChevronRight } from 'lucide-react'

/**
 * Paginação server-side opcional do `DataTable` (ver `DataTableProps.pagination`).
 * Quando fornecida, `data` deve conter apenas os registros da página atual
 * (ex.: o `items` de `listHistoryPaginated({ limit, offset })`), e `total` é a
 * contagem real no servidor — não `data.length`.
 */
export interface DataTablePaginationProps {
  /** Total de registros no servidor (não apenas os da página atual). */
  total: number
  /** Página atual, base 0. */
  pageIndex: number
  /** Quantidade de registros por página. */
  pageSize: number
  /** Chamado com o novo `pageIndex` quando o usuário navega (Anterior/Próxima). */
  onPageChange: (index: number) => void
}

interface DataTableProps<TData> {
  data: TData[]
  columns: ColumnDef<TData, unknown>[]
  searchPlaceholder?: string
  className?: string
  /**
   * Ausente (padrão): comportamento inalterado — client-side, `data` é a lista
   * inteira, filtro/ordenação rodam no navegador e o rodapé mostra "N de M
   * registros" com base em `data.length`. Nenhum chamador existente precisa
   * mudar.
   *
   * Presente: assume que `data` é só a página atual vinda do servidor. Desliga
   * o rodapé de contagem client-side e renderiza controles de navegação
   * (Anterior/Próxima + "Página X de Y") usando `total`/`pageIndex`/`pageSize`.
   * A busca por texto livre também é desativada nesse modo: sem um parâmetro
   * de busca no backend, filtrar no cliente só filtraria a página carregada,
   * o que enganaria o usuário (pareceria buscar em tudo, mas ignora as demais
   * páginas). Ordenação por coluna continua ativa, mas também só age sobre a
   * página atual.
   */
  pagination?: DataTablePaginationProps
}

export function DataTable<TData>({
  data,
  columns,
  searchPlaceholder = 'Buscar...',
  className,
  pagination,
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

  // Controles de paginação server-side (só fazem sentido quando `pagination` é passado).
  const pageSize = pagination?.pageSize || 1
  const pageCount = pagination ? Math.max(1, Math.ceil(pagination.total / pageSize)) : 0
  const currentPage = (pagination?.pageIndex ?? 0) + 1
  const rangeStart = pagination && pagination.total > 0 ? pagination.pageIndex * pageSize + 1 : 0
  const rangeEnd = pagination ? Math.min(pagination.total, (pagination.pageIndex + 1) * pageSize) : 0

  return (
    <div className={cn('space-y-3', className)}>
      {!pagination && (
        <Input
          placeholder={searchPlaceholder}
          value={globalFilter}
          onChange={(e) => setGlobalFilter(e.target.value)}
          className="max-w-sm"
        />
      )}
      <div className="rounded-md border overflow-x-auto">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 border-b">
            {table.getHeaderGroups().map((headerGroup) => (
              <tr key={headerGroup.id}>
                {headerGroup.headers.map((header) => (
                  <th
                    key={header.id}
                    className="px-4 py-3 text-left font-medium text-slate-600 whitespace-nowrap select-none"
                    onClick={header.column.getToggleSortingHandler()}
                    style={{ cursor: header.column.getCanSort() ? 'pointer' : 'default' }}
                  >
                    <div className="flex items-center gap-1">
                      {flexRender(header.column.columnDef.header, header.getContext())}
                      {header.column.getCanSort() && <ArrowUpDown size={12} className="opacity-40" />}
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody className="divide-y divide-slate-100">
            {rows.length === 0 ? (
              <tr>
                <td colSpan={columns.length} className="px-4 py-8 text-center text-slate-400">
                  Nenhum resultado encontrado.
                </td>
              </tr>
            ) : (
              rows.map((row) => (
                <tr key={row.id} className="hover:bg-slate-50 transition-colors">
                  {row.getVisibleCells().map((cell) => (
                    <td key={cell.id} className="px-4 py-2.5 whitespace-nowrap">
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {pagination ? (
        <div className="flex items-center justify-between gap-4">
          <p className="text-xs text-slate-400 whitespace-nowrap">
            {pagination.total === 0 ? '0 registros' : `${rangeStart}–${rangeEnd} de ${pagination.total} registros`}
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
            <span className="text-xs text-slate-500 px-1 whitespace-nowrap">
              Página {currentPage} de {pageCount}
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
        <p className="text-xs text-slate-400">
          {table.getFilteredRowModel().rows.length} de {data.length} registros
        </p>
      )}
    </div>
  )
}
