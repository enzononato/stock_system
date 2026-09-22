import { useEffect, useRef, useState } from 'react'
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

// Permite que cada coluna declare classes extras para sua <th>/<td> — usado
// pelas telas de alta densidade (ex.: StockPage) para esconder colunas por
// breakpoint (`hidden md:table-cell` etc.) sem duplicar a tabela numa versão
// escrita à mão.
declare module '@tanstack/react-table' {
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  interface ColumnMeta<TData, TValue> {
    className?: string
  }
}

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
  /** Tamanho de página da paginação client-side (só entra em ação quando `pagination` não é passada). Padrão: 10. */
  clientPageSize?: number
}

export function DataTable<TData>({
  data,
  columns,
  searchPlaceholder = 'Buscar...',
  className,
  pagination,
  onRowClick,
  clientPageSize = 10,
}: DataTableProps<TData>) {
  const [globalFilter, setGlobalFilter] = useState('')
  const [sorting, setSorting] = useState<SortingState>([])
  // Página atual da paginação client-side (só usada quando `pagination` não é passada).
  const [clientPage, setClientPage] = useState(0)

  // Se `data` mudar de tamanho por fora (item criado/removido, recarregado
  // etc.), a página em que o usuário está pode deixar de existir — por
  // exemplo, a lista encolhe de 5 para 2 páginas enquanto ele está na página
  // 5, e a tabela renderizaria vazia sem nenhuma pista do motivo. Volta à
  // primeira página sempre que o tamanho do array mudar.
  const previousDataLength = useRef(data.length)
  useEffect(() => {
    if (previousDataLength.current !== data.length) {
      previousDataLength.current = data.length
      setClientPage(0)
    }
  }, [data.length])

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

  const allRows = table.getRowModel().rows

  // Paginação client-side: fatia as linhas já filtradas/ordenadas em páginas.
  // Só entra em ação quando ninguém passou a prop `pagination` (caminho
  // server-side, usado por HistoryPage/StockPage/ReportPage etc., que
  // continua intocado abaixo).
  const isClientPaginated = !pagination
  const clientPageCount = isClientPaginated ? Math.max(1, Math.ceil(allRows.length / clientPageSize)) : 0
  const rows = isClientPaginated
    ? allRows.slice(clientPage * clientPageSize, (clientPage + 1) * clientPageSize)
    : allRows

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
      setClientPage(0) // nova busca sempre volta para a primeira página
    }
  }

  const showSearchInput = pagination ? Boolean(pagination.onSearchChange) : true

  return (
    <div className={cn('space-y-4', className)}>
      {/*
       * Busca em painel próprio (toolbar), separado do painel da tabela — o
       * mesmo padrão da referência: a busca não compartilha borda com os
       * dados, e a tabela fica livre para correr de ponta a ponta no painel
       * abaixo. Sem caption de PanelHeader aqui de propósito: toda página que
       * chama DataTable já renderiza um PanelHeader ou <h3> imediatamente
       * acima (ex.: "Empréstimos Ativos (N)", "Unidades cadastradas") — uma
       * legenda aqui duplicaria essa faixa.
       */}
      {showSearchInput && (
        <div className="surface-panel p-3">
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

      {/* Painel da tabela: overflow-hidden para o cabeçalho (bg-surface-alt) chegar até a borda do painel. */}
      <div className="surface-panel overflow-hidden">
        {/* overflow-x-auto fica só ao redor de <table> — preserva a correção
            anterior de tabela larga rolar sem arrastar busca/paginação para fora da tela. */}
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-surface-alt border-b border-border">
              {table.getHeaderGroups().map((headerGroup) => (
                <tr key={headerGroup.id}>
                  {headerGroup.headers.map((header) => (
                    <th
                      key={header.id}
                      className={cn(
                        'text-caption text-muted-foreground px-4 py-2.5 text-left whitespace-nowrap select-none',
                        header.column.columnDef.meta?.className
                      )}
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
                      // `group`: permite que células desta linha (ex.: as ações
                      // que só aparecem no hover) reajam ao hover da linha
                      // inteira, não só do próprio elemento.
                      'group border-b border-border last:border-0 hover:bg-surface-alt transition-colors duration-micro',
                      onRowClick && 'cursor-pointer'
                    )}
                  >
                    {row.getVisibleCells().map((cell) => (
                      <td
                        key={cell.id}
                        className={cn(
                          'px-4 py-2.5 text-body-sm whitespace-nowrap',
                          cell.column.columnDef.meta?.className
                        )}
                      >
                        {flexRender(cell.column.columnDef.cell, cell.getContext())}
                      </td>
                    ))}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination Footer — dentro do painel da tabela, com padding próprio já que o painel não tem mais p-5. */}
        {pagination ? (
          <div className="flex items-center justify-between gap-4 px-4 py-3 border-t border-border">
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
                <ChevronLeft size={14} />Anterior
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
                Próxima<ChevronRight size={14} />
              </Button>
            </div>
          </div>
        ) : (
          // Paginação client-side — o contador aparece sempre; os botões de
          // navegação só quando há mais de uma página, para não poluir listas
          // curtas que já cabem inteiras na tela.
          <div className="flex items-center justify-between gap-4 px-4 py-3 border-t border-border">
            <p className="text-caption text-muted-foreground whitespace-nowrap">
              <span className="num">{allRows.length}</span> de{' '}
              <span className="num">{data.length}</span> registros
            </p>
            {clientPageCount > 1 && (
              <div className="flex items-center gap-2">
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onClick={() => setClientPage((p) => Math.max(0, p - 1))}
                  disabled={clientPage <= 0}
                >
                  <ChevronLeft size={14} />Anterior
                </Button>
                <span className="text-caption text-muted-foreground px-2 py-1 rounded-md bg-surface-alt whitespace-nowrap">
                  Página <span className="num">{clientPage + 1}</span> de <span className="num">{clientPageCount}</span>
                </span>
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onClick={() => setClientPage((p) => Math.min(clientPageCount - 1, p + 1))}
                  disabled={clientPage + 1 >= clientPageCount}
                >
                  Próxima<ChevronRight size={14} />
                </Button>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  )
}

