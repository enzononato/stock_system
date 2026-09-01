import { useMemo, useState, type ReactNode } from "react";
import { ArrowDown, ArrowUp, ChevronsUpDown, ChevronLeft, ChevronRight } from "lucide-react";

import { cn } from "@/lib/utils";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import { EmptyState, ErrorState } from "@/components/app/StateBlocks";

export interface Column<T> {
  /** Identificador único da coluna. */
  key: string;
  header: string;
  cell: (row: T) => ReactNode;
  /** Valor usado na ordenação client-side (quando `sortable`). */
  sortValue?: ((row: T) => string | number | undefined) | undefined;
  sortable?: boolean | undefined;
  align?: ("left" | "right") | undefined;
  /** Esconde a coluna abaixo do breakpoint informado (desktop-first). */
  hideBelow?: ("md" | "lg" | "xl") | undefined;
  /** Título do card no layout mobile. */
  primary?: boolean | undefined;
  /** Não exibir no layout mobile (ex.: coluna de ações já renderizada no rodapé do card). */
  hideOnMobile?: boolean | undefined;
}

interface DataTableProps<T> {
  data: T[];
  columns: Column<T>[];
  rowKey: (row: T) => string | number;
  isLoading?: boolean | undefined;
  error?: unknown;
  onRetry?: (() => void) | undefined;
  emptyTitle?: string | undefined;
  emptyDescription?: string | undefined;
  emptyAction?: ReactNode | undefined;
  /** Ações renderizadas no rodapé de cada card no mobile. */
  mobileActions?: ((row: T) => ReactNode) | undefined;
  onRowClick?: ((row: T) => void) | undefined;
  caption?: string | undefined;
  /** Paginação server-side. Quando ausente, a tabela renderiza tudo que recebe. */
  pagination?:
    | {
        page: number;
        pageSize: number;
        total: number;
        onPageChange: (page: number) => void;
      }
    | undefined;
}

const HIDE_CLASS: Record<NonNullable<Column<unknown>["hideBelow"]>, string> = {
  md: "hidden md:table-cell",
  lg: "hidden lg:table-cell",
  xl: "hidden xl:table-cell",
};

export function DataTable<T>({
  data,
  columns,
  rowKey,
  isLoading,
  error,
  onRetry,
  emptyTitle = "Nenhum registro encontrado",
  emptyDescription = "Ajuste os filtros ou cadastre um novo registro.",
  emptyAction,
  mobileActions,
  onRowClick,
  caption,
  pagination,
}: DataTableProps<T>) {
  const [sort, setSort] = useState<{ key: string; dir: "asc" | "desc" } | null>(null);

  const rows = useMemo(() => {
    if (!sort) return data;
    const col = columns.find((c) => c.key === sort.key);
    if (!col?.sortValue) return data;
    const factor = sort.dir === "asc" ? 1 : -1;
    return [...data].sort((a, b) => {
      const av = col.sortValue!(a);
      const bv = col.sortValue!(b);
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      if (typeof av === "number" && typeof bv === "number") return (av - bv) * factor;
      return String(av).localeCompare(String(bv), "pt-BR") * factor;
    });
  }, [data, sort, columns]);

  function toggleSort(key: string) {
    setSort((prev) =>
      prev?.key === key ? (prev.dir === "asc" ? { key, dir: "desc" } : null) : { key, dir: "asc" },
    );
  }

  if (error) {
    return <ErrorState error={error} onRetry={onRetry} />;
  }

  if (isLoading) {
    return (
      <div className="surface-panel overflow-hidden" aria-busy="true" aria-live="polite">
        <div className="space-y-px">
          {Array.from({ length: 6 }).map((_, i) => (
            <div key={i} className="flex items-center gap-4 px-4 py-3.5">
              <Skeleton className="h-4 w-10" />
              <Skeleton className="h-4 flex-1" />
              <Skeleton className="hidden h-4 w-32 md:block" />
              <Skeleton className="hidden h-4 w-24 lg:block" />
              <Skeleton className="h-4 w-16" />
            </div>
          ))}
        </div>
        <span className="sr-only">Carregando registros…</span>
      </div>
    );
  }

  if (rows.length === 0) {
    return <EmptyState title={emptyTitle} description={emptyDescription} action={emptyAction} />;
  }

  const totalPages = pagination ? Math.max(1, Math.ceil(pagination.total / pagination.pageSize)) : 1;

  return (
    <div className="space-y-3">
      {/* Desktop / tablet */}
      <div className="surface-panel hidden overflow-hidden sm:block">
        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-sm">
            {caption ? <caption className="sr-only">{caption}</caption> : null}
            <thead>
              <tr className="border-b border-border bg-muted/40">
                {columns.map((col) => (
                  <th
                    key={col.key}
                    scope="col"
                    aria-sort={
                      sort?.key === col.key ? (sort.dir === "asc" ? "ascending" : "descending") : undefined
                    }
                    className={cn(
                      "px-4 py-2.5 text-left text-[11px] font-semibold uppercase tracking-wider text-muted-foreground whitespace-nowrap",
                      col.align === "right" && "text-right",
                      col.hideBelow && HIDE_CLASS[col.hideBelow],
                    )}
                  >
                    {col.sortable && col.sortValue ? (
                      <button
                        type="button"
                        onClick={() => toggleSort(col.key)}
                        className="inline-flex items-center gap-1 rounded-sm transition-colors hover:text-foreground"
                      >
                        {col.header}
                        {sort?.key === col.key ? (
                          sort.dir === "asc" ? (
                            <ArrowUp className="size-3" aria-hidden />
                          ) : (
                            <ArrowDown className="size-3" aria-hidden />
                          )
                        ) : (
                          <ChevronsUpDown className="size-3 opacity-50" aria-hidden />
                        )}
                      </button>
                    ) : (
                      col.header
                    )}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr
                  key={rowKey(row)}
                  onClick={onRowClick ? () => onRowClick(row) : undefined}
                  className={cn(
                    "border-b border-border/70 last:border-0 transition-colors hover:bg-accent/40",
                    onRowClick && "cursor-pointer",
                  )}
                >
                  {columns.map((col) => (
                    <td
                      key={col.key}
                      className={cn(
                        "px-4 py-3 align-middle text-foreground",
                        col.align === "right" && "text-right",
                        col.hideBelow && HIDE_CLASS[col.hideBelow],
                      )}
                    >
                      {col.cell(row)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Mobile — cada linha vira um card */}
      <ul className="space-y-2.5 sm:hidden">
        {rows.map((row) => {
          const primary = columns.find((c) => c.primary);
          const rest = columns.filter((c) => !c.primary && !c.hideOnMobile);
          return (
            <li key={rowKey(row)} className="surface-panel p-4">
              {primary ? <div className="mb-2 text-sm font-semibold">{primary.cell(row)}</div> : null}
              <dl className="grid grid-cols-2 gap-x-3 gap-y-2">
                {rest.map((col) => (
                  <div key={col.key} className="min-w-0">
                    <dt className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
                      {col.header}
                    </dt>
                    <dd className="truncate text-sm text-foreground">{col.cell(row)}</dd>
                  </div>
                ))}
              </dl>
              {mobileActions ? (
                <div className="mt-3 flex flex-wrap gap-2 border-t border-border pt-3">{mobileActions(row)}</div>
              ) : null}
            </li>
          );
        })}
      </ul>

      {pagination ? (
        <nav
          aria-label="Paginação"
          className="flex flex-col items-center justify-between gap-3 sm:flex-row"
        >
          <p className="text-xs text-muted-foreground">
            Página <span className="num">{pagination.page + 1}</span> de{" "}
            <span className="num">{totalPages}</span> · <span className="num">{pagination.total}</span> registros
          </p>
          <div className="flex gap-2">
            <Button
              variant="outline"
              size="sm"
              disabled={pagination.page === 0}
              onClick={() => pagination.onPageChange(pagination.page - 1)}
            >
              <ChevronLeft className="size-4" aria-hidden />
              Anterior
            </Button>
            <Button
              variant="outline"
              size="sm"
              disabled={pagination.page + 1 >= totalPages}
              onClick={() => pagination.onPageChange(pagination.page + 1)}
            >
              Próxima
              <ChevronRight className="size-4" aria-hidden />
            </Button>
          </div>
        </nav>
      ) : null}
    </div>
  );
}
