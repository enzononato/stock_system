import { useEffect, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { useNavigate } from "@tanstack/react-router";
import {
  Boxes,
  Eye,
  Filter,
  Pencil,
  Plus,
  RefreshCw,
  Search,
  X,
} from "lucide-react";

import { listItemsPaginated, type Item } from "@/api/items";
import { listUnidades } from "@/api/unidades";
import { DataTable, type Column } from "@/components/app/DataTable";
import { AssetDetailsPanel } from "@/components/app/AssetDetailsPanel";
import { EditItemModal } from "@/components/app/EditItemModal";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { useAuth } from "@/lib/auth";
import { useConstants } from "@/hooks/useConstants";
import { cn, formatDate } from "@/lib/utils";

const STATUS_OPTIONS = ["Disponível", "Indisponível", "Pendente", "Pendente Devolução"];
const PAGE_SIZE = 7;

function MonochromaticStatusBadge({ status }: { status?: string }) {
  if (!status) return <span className="text-muted-foreground text-xs">—</span>;

  let symbol = "●";
  let label = status;

  if (status === "Disponível") {
    symbol = "○";
  } else if (status.startsWith("Pendente")) {
    symbol = "!";
  } else if (status === "Baixado" || status === "Descartado") {
    symbol = "×";
  }

  return (
    <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-sm border border-border text-foreground font-mono text-xs font-medium">
      <span className="font-bold text-[11px]">{symbol}</span>
      <span className="uppercase tracking-wider text-[11px]">{label}</span>
    </span>
  );
}

export function StockPage() {
  const { hasRole } = useAuth();
  const navigate = useNavigate();
  const { equipmentTypes, isLoading: constantsLoading } = useConstants();

  const [search, setSearch] = useState("");
  const [filterTipo, setFilterTipo] = useState("all");
  const [filterStatus, setFilterStatus] = useState("all");
  const [filterRevenda, setFilterRevenda] = useState("all");
  const [page, setPage] = useState(0);

  const [selectedAsset, setSelectedAsset] = useState<Item | null>(null);
  const [editingItem, setEditingItem] = useState<Item | null>(null);

  const { data: unidades = [] } = useQuery({
    queryKey: ["unidades-filter"],
    queryFn: () => listUnidades(),
  });

  const { data, isLoading, isFetching, error, refetch } = useQuery({
    queryKey: ["items", filterTipo, filterStatus, filterRevenda, page, search],
    queryFn: () =>
      listItemsPaginated({
        tipo: filterTipo !== "all" ? filterTipo : undefined,
        status: filterStatus !== "all" ? filterStatus : undefined,
        revenda: filterRevenda !== "all" ? filterRevenda : undefined,
        search: search.trim() || undefined,
        limit: PAGE_SIZE,
        offset: page * PAGE_SIZE,
      }),
  });

  const items = data?.items ?? [];
  const total = data?.total ?? 0;
  const hasFilters =
    filterTipo !== "all" || filterStatus !== "all" || filterRevenda !== "all" || search !== "";

  useEffect(() => {
    if (data && page > 0 && items.length === 0 && total > 0) {
      setPage(Math.max(0, Math.ceil(total / PAGE_SIZE) - 1));
    }
  }, [data, items.length, page, total]);

  const resetFilters = () => {
    setSearch("");
    setFilterTipo("all");
    setFilterStatus("all");
    setFilterRevenda("all");
    setPage(0);
  };

  const columns: Column<Item>[] = [
    {
      key: "id",
      header: "Patrimônio",
      primary: true,
      cell: (row) => (
        <button
          type="button"
          onClick={() => setSelectedAsset(row)}
          className="num font-mono text-xs font-semibold text-foreground hover:underline cursor-pointer flex items-center gap-1"
        >
          <span>#{row.id}</span>
          {row.identificador && (
            <span className="text-muted-foreground font-normal">({row.identificador})</span>
          )}
        </button>
      ),
    },
    {
      key: "tipo",
      header: "Tipo",
      cell: (row) => (
        <span className="font-medium text-foreground truncate max-w-[140px] block">
          {row.tipo || "—"}
        </span>
      ),
    },
    {
      key: "brand_model",
      header: "Equipamento",
      cell: (row) => (
        <span className="text-foreground truncate max-w-[200px] block" title={`${row.brand} ${row.model}`}>
          {row.brand || ""} {row.model || "—"}
        </span>
      ),
    },
    {
      key: "status",
      header: "Status",
      cell: (row) => <MonochromaticStatusBadge status={row.status} />,
    },
    {
      key: "assigned_to",
      header: "Alocado Para",
      cell: (row) => (
        <span className="text-muted-foreground truncate max-w-[160px] block">
          {row.assigned_to || "—"}
        </span>
      ),
      hideBelow: "md",
    },
    {
      key: "revenda",
      header: "Unidade",
      cell: (row) => (
        <span className="text-secondary text-xs truncate max-w-[120px] block">
          {row.revenda || "—"}
        </span>
      ),
      hideBelow: "lg",
    },
    {
      key: "date_registered",
      header: "Cadastro",
      cell: (row) => (
        <span className="num font-mono text-xs text-muted-foreground">
          {formatDate(row.date_registered)}
        </span>
      ),
      hideBelow: "xl",
    },
    {
      key: "actions",
      header: "",
      align: "right",
      cell: (row) => (
        <div className="flex items-center justify-end gap-1">
          <Button
            variant="ghost"
            size="sm"
            className="size-8 p-0 text-muted-foreground hover:text-foreground"
            onClick={(e) => {
              e.stopPropagation();
              setSelectedAsset(row);
            }}
            title="Ver ficha técnica"
            aria-label={`Ver ficha de ${row.tipo} #${row.id}`}
          >
            <Eye className="size-3.5" />
          </Button>

          {hasRole("Gestor", "Técnico") && (
            <Button
              variant="ghost"
              size="sm"
              className="size-8 p-0 text-muted-foreground hover:text-foreground"
              onClick={(e) => {
                e.stopPropagation();
                setEditingItem(row);
              }}
              title="Edição rápida"
              aria-label={`Editar ${row.tipo} #${row.id}`}
            >
              <Pencil className="size-3.5" />
            </Button>
          )}
        </div>
      ),
    },
  ];

  return (
    <div className="space-y-6">
      {/* Topo Operacional (Seção 24: Contexto + Ações) */}
      <div className="border-b border-border pb-5">
        <div className="flex flex-col sm:flex-row sm:items-baseline sm:justify-between gap-3">
          <div>
            <span className="text-caption text-muted-foreground font-mono">
              INVENTÁRIO · DATA GRID OPERACIONAL
            </span>
            <div className="flex items-baseline gap-3 mt-0.5">
              <h1 className="text-heading-lg font-semibold tracking-tight text-foreground">
                Estoque de Equipamentos
              </h1>
              <span className="text-caption font-mono num text-muted-foreground">
                ({total} ativos registrados)
              </span>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => refetch()}
              disabled={isFetching}
              className="h-8 text-xs"
            >
              <RefreshCw className={cn("mr-1.5 size-3.5", isFetching && "animate-spin")} />
              {isFetching ? "Atualizando…" : "Atualizar"}
            </Button>

            {hasRole("Gestor", "Técnico") && (
              <Button
                size="sm"
                onClick={() => navigate({ to: "/register" })}
                className="h-8 text-xs"
              >
                <Plus className="mr-1.5 size-3.5" />
                Novo Ativo
              </Button>
            )}
          </div>
        </div>
      </div>

      {/* Toolbar Operacional de Filtros e Busca (Sem Cards Decorativos) */}
      <div className="bg-surface border border-border rounded-md p-4 space-y-3">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
          {/* Campo de Busca Direta */}
          <div className="relative flex-1 max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
            <Input
              value={search}
              onChange={(e) => {
                setSearch(e.target.value);
                setPage(0);
              }}
              placeholder="Buscar por patrimônio, marca, modelo, usuário ou serial..."
              className="h-9 pl-9 text-xs"
            />
            {search && (
              <button
                type="button"
                onClick={() => setSearch("")}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
              >
                <X className="size-3.5" />
              </button>
            )}
          </div>

          {/* Filtros em Linha */}
          <div className="flex flex-wrap items-center gap-2">
            <Select
              value={filterTipo}
              onValueChange={(val) => {
                setFilterTipo(val);
                setPage(0);
              }}
            >
              <SelectTrigger className="h-9 w-36 text-xs">
                <SelectValue placeholder="Tipo" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todos os tipos</SelectItem>
                {equipmentTypes.map((t) => (
                  <SelectItem key={t} value={t}>
                    {t}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>

            <Select
              value={filterStatus}
              onValueChange={(val) => {
                setFilterStatus(val);
                setPage(0);
              }}
            >
              <SelectTrigger className="h-9 w-36 text-xs">
                <SelectValue placeholder="Status" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todos status</SelectItem>
                {STATUS_OPTIONS.map((s) => (
                  <SelectItem key={s} value={s}>
                    {s}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>

            <Select
              value={filterRevenda}
              onValueChange={(val) => {
                setFilterRevenda(val);
                setPage(0);
              }}
            >
              <SelectTrigger className="h-9 w-40 text-xs">
                <SelectValue placeholder="Unidade" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todas as unidades</SelectItem>
                {unidades.map((u) => (
                  <SelectItem key={u.id} value={u.nome}>
                    {u.nome}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>

            {hasFilters && (
              <Button
                variant="ghost"
                size="sm"
                onClick={resetFilters}
                className="h-9 px-2 text-xs text-muted-foreground hover:text-foreground"
              >
                <X className="mr-1 size-3" /> Limpar
              </Button>
            )}
          </div>
        </div>
      </div>

      {/* Tabela Protagonista de Alta Densidade (7 itens por página) */}
      <div className="bg-surface border border-border rounded-md overflow-hidden">
        <DataTable
          data={items}
          columns={columns}
          rowKey={(row) => row.id}
          isLoading={isLoading}
          error={error}
          onRetry={() => void refetch()}
          onRowClick={(row) => setSelectedAsset(row)}
          emptyTitle="Nenhum ativo localizado"
          emptyDescription={
            hasFilters
              ? "Nenhum patrimônio corresponde aos filtros aplicados. Tente ajustar os parâmetros."
              : "Não há equipamentos cadastrados no estoque de TI."
          }
          pagination={{
            page,
            pageSize: PAGE_SIZE,
            total,
            onPageChange: setPage,
          }}
        />
      </div>

      {/* Ficha Técnica Lateral do Ativo Selecionado (AssetDetailsPanel) */}
      <AssetDetailsPanel
        item={selectedAsset}
        open={selectedAsset !== null}
        onOpenChange={(open) => !open && setSelectedAsset(null)}
      />

      {/* Modal de Edição Rápida */}
      <EditItemModal
        item={editingItem}
        open={editingItem !== null}
        onClose={() => setEditingItem(null)}
      />
    </div>
  );
}
