import { useEffect, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { Link, useNavigate } from "@tanstack/react-router";
import { Eye, Pencil, Plus, RefreshCw, X } from "lucide-react";
import { listItemsPaginated, type Item } from "@/api/items";
import { listUnidades } from "@/api/unidades";
import { DataTable, type Column } from "@/components/app/DataTable";
import { PageHeader } from "@/components/app/PageHeader";
import { StatusBadge } from "@/components/app/StatusBadge";
import { KpiCard } from "@/components/app/KpiCard";
import { ItemDetailsModal } from "@/components/app/ItemDetailsModal";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Button } from "@/components/ui/button";
import { useAuth } from "@/lib/auth";
import { useConstants } from "@/hooks/useConstants";
import { formatDate } from "@/lib/utils";

const STATUS_OPTIONS = ["Disponível", "Indisponível", "Pendente", "Pendente Devolução"];
const PAGE_SIZE = 20;

export function StockPage() {
  const { hasRole } = useAuth();
  const navigate = useNavigate();
  const { equipmentTypes, isLoading: constantsLoading } = useConstants();
  const [filterTipo, setFilterTipo] = useState("all");
  const [filterStatus, setFilterStatus] = useState("all");
  const [filterRevenda, setFilterRevenda] = useState("all");
  const [page, setPage] = useState(0);
  const [selectedItem, setSelectedItem] = useState<Item | null>(null);
  const { data: unidades = [] } = useQuery({ queryKey: ["unidades-filter"], queryFn: listUnidades });
  const { data, isLoading, isFetching, error, refetch } = useQuery({ queryKey: ["items", filterTipo, filterStatus, filterRevenda, page], queryFn: () => listItemsPaginated({ tipo: filterTipo !== "all" ? filterTipo : undefined, status: filterStatus !== "all" ? filterStatus : undefined, revenda: filterRevenda !== "all" ? filterRevenda : undefined, limit: PAGE_SIZE, offset: page * PAGE_SIZE }) });
  const items = data?.items ?? []; const total = data?.total ?? 0;
  const disponiveisCount = items.filter((i) => i.status === "Disponível").length;
  const indisponiveisCount = items.filter((i) => i.status === "Indisponível").length;
  const pendentesCount = items.filter((i) => i.status?.startsWith("Pendente")).length;
  const hasFilters = filterTipo !== "all" || filterStatus !== "all" || filterRevenda !== "all";
  useEffect(() => { if (data && page > 0 && items.length === 0 && total > 0) setPage(Math.max(0, Math.ceil(total / PAGE_SIZE) - 1)); }, [data, items.length, page, total]);
  const resetFilters = () => { setFilterTipo("all"); setFilterStatus("all"); setFilterRevenda("all"); setPage(0); };
  const columns: Column<Item>[] = [
    { key: "id", header: "ID", primary: true, cell: (row) => <button type="button" onClick={() => setSelectedItem(row)} className="font-mono text-xs font-semibold text-primary hover:underline">#{row.id}</button> },
    { key: "tipo", header: "Tipo", cell: (row) => <span className="block max-w-[180px] truncate" title={row.tipo || undefined}>{row.tipo || "-"}</span> },
    { key: "brand", header: "Marca", cell: (row) => row.brand || "-", hideBelow: "md" },
    { key: "model", header: "Modelo", cell: (row) => <span className="block max-w-[180px] truncate" title={row.model || undefined}>{row.model || "-"}</span>, hideBelow: "md" },
    { key: "status", header: "Status", cell: (row) => <StatusBadge status={row.status} /> },
    { key: "peripheral_count", header: "Periféricos", cell: (row) => row.peripheral_count ?? 0, hideBelow: "lg", align: "right" },
    { key: "assigned_to", header: "Usuário Alocado", cell: (row) => <span className="block max-w-[160px] truncate" title={row.assigned_to || undefined}>{row.assigned_to || "-"}</span>, hideBelow: "lg" },
    { key: "revenda", header: "Unidade", cell: (row) => row.revenda || "-", hideBelow: "md" },
    { key: "date_registered", header: "Data Cadastro", cell: (row) => formatDate(row.date_registered), hideBelow: "xl" },
    { key: "actions", header: "Ações", hideOnMobile: true, align: "right", cell: (row) => <div className="flex items-center justify-end gap-1"><Button variant="ghost" size="icon" className="size-8" onClick={(e) => { e.stopPropagation(); setSelectedItem(row); }} title="Ver detalhes" aria-label={`Ver detalhes de ${row.tipo || "equipamento"} #${row.id}`}><Eye className="size-4" aria-hidden /></Button>{hasRole("Gestor", "Técnico") && <Button variant="ghost" size="icon" className="size-8" asChild onClick={(e) => e.stopPropagation()} title="Editar equipamento" aria-label={`Editar equipamento #${row.id}`}><Link to="/edit/$id" params={{ id: String(row.id) }}><Pencil className="size-4" aria-hidden /></Link></Button>}</div> },
  ];
  return <div className="space-y-6">
    <PageHeader eyebrow="Inventário" title="Estoque de Equipamentos" description="Visão operacional do inventário de TI, sincronizada com os dados reais do sistema." actions={<><Button variant="outline" size="sm" onClick={() => refetch()} disabled={isFetching}><RefreshCw className={`mr-2 size-4 ${isFetching ? "animate-spin" : ""}`} aria-hidden />{isFetching ? "Atualizando…" : "Atualizar"}</Button>{hasRole("Gestor", "Técnico") && <Button size="sm" onClick={() => navigate({ to: "/register" })}><Plus className="mr-2 size-4" aria-hidden />Novo Equipamento</Button>}</>} />
    <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4"><KpiCard label="Total em Estoque" value={total} hint="Resultados conforme filtros" /><KpiCard label="Disponíveis" value={disponiveisCount} hint="Nesta página" /><KpiCard label="Indisponíveis" value={indisponiveisCount} hint="Nesta página" /><KpiCard label="Ações Pendentes" value={pendentesCount} hint="Nesta página" /></div>
    <div className="surface-panel p-3"><div className="flex flex-col gap-3 lg:flex-row lg:items-center"><span className="px-1 text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">Filtros</span><div className="grid min-w-0 flex-1 grid-cols-1 gap-2 sm:grid-cols-3"><Select value={filterTipo} onValueChange={(v) => { setFilterTipo(v); setPage(0); }} disabled={constantsLoading}><SelectTrigger className="w-full"><SelectValue placeholder="Tipo de Equipamento" /></SelectTrigger><SelectContent><SelectItem value="all">Todos os tipos</SelectItem>{equipmentTypes.map((t) => <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent></Select><Select value={filterStatus} onValueChange={(v) => { setFilterStatus(v); setPage(0); }}><SelectTrigger className="w-full"><SelectValue placeholder="Status do Item" /></SelectTrigger><SelectContent><SelectItem value="all">Todos os status</SelectItem>{STATUS_OPTIONS.map((s) => <SelectItem key={s} value={s}>{s}</SelectItem>)}</SelectContent></Select><Select value={filterRevenda} onValueChange={(v) => { setFilterRevenda(v); setPage(0); }}><SelectTrigger className="w-full"><SelectValue placeholder="Unidade" /></SelectTrigger><SelectContent><SelectItem value="all">Todas as unidades</SelectItem>{unidades.map((u) => <SelectItem key={u.id} value={u.nome}>{u.nome}</SelectItem>)}</SelectContent></Select></div>{hasFilters && <Button variant="ghost" size="sm" className="shrink-0" onClick={resetFilters}><X className="mr-1.5 size-3.5" />Limpar</Button>}</div></div>
    <DataTable data={items} columns={columns} rowKey={(row) => row.id} isLoading={isLoading} error={error} onRetry={() => refetch()} onRowClick={setSelectedItem} emptyTitle="Nenhum equipamento encontrado" emptyDescription={hasFilters ? "Nenhum equipamento corresponde aos filtros selecionados." : "Cadastre um equipamento para começar o inventário."} pagination={{ page, pageSize: PAGE_SIZE, total, onPageChange: setPage }} />
    <ItemDetailsModal item={selectedItem} onClose={() => setSelectedItem(null)} />
  </div>;
}
