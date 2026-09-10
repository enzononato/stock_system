import { useMemo, useState, type FormEvent } from "react";
import { useMutation, useQueries, useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { Download, Link2, Trash2, Unlink, RefreshCw, Plus, Search, ChevronRight } from "lucide-react";

import {
  createPeripheral,
  deletePeripheral,
  listPeripherals,
  listItemPeripherals,
  linkPeripheral,
  unlinkPeripheral,
  replacePeripheral,
  type Peripheral,
} from "@/api/peripherals";
import { listItemsPaginated } from "@/api/items";
import { listUnidades } from "@/api/unidades";
import { useAuth } from "@/lib/auth";
import { useConstants } from "@/hooks/useConstants";
import { getErrorMessage } from "@/lib/api-error";
import { exportToCsv, formatDate } from "@/lib/utils";
import { SearchableSelect } from "@/components/app/SearchableSelect";
import { FileUpload } from "@/components/app/FileUpload";
import { EmptyState } from "@/components/app/StateBlocks";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Badge } from "@/components/ui/badge";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from "@/components/ui/alert-dialog";

const LINK_ALLOWED_TYPES = ["Desktop", "Notebook", "Switch", "Impressora"];
const FETCH_ALL_LIMIT = 500;

function StatusBadge({ status }: { status?: string }) {
  const symbol = status === "Disponível" ? "○" : status === "Em Uso" ? "●" : "×";
  return (
    <span
      className={`font-mono text-caption font-semibold ${
        status === "Disponível"
          ? "text-foreground"
          : status === "Em Uso"
            ? "text-foreground"
            : "text-muted-foreground"
      }`}
    >
      {symbol} {status ?? "—"}
    </span>
  );
}

export function PeripheralsPage() {
  const queryClient = useQueryClient();
  const { hasRole } = useAuth();
  const { peripheralTypes, revendas, isLoading: constantsLoading } = useConstants();

  // Form: novo periférico
  const [tipo, setTipo] = useState("");
  const [brand, setBrand] = useState("");
  const [model, setModel] = useState("");
  const [identificador, setIdentificador] = useState("");
  const [showCreateForm, setShowCreateForm] = useState(false);

  // Filters
  const [filterTipo, setFilterTipo] = useState("all");
  const [filterStatus, setFilterStatus] = useState("all");
  const [searchQuery, setSearchQuery] = useState("");

  // Detail / relação
  const [selectedPeripheralId, setSelectedPeripheralId] = useState<number | null>(null);
  const [selectedItemId, setSelectedItemId] = useState("");

  // Substituição
  const [replacingOldId, setReplacingOldId] = useState<number | null>(null);
  const [replaceNewId, setReplaceNewId] = useState("");
  const [replaceReason, setReplaceReason] = useState("");
  const [replaceAttachment, setReplaceAttachment] = useState<File | null>(null);

  const {
    data: peripherals = [],
    isLoading,
    error,
    refetch,
  } = useQuery({ queryKey: ["peripherals"], queryFn: () => listPeripherals() });

  const { data: itemsData } = useQuery({
    queryKey: ["items"],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  });
  const items = itemsData?.items ?? [];
  const linkableItems = items.filter((i) => LINK_ALLOWED_TYPES.includes(i.tipo ?? ""));

  const { data: linkedPeripherals = [], refetch: refetchLinked } = useQuery({
    queryKey: ["item-peripherals", selectedItemId],
    queryFn: () => listItemPeripherals(Number(selectedItemId)),
    enabled: !!selectedItemId,
  });

  const { data: unidades = [] } = useQuery({
    queryKey: ["unidades"],
    queryFn: () => listUnidades(),
  });

  const availablePeripherals = peripherals.filter((p) => p.status === "Disponível" && !p.link_id);
  const canManage = hasRole("Gestor", "Técnico");

  const itemsWithPeripherals = useMemo(
    () => items.filter((i) => (i.peripheral_count ?? 0) > 0),
    [items],
  );
  const peripheralQueries = useQueries({
    queries: itemsWithPeripherals.map((item) => ({
      queryKey: ["item-peripherals-map", item.id],
      queryFn: () => listItemPeripherals(item.id),
      staleTime: 5 * 60 * 1000,
    })),
  });
  const peripheralRevendaMap = useMemo(() => {
    const map = new Map<number, string>();
    for (let i = 0; i < itemsWithPeripherals.length; i++) {
      const item = itemsWithPeripherals[i];
      const result = peripheralQueries[i];
      if (item && result?.data && item.revenda) {
        for (const p of result.data) {
          map.set(p.id, item.revenda);
        }
      }
    }
    return map;
  }, [itemsWithPeripherals, peripheralQueries]);

  const filteredPeripherals = useMemo(() => {
    return peripherals.filter((p) => {
      if (filterTipo !== "all" && p.tipo !== filterTipo) return false;
      if (filterStatus !== "all" && p.status !== filterStatus) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const match =
          String(p.id).includes(q) ||
          (p.tipo ?? "").toLowerCase().includes(q) ||
          (p.brand ?? "").toLowerCase().includes(q) ||
          (p.model ?? "").toLowerCase().includes(q) ||
          (p.identificador ?? "").toLowerCase().includes(q);
        if (!match) return false;
      }
      return true;
    });
  }, [peripherals, filterTipo, filterStatus, searchQuery, peripheralRevendaMap]);

  const [tablePage, setTablePage] = useState(0);
  const TABLE_PAGE_SIZE = 7;
  const totalTablePages = Math.max(1, Math.ceil(filteredPeripherals.length / TABLE_PAGE_SIZE));
  const paginatedPeripherals = filteredPeripherals.slice(
    tablePage * TABLE_PAGE_SIZE,
    (tablePage + 1) * TABLE_PAGE_SIZE
  );

  const selectedPeripheral = peripherals.find((p) => p.id === selectedPeripheralId);
  // find which item is linked to this peripheral
  const linkedToItem = selectedPeripheral?.link_id
    ? items.find((i) =>
        linkedPeripherals.some((lp) => lp.id === selectedPeripheral.id && String(i.id) === selectedItemId),
      ) || items.find((i) => {
        // fallback: search via peripheral queries
        for (let idx = 0; idx < itemsWithPeripherals.length; idx++) {
          const item = itemsWithPeripherals[idx];
          const result = peripheralQueries[idx];
          if (result?.data?.some((lp) => lp.id === selectedPeripheral.id)) return item?.id === i.id;
        }
        return false;
      })
    : null;

  const createMutation = useMutation({
    mutationFn: createPeripheral,
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["peripherals"] });
      setTipo(""); setBrand(""); setModel(""); setIdentificador("");
      setShowCreateForm(false);
      toast.success("Periférico cadastrado com sucesso!");
    },
    onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao cadastrar periférico.")),
  });

  const linkMutation = useMutation({
    mutationFn: (pid: number) => linkPeripheral(Number(selectedItemId), pid),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["peripherals"] });
      void queryClient.invalidateQueries({ queryKey: ["item-peripherals", selectedItemId] });
      toast.success("Periférico vinculado!");
    },
    onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao vincular.")),
  });

  const unlinkMutation = useMutation({
    mutationFn: (id: number) => unlinkPeripheral(id),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["peripherals"] });
      void queryClient.invalidateQueries({ queryKey: ["item-peripherals", selectedItemId] });
      toast.success("Periférico desvinculado.");
    },
    onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao desvincular.")),
  });

  const replaceMutation = useMutation({
    mutationFn: () =>
      replacePeripheral(
        Number(selectedItemId),
        replacingOldId!,
        Number(replaceNewId),
        replaceReason,
        replaceAttachment ?? undefined,
      ),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["peripherals"] });
      void queryClient.invalidateQueries({ queryKey: ["item-peripherals", selectedItemId] });
      setReplacingOldId(null); setReplaceNewId(""); setReplaceReason(""); setReplaceAttachment(null);
      toast.success("Periférico substituído com sucesso!");
    },
    onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao substituir.")),
  });

  const deleteMutation = useMutation({
    mutationFn: (id: number) => deletePeripheral(id),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["peripherals"] });
      setSelectedPeripheralId(null);
      toast.success("Periférico inativado.");
    },
    onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao inativar.")),
  });

  function handleSubmit(e: FormEvent) {
    e.preventDefault();
    createMutation.mutate({ tipo, brand, model, identificador });
  }

  return (
    <div className="page-container-dense space-y-6">
      {/* Header */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Inventário Relacional
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Master / Detail</span>
          </div>
          <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">
            Periféricos e Vínculos
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Gestão de periféricos — mouse, teclado, monitor, headset — com rastreabilidade relacional ao ativo principal.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <Button
            variant="outline"
            size="sm"
            disabled={!peripherals.length}
            onClick={() =>
              exportToCsv("perifericos", filteredPeripherals as unknown as Record<string, unknown>[], [
                { key: "id", label: "ID" },
                { key: "tipo", label: "Tipo" },
                { key: "brand", label: "Marca" },
                { key: "model", label: "Modelo" },
                { key: "identificador", label: "S/N" },
                { key: "status", label: "Status" },
              ])
            }
            className="rounded-[4px] border-border text-foreground hover:bg-muted text-xs h-8"
          >
            <Download className="mr-1.5 size-3.5" />
            Exportar CSV
          </Button>
          {canManage && (
            <Button
              size="sm"
              onClick={() => setShowCreateForm((p) => !p)}
              className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs h-8"
            >
              <Plus className="mr-1.5 size-3.5" />
              Novo Periférico
            </Button>
          )}
        </div>
      </div>

      {/* Painel de Cadastro Rápido (colapsável) */}
      {showCreateForm && canManage && (
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-4">
          <div className="border-b border-border pb-2.5">
            <h2 className="text-body font-semibold text-foreground">Cadastrar Novo Periférico</h2>
            <p className="text-caption text-muted-foreground">
              Registre um periférico individual para posterior vinculação a um ativo.
            </p>
          </div>
          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="grid grid-cols-1 sm:grid-cols-4 gap-4">
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Tipo *</Label>
                <Select value={tipo} onValueChange={setTipo} disabled={constantsLoading}>
                  <SelectTrigger className="rounded-[4px] border-border bg-background text-body-sm h-9">
                    <SelectValue placeholder="Selecione o tipo" />
                  </SelectTrigger>
                  <SelectContent className="rounded-[4px] border-border bg-surface">
                    {peripheralTypes.map((t) => (
                      <SelectItem key={t} value={t} className="text-body-sm rounded-[2px]">{t}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">S/N (Identificador) *</Label>
                <Input
                  value={identificador}
                  onChange={(e) => setIdentificador(e.target.value)}
                  required
                  placeholder="Número de série único"
                  className="font-mono rounded-[4px] border-border bg-background text-body-sm h-9"
                />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Marca</Label>
                <Input value={brand} onChange={(e) => setBrand(e.target.value)} placeholder="Ex: Logitech" className="rounded-[4px] border-border bg-background text-body-sm h-9" />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Modelo</Label>
                <Input value={model} onChange={(e) => setModel(e.target.value)} placeholder="Ex: MX Keys" className="rounded-[4px] border-border bg-background text-body-sm h-9" />
              </div>
            </div>
            <div className="flex justify-end gap-2">
              <Button type="button" variant="outline" size="sm" onClick={() => setShowCreateForm(false)} className="rounded-[4px] border-border text-xs h-8">
                Cancelar
              </Button>
              <Button type="submit" size="sm" disabled={createMutation.isPending} className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs h-8">
                {createMutation.isPending ? "Cadastrando…" : "Cadastrar Periférico"}
              </Button>
            </div>
          </form>
        </div>
      )}

      {/* Grid Master / Detail */}
      <div className="grid grid-cols-1 lg:grid-cols-[minmax(0,1fr)_360px] gap-6 items-start">
        {/* MASTER — Lista de Periféricos */}
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-4">
          {/* Toolbar de Filtros */}
          <div className="flex flex-col sm:flex-row gap-3 border-b border-border pb-4">
            <div className="relative flex-1">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
              <Input
                value={searchQuery}
                onChange={(e) => {
                  setSearchQuery(e.target.value);
                  setTablePage(0);
                }}
                placeholder="Buscar por ID, tipo, marca, S/N…"
                className="pl-8 h-8 rounded-[4px] border-border bg-background text-body-sm"
              />
            </div>
            <Select
              value={filterTipo}
              onValueChange={(val) => {
                setFilterTipo(val);
                setTablePage(0);
              }}
            >
              <SelectTrigger className="w-36 rounded-[4px] border-border bg-background text-body-sm h-8">
                <SelectValue placeholder="Tipo" />
              </SelectTrigger>
              <SelectContent className="rounded-[4px] border-border bg-surface">
                <SelectItem value="all" className="text-body-sm rounded-[2px]">Todos os tipos</SelectItem>
                {peripheralTypes.map((t) => (
                  <SelectItem key={t} value={t} className="text-body-sm rounded-[2px]">{t}</SelectItem>
                ))}
              </SelectContent>
            </Select>
            <Select
              value={filterStatus}
              onValueChange={(val) => {
                setFilterStatus(val);
                setTablePage(0);
              }}
            >
              <SelectTrigger className="w-36 rounded-[4px] border-border bg-background text-body-sm h-8">
                <SelectValue placeholder="Status" />
              </SelectTrigger>
              <SelectContent className="rounded-[4px] border-border bg-surface">
                <SelectItem value="all" className="text-body-sm rounded-[2px]">Todos</SelectItem>
                <SelectItem value="Disponível" className="text-body-sm rounded-[2px]">Disponível</SelectItem>
                <SelectItem value="Em Uso" className="text-body-sm rounded-[2px]">Em Uso</SelectItem>
                <SelectItem value="Substituido" className="text-body-sm rounded-[2px]">Substituído</SelectItem>
              </SelectContent>
            </Select>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-body-sm">
              <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground">
                <tr>
                  <th className="py-2 px-3">ID</th>
                  <th className="py-2 px-3">Tipo</th>
                  <th className="py-2 px-3">Marca / Modelo</th>
                  <th className="py-2 px-3">S/N</th>
                  <th className="py-2 px-3">Status</th>
                  <th className="py-2 px-3">Cadastro</th>
                  <th className="py-2 px-3 text-right">Detalhe</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {filteredPeripherals.length === 0 ? (
                  <tr>
                    <td colSpan={7} className="py-8 text-center text-caption text-muted-foreground">
                      Nenhum periférico encontrado.
                    </td>
                  </tr>
                ) : (
                  paginatedPeripherals.map((p) => {
                    const isSelected = selectedPeripheralId === p.id;
                    return (
                      <tr
                        key={p.id}
                        onClick={() => setSelectedPeripheralId(isSelected ? null : p.id)}
                        className={`cursor-pointer transition-colors ${isSelected ? "bg-surface-alt" : "hover:bg-muted/30"}`}
                      >
                        <td className="py-2 px-3 font-mono font-semibold">#{p.id}</td>
                        <td className="py-2 px-3 font-medium">{p.tipo}</td>
                        <td className="py-2 px-3 text-muted-foreground">{p.brand} {p.model}</td>
                        <td className="py-2 px-3 font-mono text-foreground">{p.identificador || "—"}</td>
                        <td className="py-2 px-3"><StatusBadge status={p.status} /></td>
                        <td className="py-2 px-3 font-mono text-muted-foreground">{formatDate(p.date_registered)}</td>
                        <td className="py-2 px-3 text-right">
                          <ChevronRight className={`size-4 ml-auto transition-transform ${isSelected ? "rotate-90 text-foreground" : "text-muted-foreground"}`} />
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          <div className="flex items-center justify-between border-t border-border pt-3">
            <span className="text-caption text-muted-foreground">
              {totalTablePages > 1
                ? `Página ${tablePage + 1} de ${totalTablePages} (${filteredPeripherals.length} periféricos)`
                : `${filteredPeripherals.length} periférico${filteredPeripherals.length !== 1 ? "s" : ""}`}
            </span>
            {totalTablePages > 1 && (
              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant="outline"
                  disabled={tablePage === 0}
                  onClick={() => setTablePage((p) => Math.max(0, p - 1))}
                  className="rounded-[4px] border-border text-xs h-7 px-2.5"
                >
                  ← Anterior
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  disabled={tablePage >= totalTablePages - 1}
                  onClick={() => setTablePage((p) => p + 1)}
                  className="rounded-[4px] border-border text-xs h-7 px-2.5"
                >
                  Próxima →
                </Button>
              </div>
            )}
          </div>
        </div>

        {/* DETAIL — Painel de Relações */}
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-4 min-h-[400px]">
          {selectedPeripheral ? (
            <>
              <div className="border-b border-border pb-3">
                <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                  Ficha do Periférico
                </span>
                <p className="text-body-lg font-semibold text-foreground mt-0.5">
                  #{selectedPeripheral.id} — {selectedPeripheral.tipo}
                </p>
                <p className="text-caption text-muted-foreground">
                  {selectedPeripheral.brand} {selectedPeripheral.model} • S/N:{" "}
                  <span className="font-mono text-foreground">{selectedPeripheral.identificador || "—"}</span>
                </p>
              </div>

              <div className="space-y-2 text-caption">
                <div className="flex justify-between py-1 border-b border-border/50">
                  <span className="text-muted-foreground">Status</span>
                  <StatusBadge status={selectedPeripheral.status} />
                </div>
                <div className="flex justify-between py-1 border-b border-border/50">
                  <span className="text-muted-foreground">Vinculado</span>
                  <span className="font-semibold text-foreground">
                    {selectedPeripheral.link_id ? `● Sim (link #${selectedPeripheral.link_id})` : "○ Livre"}
                  </span>
                </div>
                <div className="flex justify-between py-1 border-b border-border/50">
                  <span className="text-muted-foreground">Cadastro</span>
                  <span className="font-mono text-foreground">{formatDate(selectedPeripheral.date_registered)}</span>
                </div>
              </div>

              {/* Ações de Gestão */}
              {canManage && (
                <div className="space-y-3 pt-3 border-t border-border">
                  <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                    Ações de Gestão
                  </span>

                  {selectedPeripheral.link_id ? (
                    <Button
                      size="sm"
                      variant="outline"
                      disabled={unlinkMutation.isPending}
                      onClick={() => unlinkMutation.mutate(selectedPeripheral.link_id!)}
                      className="w-full rounded-[4px] border-border text-foreground hover:bg-muted text-xs h-8"
                    >
                      <Unlink className="mr-1.5 size-3.5" />
                      Desvincular do Ativo
                    </Button>
                  ) : (
                    <div className="space-y-2">
                      <Label className="text-caption font-medium">Vincular a um ativo:</Label>
                      <SearchableSelect
                        options={linkableItems.map((i) => ({
                          value: String(i.id),
                          label: `#${i.id} — ${i.tipo ?? ""} ${i.brand ?? ""} ${i.model ?? ""}`,
                          subtitle: i.revenda || undefined,
                        }))}
                        value={selectedItemId}
                        onValueChange={setSelectedItemId}
                        placeholder="Selecione o equipamento…"
                        searchPlaceholder="Buscar ativo…"
                      />
                      {selectedItemId && (
                        <Button
                          size="sm"
                          disabled={linkMutation.isPending}
                          onClick={() => linkMutation.mutate(selectedPeripheral.id)}
                          className="w-full rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs h-8"
                        >
                          <Link2 className="mr-1.5 size-3.5" />
                          Confirmar Vínculo
                        </Button>
                      )}
                    </div>
                  )}

                  <AlertDialog>
                    <AlertDialogTrigger asChild>
                      <Button
                        size="sm"
                        variant="outline"
                        className="w-full rounded-[4px] border-border text-muted-foreground hover:text-foreground text-xs h-8"
                      >
                        <Trash2 className="mr-1.5 size-3.5" />
                        Inativar Periférico
                      </Button>
                    </AlertDialogTrigger>
                    <AlertDialogContent className="rounded-[6px] border-border bg-surface">
                      <AlertDialogHeader>
                        <AlertDialogTitle className="text-body-lg font-semibold">
                          Inativar Periférico #{selectedPeripheral.id}?
                        </AlertDialogTitle>
                        <AlertDialogDescription className="text-caption text-muted-foreground">
                          Periféricos em uso ativo não podem ser inativados. Esta ação registra o equipamento como baixado no sistema.
                        </AlertDialogDescription>
                      </AlertDialogHeader>
                      <AlertDialogFooter>
                        <AlertDialogCancel className="rounded-[4px] border-border text-xs">Cancelar</AlertDialogCancel>
                        <AlertDialogAction
                          className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs"
                          onClick={() => deleteMutation.mutate(selectedPeripheral.id)}
                        >
                          Confirmar Inativação
                        </AlertDialogAction>
                      </AlertDialogFooter>
                    </AlertDialogContent>
                  </AlertDialog>
                </div>
              )}
            </>
          ) : (
            <div className="flex flex-col items-center justify-center h-full py-12 text-center text-caption text-muted-foreground space-y-2">
              <ChevronRight className="size-8 text-border" />
              <p className="font-medium text-body-sm text-foreground">Selecione um periférico</p>
              <p>Clique em qualquer linha da tabela para visualizar a ficha técnica e opções de vínculo relacional.</p>
            </div>
          )}
        </div>
      </div>

      {/* Seção: Vincular Periférico a Equipamento (por Equipamento) */}
      {canManage && (
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-4">
          <div className="border-b border-border pb-2.5">
            <h2 className="text-body font-semibold text-foreground">Visualização por Equipamento</h2>
            <p className="text-caption text-muted-foreground">
              Selecione um ativo para ver todos os periféricos vinculados e gerenciar cada relação.
            </p>
          </div>

          <div className="max-w-xl">
            <SearchableSelect
              options={linkableItems.map((i) => ({
                value: String(i.id),
                label: `#${i.id} — ${i.tipo ?? ""} ${i.brand ?? ""} ${i.model ?? ""}`,
                subtitle: [i.revenda, i.identificador].filter(Boolean).join(" • "),
              }))}
              value={selectedItemId}
              onValueChange={setSelectedItemId}
              placeholder="Selecione ou busque um equipamento…"
              searchPlaceholder="Buscar por ID, tipo, marca, modelo…"
            />
          </div>

          {selectedItemId && (
            <div className="space-y-3">
              {linkedPeripherals.length === 0 ? (
                <p className="text-caption text-muted-foreground italic py-4 text-center">
                  Nenhum periférico vinculado a este ativo.
                </p>
              ) : (
                <div className="space-y-2">
                  <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                    Periféricos Vinculados ({linkedPeripherals.length})
                  </span>
                  {linkedPeripherals.map((p) => (
                    <div
                      key={p.id}
                      className="flex items-center justify-between gap-3 rounded-[4px] border border-border bg-surface-alt px-3 py-2.5 text-body-sm"
                    >
                      <div>
                        <span className="font-medium text-foreground">{p.tipo} — {p.brand} {p.model}</span>
                        <span className="font-mono text-caption text-muted-foreground ml-2">S/N: {p.identificador || "—"}</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <StatusBadge status={p.status} />
                        <Button
                          size="sm"
                          variant="outline"
                          disabled={unlinkMutation.isPending}
                          onClick={() => p.link_id && unlinkMutation.mutate(p.link_id)}
                          className="rounded-[4px] border-border text-xs h-7 px-2"
                        >
                          <Unlink className="size-3 mr-1" />
                          Desvincular
                        </Button>
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {availablePeripherals.length > 0 && (
                <div className="space-y-2 pt-3 border-t border-border">
                  <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                    Adicionar Periférico Disponível
                  </span>
                  <div className="flex gap-2">
                    <SearchableSelect
                      options={availablePeripherals.map((p) => ({
                        value: String(p.id),
                        label: `#${p.id} — ${p.tipo} ${p.brand} ${p.model}`,
                        subtitle: `S/N: ${p.identificador || "—"}`,
                      }))}
                      value=""
                      onValueChange={(pid) => linkMutation.mutate(Number(pid))}
                      placeholder="Selecione um periférico disponível…"
                      searchPlaceholder="Buscar periférico livre…"
                    />
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
