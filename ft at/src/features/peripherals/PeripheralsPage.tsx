import { useState, type FormEvent } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { Download, Link2, Trash2, Unlink, RefreshCw } from "lucide-react";

import { createPeripheral, deletePeripheral, listPeripherals, listItemPeripherals, linkPeripheral, unlinkPeripheral, replacePeripheral, type Peripheral } from "@/api/peripherals";
import { listItemsPaginated } from "@/api/items";
import { useAuth } from "@/lib/auth";
import { useConstants } from "@/hooks/useConstants";
import { getErrorMessage } from "@/lib/api-error";
import { exportToCsv, formatDate } from "@/lib/utils";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { DataTable, type Column } from "@/components/app/DataTable";
import { KpiCard } from "@/components/app/KpiCard";
import { SearchableSelect } from "@/components/app/SearchableSelect";
import { FileUpload } from "@/components/app/FileUpload";
import { EmptyState } from "@/components/app/StateBlocks";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Badge } from "@/components/ui/badge";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger } from "@/components/ui/alert-dialog";

const LINK_ALLOWED_TYPES = ["Desktop", "Notebook", "Switch", "Impressora"];
const FETCH_ALL_LIMIT = 500;

function PeripheralStatusBadge({ status }: { status?: string }) {
  if (status === "Disponível") return <Badge className="border-emerald-500/25 bg-emerald-500/12 text-emerald-700 dark:text-emerald-400">{status}</Badge>;
  if (status === "Em Uso") return <Badge className="border-amber-500/30 bg-amber-500/14 text-amber-700 dark:text-amber-400">{status}</Badge>;
  if (status === "Substituido") return <Badge variant="destructive">{status}</Badge>;
  return <Badge variant="outline">{status ?? "-"}</Badge>;
}

export function PeripheralsPage() {
  const queryClient = useQueryClient();
  const { hasRole } = useAuth();
  const { peripheralTypes, isLoading: constantsLoading } = useConstants();
  const [tipo, setTipo] = useState(""); const [brand, setBrand] = useState(""); const [model, setModel] = useState(""); const [identificador, setIdentificador] = useState("");
  const [selectedItemId, setSelectedItemId] = useState(""); const [replacingOldId, setReplacingOldId] = useState<number | null>(null); const [replaceNewId, setReplaceNewId] = useState(""); const [replaceReason, setReplaceReason] = useState(""); const [replaceAttachment, setReplaceAttachment] = useState<File | null>(null);
  const { data: peripherals = [], isLoading, error, refetch } = useQuery({ queryKey: ["peripherals"], queryFn: () => listPeripherals() });
  const { data: itemsData } = useQuery({ queryKey: ["items"], queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }) });
  const items = itemsData?.items ?? []; const linkableItems = items.filter((i) => LINK_ALLOWED_TYPES.includes(i.tipo ?? ""));
  const { data: linkedPeripherals = [], refetch: refetchLinked } = useQuery({ queryKey: ["item-peripherals", selectedItemId], queryFn: () => listItemPeripherals(Number(selectedItemId)), enabled: !!selectedItemId });
  const availablePeripherals = peripherals.filter((p) => p.status === "Disponível" && !p.link_id);
  const selectedItem = items.find((i) => String(i.id) === selectedItemId);
  const canManage = hasRole("Gestor", "Técnico");

  const createMutation = useMutation({ mutationFn: createPeripheral, onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["peripherals"] }); setTipo(""); setBrand(""); setModel(""); setIdentificador(""); toast.success("Periférico cadastrado com sucesso!"); }, onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao cadastrar.")) });
  const linkMutation = useMutation({ mutationFn: (pid: number) => linkPeripheral(Number(selectedItemId), pid), onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["peripherals"] }); queryClient.invalidateQueries({ queryKey: ["item-peripherals", selectedItemId] }); toast.success("Periférico vinculado!"); }, onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao vincular.")) });
  const unlinkMutation = useMutation({ mutationFn: (id: number) => unlinkPeripheral(id), onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["peripherals"] }); queryClient.invalidateQueries({ queryKey: ["item-peripherals", selectedItemId] }); toast.success("Periférico desvinculado."); }, onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao desvincular.")) });
  const replaceMutation = useMutation({ mutationFn: () => replacePeripheral(Number(selectedItemId), replacingOldId!, Number(replaceNewId), replaceReason, replaceAttachment ?? undefined), onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["peripherals"] }); queryClient.invalidateQueries({ queryKey: ["item-peripherals", selectedItemId] }); setReplacingOldId(null); setReplaceNewId(""); setReplaceReason(""); setReplaceAttachment(null); toast.success("Periférico substituído com sucesso!"); }, onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao substituir.")) });
  const deleteMutation = useMutation({ mutationFn: (id: number) => deletePeripheral(id), onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["peripherals"] }); toast.success("Periférico inativado com sucesso."); }, onError: (e: unknown) => toast.error(getErrorMessage(e, "Erro ao inativar periférico.")) });

  function handleSubmit(e: FormEvent) { e.preventDefault(); createMutation.mutate({ tipo, brand, model, identificador }); }
  const columns: Column<Peripheral>[] = [
    { key: "id", header: "ID", cell: (p) => `#${p.id}`, primary: true }, { key: "tipo", header: "Tipo", cell: (p) => p.tipo }, { key: "brand", header: "Marca", cell: (p) => p.brand || "-", hideBelow: "md" }, { key: "model", header: "Modelo", cell: (p) => p.model || "-", hideBelow: "lg" }, { key: "identificador", header: "Identificador (S/N)", cell: (p) => p.identificador || "-" }, { key: "status", header: "Status", cell: (p) => <PeripheralStatusBadge status={p.status} /> }, { key: "date_registered", header: "Cadastro", cell: (p) => formatDate(p.date_registered), hideBelow: "lg" },
    ...(canManage ? [{ key: "actions", header: "", hideOnMobile: true, cell: (p: Peripheral) => <AlertDialog><AlertDialogTrigger asChild><Button size="sm" variant="ghost" className="text-destructive hover:text-destructive"><Trash2 className="size-3.5" /></Button></AlertDialogTrigger><AlertDialogContent><AlertDialogHeader><AlertDialogTitle>Inativar periférico #{p.id}?</AlertDialogTitle><AlertDialogDescription>Esta ação inativa o cadastro. Periféricos em uso não podem ser inativados.</AlertDialogDescription></AlertDialogHeader><AlertDialogFooter><AlertDialogCancel>Cancelar</AlertDialogCancel><AlertDialogAction onClick={() => deleteMutation.mutate(p.id)}>Inativar</AlertDialogAction></AlertDialogFooter></AlertDialogContent></AlertDialog> } satisfies Column<Peripheral>] : [])
  ];

  return <div className="space-y-8">
    <PageHeader title="Periféricos" description="Cadastre, consulte e vincule periféricos aos equipamentos em um único fluxo." actions={<Button variant="outline" size="sm" disabled={!peripherals.length} onClick={() => exportToCsv("perifericos", peripherals as unknown as Record<string, unknown>[], [{ key: "id", label: "ID" }, { key: "tipo", label: "Tipo" }, { key: "brand", label: "Marca" }, { key: "model", label: "Modelo" }, { key: "identificador", label: "Identificador" }, { key: "status", label: "Status" }, { key: "date_registered", label: "Cadastro" }])}><Download className="mr-2 size-3.5" />Exportar CSV</Button>} />
    <div className="grid grid-cols-1 gap-3 sm:grid-cols-3"><KpiCard label="Total de periféricos" value={peripherals.length} /><KpiCard label="Vinculados" value={peripherals.filter((p) => p.link_id).length} /><KpiCard label="Disponíveis" value={availablePeripherals.length} /></div>
    <Section title="Cadastrar periférico"><form onSubmit={handleSubmit} className="space-y-4"><div className="grid grid-cols-1 gap-4 sm:grid-cols-2"><div className="flex flex-col gap-1.5"><Label>Tipo *</Label><Select value={tipo} onValueChange={setTipo} disabled={constantsLoading}><SelectTrigger><SelectValue placeholder="Selecione o tipo" /></SelectTrigger><SelectContent>{peripheralTypes.map((t) => <SelectItem key={t} value={t}>{t}</SelectItem>)}</SelectContent></Select></div><div className="flex flex-col gap-1.5"><Label>Identificador (S/N) *</Label><Input value={identificador} onChange={(e) => setIdentificador(e.target.value)} required placeholder="Número de série único" /></div><div className="flex flex-col gap-1.5"><Label>Marca</Label><Input value={brand} onChange={(e) => setBrand(e.target.value)} /></div><div className="flex flex-col gap-1.5"><Label>Modelo</Label><Input value={model} onChange={(e) => setModel(e.target.value)} /></div></div><div className="flex justify-end"><Button type="submit" disabled={createMutation.isPending}>{createMutation.isPending ? "Cadastrando…" : "Cadastrar periférico"}</Button></div></form></Section>

    {canManage && <Section title="Vincular periférico" description="O vínculo agora acontece dentro da própria área de Periféricos. Não é necessário abrir outra tela."><div className="space-y-5"><div className="max-w-xl"><Label>Equipamento</Label><div className="mt-1.5"><SearchableSelect options={linkableItems.map((i) => ({ value: String(i.id), label: `#${i.id} — ${i.tipo ?? ""} ${i.brand ?? ""} ${i.model ?? ""}`, subtitle: [i.revenda, i.identificador].filter(Boolean).join(" • ") }))} value={selectedItemId} onValueChange={setSelectedItemId} placeholder="Selecione ou busque um equipamento…" searchPlaceholder="Buscar por ID, tipo, marca, modelo, patrimônio…" /></div></div>{selectedItemId && <div className="grid gap-5 lg:grid-cols-2"><div><div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-semibold">Vinculados ({linkedPeripherals.length})</h3><Button size="icon" variant="ghost" onClick={() => void refetchLinked()}><RefreshCw className="size-3.5" /></Button></div>{linkedPeripherals.length === 0 ? <EmptyState title="Nenhum periférico vinculado" /> : <div className="space-y-2">{linkedPeripherals.map((p) => <div key={p.link_id} className="rounded-lg border border-border p-3"><div className="flex items-center justify-between gap-3"><div className="min-w-0"><p className="truncate text-sm font-medium">{p.tipo} — {p.brand || "-"} {p.model || ""}</p><p className="text-xs text-muted-foreground">S/N: {p.identificador || "-"}</p></div><Button size="sm" variant="outline" onClick={() => unlinkMutation.mutate(p.link_id!)}><Unlink className="mr-1.5 size-3.5" />Desvincular</Button></div><button className="mt-2 text-xs text-primary hover:underline" onClick={() => { setReplacingOldId(p.id); setReplaceNewId(""); setReplaceReason(""); }}>Substituir…</button></div>)}</div>}</div><div><h3 className="mb-3 text-sm font-semibold">Disponíveis ({availablePeripherals.length})</h3>{availablePeripherals.length === 0 ? <EmptyState title="Nenhum periférico disponível" /> : <div className="max-h-[420px] space-y-2 overflow-y-auto">{availablePeripherals.map((p) => <div key={p.id} className="flex items-center justify-between gap-3 rounded-lg border border-border p-3"><div className="min-w-0"><p className="truncate text-sm font-medium">{p.tipo} — {p.brand || "-"} {p.model || ""}</p><p className="text-xs text-muted-foreground">S/N: {p.identificador || "-"}</p></div><Button size="sm" variant="outline" onClick={() => linkMutation.mutate(p.id)}><Link2 className="mr-1.5 size-3.5" />Vincular</Button></div>)}</div>}</div></div>}</div></Section>}

    {replacingOldId && <Section title={`Substituir periférico #${replacingOldId}`}><div className="grid gap-4 sm:grid-cols-2"><div><Label>Novo periférico</Label><Select value={replaceNewId} onValueChange={setReplaceNewId}><SelectTrigger><SelectValue placeholder="Selecione o substituto" /></SelectTrigger><SelectContent>{availablePeripherals.map((p) => <SelectItem key={p.id} value={String(p.id)}>#{p.id} — {p.tipo} {p.brand} {p.model}</SelectItem>)}</SelectContent></Select></div><div><Label>Motivo *</Label><Input value={replaceReason} onChange={(e) => setReplaceReason(e.target.value)} placeholder="Ex: Defeito, upgrade…" /></div></div><div className="mt-4"><FileUpload accept="application/pdf,image/*" onFile={setReplaceAttachment} label="Comprovante (opcional)" /></div><div className="mt-4 flex gap-3"><Button disabled={!replaceNewId || !replaceReason || replaceMutation.isPending} onClick={() => replaceMutation.mutate()}>{replaceMutation.isPending ? "Substituindo…" : "Confirmar substituição"}</Button><Button variant="ghost" onClick={() => setReplacingOldId(null)}>Cancelar</Button></div></Section>}
    <DataTable data={peripherals} columns={columns} rowKey={(p) => p.id} isLoading={isLoading} error={error} onRetry={() => void refetch()} emptyTitle="Nenhum periférico cadastrado" />
  </div>;
}
