import { useState } from "react";
import { useNavigate } from "@tanstack/react-router";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { ArrowLeft, Link2, RefreshCw, Unlink } from "lucide-react";

import { listItemsPaginated } from "@/api/items";
import {
  linkPeripheral,
  listItemPeripherals,
  listPeripherals,
  replacePeripheral,
  unlinkPeripheral,
  type Peripheral,
} from "@/api/peripherals";
import { getErrorMessage } from "@/lib/api-error";
import { PageHeader, Section } from "@/components/app/PageHeader";
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

// Tipos de equipamento que aceitam periféricos. Regra própria desta tela
// (subconjunto curado, não um cadastro de domínio), por isso continua fixa.
const LINK_ALLOWED_TYPES = ["Desktop", "Notebook", "Switch", "Impressora"];

// Sem paginação nesta tela (filtra os equipamentos linkáveis client-side a
// partir da lista completa) — usamos o teto de página do backend para não
// truncar em 50 itens (default de GET /api/items) como aconteceria chamando
// listItemsPaginated() sem limit.
const FETCH_ALL_LIMIT = 500;

function peripheralBadgeVariant(status?: string) {
  if (status === "Disponível") return "badge-success";
  if (status === "Em Uso") return "badge-warning";
  return "badge-muted";
}

function PeripheralCard({
  peripheral,
  action,
  actionLabel,
  actionIcon,
  variant = "outline",
}: {
  peripheral: Peripheral;
  action: () => void;
  actionLabel: string;
  actionIcon: React.ReactNode;
  variant?: "outline" | "destructive";
}) {
  return (
    <div className="flex items-center justify-between gap-3 rounded-lg border border-border bg-card px-4 py-3">
      <div className="flex min-w-0 flex-col gap-0.5">
        <span className="truncate text-sm font-medium text-foreground">
          {peripheral.tipo} — {peripheral.brand || "-"} {peripheral.model || ""}
        </span>
        <span className="text-xs text-muted-foreground">
          S/N: {peripheral.identificador || "-"}
        </span>
      </div>
      <div className="flex shrink-0 items-center gap-3">
        <Badge
          variant={peripheral.status === "Substituido" ? "destructive" : "outline"}
          className={peripheralBadgeVariant(peripheral.status)}
        >
          {peripheral.status}
        </Badge>
        <Button size="sm" variant={variant} onClick={action}>
          {actionIcon}
          <span className="ml-1">{actionLabel}</span>
        </Button>
      </div>
    </div>
  );
}

export function LinkPeripheralPage() {
  const queryClient = useQueryClient();
  const navigate = useNavigate();
  const [selectedItemId, setSelectedItemId] = useState("");

  // Substituição
  const [replacingLinkId, setReplacingLinkId] = useState<number | null>(null);
  const [replacingOldId, setReplacingOldId] = useState<number | null>(null);
  const [replaceNewId, setReplaceNewId] = useState("");
  const [replaceReason, setReplaceReason] = useState("");
  const [replaceAttachment, setReplaceAttachment] = useState<File | null>(null);

  const { data } = useQuery({
    queryKey: ["items"],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  });
  const items = data?.items ?? [];
  const linkableItems = items.filter((i) => LINK_ALLOWED_TYPES.includes(i.tipo ?? ""));

  const { data: linkedPeripherals = [], refetch: refetchLinked } = useQuery({
    queryKey: ["item-peripherals", selectedItemId],
    queryFn: () => listItemPeripherals(Number(selectedItemId)),
    enabled: !!selectedItemId,
  });

  const { data: availablePeripherals = [] } = useQuery({
    queryKey: ["peripherals", "Disponível"],
    queryFn: () => listPeripherals({ status: "Disponível" }),
  });

  const linkMutation = useMutation({
    mutationFn: (pid: number) => linkPeripheral(Number(selectedItemId), pid),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["item-peripherals", selectedItemId] });
      queryClient.invalidateQueries({ queryKey: ["peripherals"] });
      toast.success("Periférico vinculado!");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao vincular.")),
  });

  const unlinkMutation = useMutation({
    mutationFn: (linkId: number) => unlinkPeripheral(linkId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["item-peripherals", selectedItemId] });
      queryClient.invalidateQueries({ queryKey: ["peripherals"] });
      toast.success("Periférico desvinculado.");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao desvincular.")),
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
      queryClient.invalidateQueries({ queryKey: ["item-peripherals", selectedItemId] });
      queryClient.invalidateQueries({ queryKey: ["peripherals"] });
      setReplacingLinkId(null);
      setReplacingOldId(null);
      setReplaceNewId("");
      setReplaceReason("");
      setReplaceAttachment(null);
      toast.success("Periférico substituído com sucesso!");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao substituir.")),
  });

  const selectedItem = items.find((i) => String(i.id) === selectedItemId);

  return (
    <div className="space-y-6">
      <div className="flex items-center gap-3">
        <Button
          variant="ghost"
          size="icon"
          onClick={() => navigate({ to: "/peripherals" })}
          aria-label="Voltar para periféricos"
        >
          <ArrowLeft className="size-4" aria-hidden />
        </Button>
        <PageHeader
          className="pb-0"
          title="Vincular periféricos"
          description="Associe periféricos a equipamentos como desktops, notebooks, switches e impressoras."
        />
      </div>

      <Section title="Equipamento">
        <div className="flex flex-col gap-2">
          <Label>Selecione o equipamento</Label>
          <div className="max-w-md">
            <SearchableSelect
              options={linkableItems.map((i) => ({
                value: String(i.id),
                label: `#${i.id} — ${i.tipo ?? ""} ${i.brand ?? ""} ${i.model ?? ""}`,
                subtitle: [i.revenda, i.identificador].filter(Boolean).join(" • "),
              }))}
              value={selectedItemId}
              onValueChange={setSelectedItemId}
              placeholder="Selecione ou busque um equipamento…"
              searchPlaceholder="Buscar por ID, tipo, marca, modelo, patrimônio…"
            />
          </div>
          {selectedItem && (
            <p className="text-xs text-muted-foreground">
              Status: <strong className="text-foreground">{selectedItem.status}</strong> · Revenda:{" "}
              <strong className="text-foreground">{selectedItem.revenda}</strong>
            </p>
          )}
        </div>
      </Section>

      {selectedItemId && (
        <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
          <Section
            title={`Vinculados (${linkedPeripherals.length})`}
            actions={
              <Button
                size="sm"
                variant="ghost"
                onClick={() => void refetchLinked()}
                aria-label="Atualizar vinculados"
              >
                <RefreshCw className="size-3.5" aria-hidden />
              </Button>
            }
          >
            {linkedPeripherals.length === 0 ? (
              <EmptyState title="Nenhum periférico vinculado" />
            ) : (
              <div className="space-y-2">
                {linkedPeripherals.map((p) => (
                  <div key={p.link_id} className="flex flex-col gap-1">
                    <PeripheralCard
                      peripheral={p}
                      action={() => unlinkMutation.mutate(p.link_id!)}
                      actionLabel="Desvincular"
                      actionIcon={<Unlink className="size-3.5" aria-hidden />}
                      variant="destructive"
                    />
                    <button
                      type="button"
                      className="ml-1 text-left text-xs text-primary hover:underline"
                      onClick={() => {
                        setReplacingLinkId(p.link_id!);
                        setReplacingOldId(p.id);
                      }}
                    >
                      Substituir…
                    </button>
                  </div>
                ))}
              </div>
            )}
          </Section>

          <Section title={`Disponíveis (${availablePeripherals.length})`}>
            {availablePeripherals.length === 0 ? (
              <EmptyState title="Nenhum periférico disponível" />
            ) : (
              <div className="max-h-[420px] space-y-2 overflow-y-auto pr-1">
                {availablePeripherals.map((p) => (
                  <PeripheralCard
                    key={p.id}
                    peripheral={p}
                    action={() => linkMutation.mutate(p.id)}
                    actionLabel="Vincular"
                    actionIcon={<Link2 className="size-3.5" aria-hidden />}
                  />
                ))}
              </div>
            )}
          </Section>
        </div>
      )}

      {replacingLinkId && replacingOldId && (
        <Section
          title={`Substituir periférico #${replacingOldId}`}
          className="border-warning/30 bg-warning/5"
        >
          <div className="space-y-4">
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
              <div className="flex flex-col gap-1.5">
                <Label>Novo periférico</Label>
                <Select value={replaceNewId} onValueChange={setReplaceNewId}>
                  <SelectTrigger>
                    <SelectValue placeholder="Selecione o substituto" />
                  </SelectTrigger>
                  <SelectContent>
                    {availablePeripherals.map((p) => (
                      <SelectItem key={p.id} value={String(p.id)}>
                        #{p.id} — {p.tipo} {p.brand} {p.model}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="flex flex-col gap-1.5">
                <Label>Motivo da substituição *</Label>
                <Input
                  value={replaceReason}
                  onChange={(e) => setReplaceReason(e.target.value)}
                  placeholder="Ex: Defeito, Upgrade…"
                />
              </div>
            </div>
            <FileUpload
              accept="application/pdf,image/*"
              onFile={setReplaceAttachment}
              label="Comprovante (opcional)"
            />
            <div className="flex gap-3">
              <Button
                disabled={!replaceNewId || !replaceReason || replaceMutation.isPending}
                onClick={() => replaceMutation.mutate()}
              >
                {replaceMutation.isPending ? "Substituindo…" : "Confirmar substituição"}
              </Button>
              <Button
                variant="ghost"
                onClick={() => {
                  setReplacingLinkId(null);
                  setReplacingOldId(null);
                }}
              >
                Cancelar
              </Button>
            </div>
          </div>
        </Section>
      )}
    </div>
  );
}
