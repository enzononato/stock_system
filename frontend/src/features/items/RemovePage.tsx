import { useState, type FormEvent } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { Trash2 } from "lucide-react";
import { toast } from "sonner";

import { listItemsPaginated, removeItem } from "@/api/items";
import { PageHeader } from "@/components/app/PageHeader";
import { SearchableSelect } from "@/components/app/SearchableSelect";
import { FileUpload } from "@/components/app/FileUpload";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useConstants } from "@/hooks/useConstants";
import { getErrorMessage } from "@/lib/api-error";

// Sem paginação nesta tela (filtra "Disponível" client-side a partir da lista
// completa) — usamos um teto acima do default de página do backend (50) para
// não truncar a lista de opções.
const FETCH_ALL_LIMIT = 500;

export function RemovePage() {
  const queryClient = useQueryClient();
  const { removalReasons, removalReasonsAttachment, isLoading: constantsLoading } = useConstants();
  const [selectedId, setSelectedId] = useState("");
  const [reason, setReason] = useState("");
  const [attachment, setAttachment] = useState<File | null>(null);

  const { data } = useQuery({
    queryKey: ["items", "remove-all"],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  });
  const items = data?.items ?? [];
  const disponiveis = items.filter((i) => i.status === "Disponível");

  // O mapa de quais motivos exigem comprovante vem de /api/constants
  // (removalReasonsAttachment), não de um objeto hardcoded local.
  const needsAttachment = reason ? Boolean(removalReasonsAttachment[reason]) : false;

  const mutation = useMutation({
    mutationFn: () => removeItem(Number(selectedId), reason, attachment ?? undefined),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      setSelectedId("");
      setReason("");
      setAttachment(null);
      toast.success("Item removido do estoque.");
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, "Erro ao remover item."));
    },
  });

  function handleSubmit(e: FormEvent) {
    e.preventDefault();
    if (!selectedId) {
      toast.error("Selecione um item.");
      return;
    }
    if (!reason) {
      toast.error("Selecione o motivo.");
      return;
    }
    if (needsAttachment && !attachment) {
      toast.error("Este motivo exige um comprovante.");
      return;
    }
    mutation.mutate();
  }

  return (
    <div className="mx-auto max-w-xl space-y-6">
      <PageHeader
        eyebrow="Inventário"
        title="Remover equipamento"
        description="Remove permanentemente o item do estoque. Esta ação não pode ser desfeita."
      />

      <form onSubmit={handleSubmit} className="space-y-4 rounded-lg border border-destructive/25 bg-card p-6">
        <div className="flex flex-col gap-1.5">
          <Label>Equipamento *</Label>
          <SearchableSelect
            options={disponiveis.map((i) => ({
              value: String(i.id),
              label: `#${i.id} — ${i.tipo ?? ""} ${i.brand ?? ""} ${i.model ?? ""}`.trim(),
              subtitle: [i.revenda, i.identificador].filter(Boolean).join(" • "),
            }))}
            value={selectedId}
            onValueChange={setSelectedId}
            placeholder="Selecione ou busque um equipamento..."
            searchPlaceholder="Buscar por ID, tipo, marca, modelo, patrimônio..."
          />
        </div>

        <div className="flex flex-col gap-1.5">
          <Label>Motivo da Remoção *</Label>
          <Select value={reason} onValueChange={setReason} required disabled={constantsLoading}>
            <SelectTrigger>
              <SelectValue placeholder="Selecione o motivo" />
            </SelectTrigger>
            <SelectContent>
              {removalReasons.map((r) => (
                <SelectItem key={r} value={r}>
                  {r}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>

        {reason && (
          <div className="flex flex-col gap-1.5">
            <Label>Comprovante {needsAttachment ? "(obrigatório)" : "(opcional)"}</Label>
            <FileUpload
              accept="application/pdf,image/jpeg,image/png"
              onFile={setAttachment}
              label="Arraste ou clique para anexar comprovante"
            />
          </div>
        )}

        <Button type="submit" variant="destructive" disabled={mutation.isPending} className="w-full">
          <Trash2 className="mr-2 size-4" />
          {mutation.isPending ? "Removendo..." : "Confirmar Remoção"}
        </Button>
      </form>
    </div>
  );
}
