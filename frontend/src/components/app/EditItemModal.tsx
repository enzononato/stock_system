import { useEffect, useState, type FormEvent } from "react";
import { useMutation, useQueryClient } from "@tanstack/react-query";
import { Loader2, Save } from "lucide-react";
import { toast } from "sonner";

import { updateItem, type Item } from "@/api/items";
import {
  TypeSpecificFields,
  validateTypeSpecificFields,
} from "@/components/app/TypeSpecificFields";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { useConstants } from "@/hooks/useConstants";
import { getErrorMessage } from "@/lib/api-error";
import { isValidNotaFiscal, maskNotaFiscalInput } from "@/lib/utils";

const SPECIFIC_KEYS = [
  "identificador",
  "dominio",
  "host",
  "endereco_fisico",
  "cpu",
  "ram",
  "storage",
  "sistema",
  "licenca",
  "anydesk",
  "setor",
  "ip",
  "mac",
  "potencia_nominal",
  "autonomia_estimada",
  "ip_snmp",
  "codigo_patrimonial",
  "responsavel",
  "local_instalacao",
  "poe",
  "quantidade_portas",
] as const;

interface EditItemModalProps {
  item: Item | null;
  onClose: () => void;
}

export function EditItemModal({ item, onClose }: EditItemModalProps) {
  const queryClient = useQueryClient();
  const { equipmentTypes, revendas, isLoading: constantsLoading } = useConstants();

  const [tipo, setTipo] = useState("");
  const [brand, setBrand] = useState("");
  const [model, setModel] = useState("");
  const [revenda, setRevenda] = useState("");
  const [notaFiscal, setNotaFiscal] = useState("");
  const [fornecedor, setFornecedor] = useState("");
  const [specificFields, setSpecificFields] = useState<Record<string, string>>({});

  useEffect(() => {
    if (item) {
      setTipo(item.tipo ?? "");
      setBrand(item.brand ?? "");
      setModel(item.model ?? "");
      setRevenda(item.revenda ?? "");
      setNotaFiscal(item.nota_fiscal ?? "");
      setFornecedor(item.fornecedor ?? "");

      const fields: Record<string, string> = {};
      const record = item as unknown as Record<string, unknown>;
      for (const k of SPECIFIC_KEYS) {
        const val = record[k];
        if (val != null && val !== "") {
          fields[k] = String(val);
        }
      }
      setSpecificFields(fields);
    }
  }, [item]);

  const mutation = useMutation({
    mutationFn: (data: Record<string, unknown>) => {
      if (!item) throw new Error("Nenhum item selecionado.");
      return updateItem(item.id, data);
    },
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      toast.success("Item atualizado com sucesso!");
      onClose();
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, "Erro ao atualizar item."));
    },
  });

  function handleSubmit(e: FormEvent) {
    e.preventDefault();

    if (!tipo || !brand.trim() || !model.trim() || !revenda) {
      toast.error("Preencha todos os campos obrigatórios (*).");
      return;
    }

    if (notaFiscal && !isValidNotaFiscal(notaFiscal)) {
      toast.error("Nota fiscal inválida. Informe exatamente 9 dígitos.");
      return;
    }

    const specificFieldsError = validateTypeSpecificFields(tipo, specificFields);
    if (specificFieldsError) {
      toast.error(specificFieldsError);
      return;
    }

    const data: Record<string, unknown> = {
      tipo,
      brand: brand.trim(),
      model: model.trim(),
      revenda,
      nota_fiscal: notaFiscal || null,
      fornecedor: fornecedor.trim() || null,
      ...specificFields,
    };

    mutation.mutate(data);
  }

  return (
    <Dialog open={!!item} onOpenChange={(open) => !open && onClose()}>
      <DialogContent className="max-h-[90vh] overflow-y-auto sm:max-w-2xl">
        <DialogHeader>
          <DialogTitle>Editar Equipamento #{item?.id}</DialogTitle>
          <DialogDescription>
            Atualize as informações do ativo. Os dados serão sincronizados com o estoque.
          </DialogDescription>
        </DialogHeader>

        <form onSubmit={handleSubmit} className="space-y-4 py-2">
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
            <div className="flex flex-col gap-1.5">
              <Label>Tipo *</Label>
              <Select value={tipo} onValueChange={setTipo} required disabled={constantsLoading}>
                <SelectTrigger>
                  <SelectValue placeholder="Selecione o tipo" />
                </SelectTrigger>
                <SelectContent>
                  {equipmentTypes.map((t) => (
                    <SelectItem key={t} value={t}>
                      {t}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="flex flex-col gap-1.5">
              <Label>Revenda / Unidade *</Label>
              <Select
                value={revenda}
                onValueChange={setRevenda}
                required
                disabled={constantsLoading}
              >
                <SelectTrigger>
                  <SelectValue placeholder="Selecione a revenda" />
                </SelectTrigger>
                <SelectContent>
                  {revendas.map((r) => (
                    <SelectItem key={r} value={r}>
                      {r}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
          </div>

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="edit-brand">Marca *</Label>
              <Input
                id="edit-brand"
                value={brand}
                onChange={(e) => setBrand(e.target.value)}
                required
                placeholder="Ex.: Dell, Lenovo, HP"
              />
            </div>

            <div className="flex flex-col gap-1.5">
              <Label htmlFor="edit-model">Modelo *</Label>
              <Input
                id="edit-model"
                value={model}
                onChange={(e) => setModel(e.target.value)}
                required
                placeholder="Ex.: Latitude 3420, ThinkPad E14"
              />
            </div>
          </div>

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="edit-nota-fiscal">Nota Fiscal</Label>
              <Input
                id="edit-nota-fiscal"
                value={notaFiscal}
                onChange={(e) => setNotaFiscal(maskNotaFiscalInput(e.target.value))}
                placeholder="9 dígitos"
                inputMode="numeric"
              />
            </div>

            <div className="flex flex-col gap-1.5">
              <Label htmlFor="edit-fornecedor">Fornecedor</Label>
              <Input
                id="edit-fornecedor"
                value={fornecedor}
                onChange={(e) => setFornecedor(e.target.value)}
                placeholder="Ex.: Kalunga, Dell Brasil"
              />
            </div>
          </div>

          {tipo && (
            <div className="space-y-3 rounded-md border border-border/70 bg-muted/30 p-3.5">
              <p className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                Especificações Técnicas — {tipo}
              </p>
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                <TypeSpecificFields
                  tipo={tipo}
                  values={specificFields}
                  onChange={(k, v) => setSpecificFields((prev) => ({ ...prev, [k]: v }))}
                />
              </div>
            </div>
          )}

          <DialogFooter className="pt-3 gap-2">
            <Button type="button" variant="outline" onClick={onClose} disabled={mutation.isPending}>
              Cancelar
            </Button>
            <Button type="submit" disabled={mutation.isPending}>
              {mutation.isPending ? (
                <>
                  <Loader2 className="mr-2 size-4 animate-spin" />
                  Confirmando...
                </>
              ) : (
                <>
                  <Save className="mr-2 size-4" />
                  Confirmar
                </>
              )}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}
