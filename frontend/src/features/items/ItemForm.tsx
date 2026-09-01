import { useEffect, useState, type FormEvent } from "react";
import { useNavigate } from "@tanstack/react-router";
import { useMutation, useQueryClient } from "@tanstack/react-query";
import { ArrowLeft, Save } from "lucide-react";
import { toast } from "sonner";

import { createItem, updateItem, type Item } from "@/api/items";
import { TypeSpecificFields, validateTypeSpecificFields } from "@/components/app/TypeSpecificFields";
import { PageHeader } from "@/components/app/PageHeader";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useConstants } from "@/hooks/useConstants";
import { getErrorMessage } from "@/lib/api-error";
import { isValidNotaFiscal, maskNotaFiscalInput } from "@/lib/utils";

const SPECIFIC_KEYS = [
  "identificador", "dominio", "host", "endereco_fisico", "cpu", "ram", "storage",
  "sistema", "licenca", "anydesk", "setor", "ip", "mac", "potencia_nominal",
  "autonomia_estimada", "ip_snmp", "codigo_patrimonial", "responsavel",
  "local_instalacao", "poe", "quantidade_portas",
] as const;

function todayBr(): string {
  const now = new Date();
  return `${String(now.getDate()).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
}

interface ItemFormProps {
  mode: "create" | "edit";
  itemId?: number | undefined;
  existingItem?: Item | undefined;
}

export function ItemForm({ mode, itemId, existingItem }: ItemFormProps) {
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const { equipmentTypes, revendas, isLoading: constantsLoading } = useConstants();

  const [tipo, setTipo] = useState("");
  const [brand, setBrand] = useState("");
  const [model, setModel] = useState("");
  const [revenda, setRevenda] = useState("");
  const [notaFiscal, setNotaFiscal] = useState("");
  const [fornecedor, setFornecedor] = useState("");
  const [dateRegistered, setDateRegistered] = useState(todayBr());
  const [specificFields, setSpecificFields] = useState<Record<string, string>>({});
  const [hydrated, setHydrated] = useState(mode === "create");

  useEffect(() => {
    if (mode === "edit" && existingItem && !hydrated) {
      setTipo(existingItem.tipo ?? "");
      setBrand(existingItem.brand ?? "");
      setModel(existingItem.model ?? "");
      setRevenda(existingItem.revenda ?? "");
      setNotaFiscal(existingItem.nota_fiscal ?? "");
      setFornecedor(existingItem.fornecedor ?? "");
      if (existingItem.date_registered) {
        const d = new Date(existingItem.date_registered);
        if (!Number.isNaN(d.getTime())) {
          setDateRegistered(`${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`);
        }
      }
      const fields: Record<string, string> = {};
      const record = existingItem as unknown as Record<string, unknown>;
      for (const k of SPECIFIC_KEYS) {
        const val = record[k];
        if (val) fields[k] = String(val);
      }
      setSpecificFields(fields);
      setHydrated(true);
    }
  }, [mode, existingItem, hydrated]);

  const mutation = useMutation({
    mutationFn: (data: Record<string, unknown>) =>
      mode === "create" ? createItem(data) : updateItem(itemId as number, data),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      toast.success(mode === "create" ? "Item cadastrado com sucesso!" : "Item atualizado com sucesso!");
      void navigate({ to: "/" });
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, "Erro ao salvar item."));
    },
  });

  function handleSubmit(e: FormEvent) {
    e.preventDefault();

    // O backend rejeita nota fiscal fora de 9 dígitos e campos de MAC/IP
    // malformados — validamos aqui antes de enviar. Nota fiscal é opcional
    // (sem *): só valida se algo foi preenchido.
    if (notaFiscal && !isValidNotaFiscal(notaFiscal)) {
      toast.error("Nota fiscal inválida. Informe 9 dígitos.");
      return;
    }
    const specificFieldsError = validateTypeSpecificFields(tipo, specificFields);
    if (specificFieldsError) {
      toast.error(specificFieldsError);
      return;
    }

    const data: Record<string, unknown> = {
      tipo, brand, model, revenda,
      nota_fiscal: notaFiscal,
      fornecedor,
      ...specificFields,
    };
    if (mode === "create") data["date_registered"] = dateRegistered;
    mutation.mutate(data);
  }

  return (
    <div className="mx-auto max-w-4xl space-y-6">
      <PageHeader
        eyebrow="Inventário"
        title={mode === "create" ? "Cadastrar item" : "Editar item"}
        description={
          mode === "create"
            ? "Registre um novo equipamento no estoque de TI."
            : "Atualize as informações do equipamento selecionado."
        }
        actions={
          <Button variant="outline" size="sm" onClick={() => void navigate({ to: "/" })}>
            <ArrowLeft className="mr-2 size-4" />
            Voltar
          </Button>
        }
      />

      <form onSubmit={handleSubmit} className="space-y-4 rounded-lg border border-border bg-card p-6">
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
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
            <Label htmlFor="brand">Marca *</Label>
            <Input id="brand" value={brand} onChange={(e) => setBrand(e.target.value)} required />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="model">Modelo *</Label>
            <Input id="model" value={model} onChange={(e) => setModel(e.target.value)} required />
          </div>
        </div>

        <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
          <div className="flex flex-col gap-1.5">
            <Label>Revenda *</Label>
            <Select value={revenda} onValueChange={setRevenda} required disabled={constantsLoading}>
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
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="nota_fiscal">Nota Fiscal</Label>
            <Input
              id="nota_fiscal"
              value={notaFiscal}
              onChange={(e) => setNotaFiscal(maskNotaFiscalInput(e.target.value))}
              placeholder="9 dígitos"
              inputMode="numeric"
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="fornecedor">Fornecedor</Label>
            <Input id="fornecedor" value={fornecedor} onChange={(e) => setFornecedor(e.target.value)} />
          </div>
        </div>

        {mode === "create" && (
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="date_registered">Data de Cadastro *</Label>
              <Input
                id="date_registered"
                value={dateRegistered}
                onChange={(e) => setDateRegistered(e.target.value)}
                placeholder="dd/mm/aaaa"
                required
              />
            </div>
          </div>
        )}

        {tipo && (
          <div className="space-y-4 border-t border-border pt-4">
            <p className="text-sm font-medium text-muted-foreground">Informações específicas — {tipo}</p>
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
              <TypeSpecificFields
                tipo={tipo}
                values={specificFields}
                onChange={(k, v) => setSpecificFields((prev) => ({ ...prev, [k]: v }))}
              />
            </div>
          </div>
        )}

        <div className="flex justify-end pt-2">
          <Button type="submit" disabled={mutation.isPending}>
            <Save className="mr-2 size-4" />
            {mutation.isPending ? "Salvando..." : mode === "create" ? "Cadastrar" : "Salvar Alterações"}
          </Button>
        </div>
      </form>
    </div>
  );
}
