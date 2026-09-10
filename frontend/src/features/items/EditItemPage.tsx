import { useEffect, useMemo, useState, type FormEvent } from "react";
import { useParams, useNavigate } from "@tanstack/react-router";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { ArrowLeft, Check, RefreshCw, Save, AlertCircle } from "lucide-react";
import { toast } from "sonner";

import { getItem, updateItem, type Item } from "@/api/items";
import { useConstants } from "@/hooks/useConstants";
import { getErrorMessage } from "@/lib/api-error";
import { isValidNotaFiscal } from "@/lib/utils";
import { PageHeader } from "@/components/app/PageHeader";
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
  TypeSpecificFields,
  validateTypeSpecificFields,
} from "@/components/app/TypeSpecificFields";
import { ErrorState, LoadingState } from "@/components/app/StateBlocks";

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

export function EditItemPage() {
  const { id } = useParams({ from: "/_shell/edit/$id" });
  const itemId = Number(id);
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const { equipmentTypes, revendas, isLoading: constantsLoading } = useConstants();

  const {
    data: item,
    isLoading,
    isError,
    error,
    refetch,
  } = useQuery({
    queryKey: ["item", itemId],
    queryFn: () => getItem(itemId),
    enabled: Number.isFinite(itemId),
  });

  const [tipo, setTipo] = useState("");
  const [brand, setBrand] = useState("");
  const [model, setModel] = useState("");
  const [revenda, setRevenda] = useState("");
  const [notaFiscal, setNotaFiscal] = useState("");
  const [fornecedor, setFornecedor] = useState("");
  const [specificFields, setSpecificFields] = useState<Record<string, string>>({});
  const [hydrated, setHydrated] = useState(false);

  useEffect(() => {
    if (item && !hydrated) {
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
        if (val) fields[k] = String(val);
      }
      setSpecificFields(fields);
      setHydrated(true);
    }
  }, [item, hydrated]);

  const changedFields = useMemo(() => {
    if (!item) return [];
    const changes: { label: string; before: string; after: string }[] = [];

    if (tipo !== (item.tipo ?? ""))
      changes.push({ label: "Tipo", before: item.tipo ?? "—", after: tipo || "—" });
    if (brand !== (item.brand ?? ""))
      changes.push({ label: "Marca", before: item.brand ?? "—", after: brand || "—" });
    if (model !== (item.model ?? ""))
      changes.push({ label: "Modelo", before: item.model ?? "—", after: model || "—" });
    if (revenda !== (item.revenda ?? ""))
      changes.push({ label: "Revenda", before: item.revenda ?? "—", after: revenda || "—" });
    if (notaFiscal !== (item.nota_fiscal ?? ""))
      changes.push({ label: "Nota Fiscal", before: item.nota_fiscal ?? "—", after: notaFiscal || "—" });
    if (fornecedor !== (item.fornecedor ?? ""))
      changes.push({ label: "Fornecedor", before: item.fornecedor ?? "—", after: fornecedor || "—" });

    const record = item as unknown as Record<string, unknown>;
    for (const k of SPECIFIC_KEYS) {
      const prevVal = String(record[k] ?? "");
      const nextVal = String(specificFields[k] ?? "");
      if (prevVal !== nextVal) {
        changes.push({ label: k.replace(/_/g, " "), before: prevVal || "—", after: nextVal || "—" });
      }
    }

    return changes;
  }, [item, tipo, brand, model, revenda, notaFiscal, fornecedor, specificFields]);

  const mutation = useMutation({
    mutationFn: (data: Record<string, unknown>) => updateItem(itemId, data),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      void queryClient.invalidateQueries({ queryKey: ["item", itemId] });
      toast.success("Equipamento atualizado com sucesso!");
      void navigate({ to: "/stock" });
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, "Erro ao salvar alterações no equipamento."));
    },
  });

  function handleSubmit(e: FormEvent) {
    e.preventDefault();

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
      tipo,
      brand,
      model,
      revenda,
      nota_fiscal: notaFiscal,
      fornecedor,
      ...specificFields,
    };
    mutation.mutate(data);
  }

  if (isLoading) return <LoadingState label="Carregando dados do equipamento…" />;
  if (isError || !item) {
    return (
      <ErrorState
        error={error}
        onRetry={() => void refetch()}
        title="Não foi possível carregar o item"
      />
    );
  }

  return (
    <div className="page-container-reading space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Manutenção Patrimonial
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="font-mono text-caption text-foreground font-medium">#{item.id}</span>
          </div>
          <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">
            Editar Equipamento
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Visualização comparativa entre o estado atual em inventário e as alterações propostas.
          </p>
        </div>

        <div className="flex items-center gap-2">
          <Button
            variant="outline"
            size="sm"
            onClick={() => void navigate({ to: "/stock" })}
            className="rounded-[4px] border-border text-foreground hover:bg-muted"
          >
            <ArrowLeft className="mr-1.5 size-3.5" />
            Voltar
          </Button>
          <Button
            type="submit"
            form="edit-item-form"
            size="sm"
            disabled={mutation.isPending || changedFields.length === 0}
            className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 font-medium"
          >
            <Save className="mr-1.5 size-3.5" />
            {mutation.isPending ? "Salvando…" : `Salvar (${changedFields.length} alteraç${changedFields.length === 1 ? "ão" : "ões"})`}
          </Button>
        </div>
      </div>

      {/* Grid de Comparação 40/60 */}
      <div className="grid grid-cols-1 lg:grid-cols-[380px_1fr] gap-6 items-start">
        {/* Coluna Esquerda: Estado Atual (Read-only) */}
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-5">
          <div className="border-b border-border pb-3">
            <div className="flex items-center justify-between">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
                Estado Atual Gravado
              </span>
              <Badge variant="outline" className="rounded-[2px] font-mono text-[11px] border-border text-muted-foreground">
                {item.status}
              </Badge>
            </div>
            <p className="text-body-sm font-medium text-foreground mt-1">
              {item.tipo} {item.brand} {item.model}
            </p>
          </div>

          <div className="space-y-3 text-body-sm">
            <div className="flex justify-between py-1 border-b border-border/50">
              <span className="text-muted-foreground">ID Patrimonial</span>
              <span className="font-mono font-medium text-foreground">#{item.id}</span>
            </div>
            <div className="flex justify-between py-1 border-b border-border/50">
              <span className="text-muted-foreground">Revenda</span>
              <span className="text-foreground">{item.revenda || "—"}</span>
            </div>
            <div className="flex justify-between py-1 border-b border-border/50">
              <span className="text-muted-foreground">Nota Fiscal</span>
              <span className="font-mono text-foreground">{item.nota_fiscal || "—"}</span>
            </div>
            <div className="flex justify-between py-1 border-b border-border/50">
              <span className="text-muted-foreground">Fornecedor</span>
              <span className="text-foreground">{item.fornecedor || "—"}</span>
            </div>
            {item.identificador && (
              <div className="flex justify-between py-1 border-b border-border/50">
                <span className="text-muted-foreground">Serial / Identificador</span>
                <span className="font-mono text-foreground">{item.identificador}</span>
              </div>
            )}
            {item.assigned_to && (
              <div className="flex justify-between py-1 border-b border-border/50">
                <span className="text-muted-foreground">Alocado Para</span>
                <span className="font-medium text-foreground">{item.assigned_to}</span>
              </div>
            )}
          </div>

          {/* Diffs Ativos */}
          <div className="pt-2 border-t border-border">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block mb-2">
              Campos Modificados ({changedFields.length})
            </span>
            {changedFields.length === 0 ? (
              <p className="text-caption text-muted-foreground italic">
                Nenhum campo alterado até o momento.
              </p>
            ) : (
              <div className="space-y-2">
                {changedFields.map((change, idx) => (
                  <div key={idx} className="rounded-[4px] border border-border bg-surface-alt p-2 text-caption">
                    <span className="font-semibold text-foreground capitalize block">{change.label}</span>
                    <div className="flex items-center gap-2 mt-1 text-muted-foreground">
                      <span className="line-through text-muted-foreground">{change.before}</span>
                      <span>→</span>
                      <span className="font-semibold text-foreground">{change.after}</span>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* Coluna Direita: Formulário de Alteração */}
        <form
          id="edit-item-form"
          onSubmit={handleSubmit}
          className="rounded-[6px] border border-border bg-surface p-6 space-y-6"
        >
          <div>
            <h2 className="text-body-lg font-semibold text-foreground">
              Formulário de Alteração
            </h2>
            <p className="text-caption text-muted-foreground mt-0.5">
              Edite as propriedades abaixo. As diferenças serão sincronizadas em tempo real com a coluna à esquerda.
            </p>
          </div>

          {/* Seção Básica */}
          <div className="space-y-4 pt-2 border-t border-border">
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="space-y-1.5">
                <Label className="text-caption font-medium text-foreground">Tipo de Equipamento *</Label>
                <Select value={tipo} onValueChange={setTipo} required disabled={constantsLoading}>
                  <SelectTrigger className="rounded-[4px] border-border bg-background text-body-sm h-9">
                    <SelectValue placeholder="Selecione o tipo" />
                  </SelectTrigger>
                  <SelectContent className="rounded-[4px] border-border bg-surface">
                    {equipmentTypes.map((t) => (
                      <SelectItem key={t} value={t} className="text-body-sm rounded-[2px]">
                        {t}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-1.5">
                <Label className="text-caption font-medium text-foreground">Marca *</Label>
                <Input
                  value={brand}
                  onChange={(e) => setBrand(e.target.value)}
                  placeholder="Ex: Dell, Lenovo, HP"
                  required
                  className="rounded-[4px] border-border bg-background text-body-sm h-9"
                />
              </div>

              <div className="space-y-1.5">
                <Label className="text-caption font-medium text-foreground">Modelo *</Label>
                <Input
                  value={model}
                  onChange={(e) => setModel(e.target.value)}
                  placeholder="Ex: Latitude 3420"
                  required
                  className="rounded-[4px] border-border bg-background text-body-sm h-9"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="space-y-1.5">
                <Label className="text-caption font-medium text-foreground">Revenda / Filial *</Label>
                <Select value={revenda} onValueChange={setRevenda} required disabled={constantsLoading}>
                  <SelectTrigger className="rounded-[4px] border-border bg-background text-body-sm h-9">
                    <SelectValue placeholder="Selecione a revenda" />
                  </SelectTrigger>
                  <SelectContent className="rounded-[4px] border-border bg-surface">
                    {revendas.map((r) => (
                      <SelectItem key={r} value={r} className="text-body-sm rounded-[2px]">
                        {r}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-1.5">
                <Label className="text-caption font-medium text-foreground">Nota Fiscal (9 dígitos)</Label>
                <Input
                  value={notaFiscal}
                  onChange={(e) => setNotaFiscal(e.target.value)}
                  placeholder="123456789"
                  maxLength={9}
                  className="font-mono rounded-[4px] border-border bg-background text-body-sm h-9"
                />
              </div>

              <div className="space-y-1.5">
                <Label className="text-caption font-medium text-foreground">Fornecedor</Label>
                <Input
                  value={fornecedor}
                  onChange={(e) => setFornecedor(e.target.value)}
                  placeholder="Nome do fornecedor"
                  className="rounded-[4px] border-border bg-background text-body-sm h-9"
                />
              </div>
            </div>
          </div>

          {/* Especificações Técnicas por Tipo */}
          {tipo && (
            <div className="space-y-4 pt-4 border-t border-border">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-body font-semibold text-foreground">
                    Especificações Técnicas ({tipo})
                  </h3>
                  <p className="text-caption text-muted-foreground">
                    Campos de configuração e identificação de hardware.
                  </p>
                </div>
                <Badge variant="outline" className="rounded-[2px] font-mono text-[11px] border-border">
                  HARDWARE
                </Badge>
              </div>

              <TypeSpecificFields
                tipo={tipo}
                values={specificFields}
                onChange={(k, v) => setSpecificFields((prev) => ({ ...prev, [k]: v }))}
              />
            </div>
          )}

          <div className="pt-4 border-t border-border flex justify-end gap-3">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => void navigate({ to: "/stock" })}
              className="rounded-[4px] border-border"
            >
              Cancelar
            </Button>
            <Button
              type="submit"
              size="sm"
              disabled={mutation.isPending || changedFields.length === 0}
              className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90"
            >
              <Save className="mr-1.5 size-3.5" />
              {mutation.isPending ? "Salvando alterações…" : "Confirmar Alterações"}
            </Button>
          </div>
        </form>
      </div>
    </div>
  );
}
