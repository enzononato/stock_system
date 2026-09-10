import { useState, type FormEvent } from "react";
import { useNavigate } from "@tanstack/react-router";
import { useMutation, useQueryClient } from "@tanstack/react-query";
import { ArrowLeft, Check, Info, PackagePlus, ShieldAlert, Sparkles } from "lucide-react";
import { toast } from "sonner";

import { createItem } from "@/api/items";
import { TypeSpecificFields, validateTypeSpecificFields } from "@/components/app/TypeSpecificFields";
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
import { useConstants } from "@/hooks/useConstants";
import { getErrorMessage } from "@/lib/api-error";
import { isValidNotaFiscal, maskNotaFiscalInput } from "@/lib/utils";

function todayBr(): string {
  const now = new Date();
  return `${String(now.getDate()).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
}

export function RegisterItemPage() {
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

  const mutation = useMutation({
    mutationFn: (data: Record<string, unknown>) => createItem(data),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      toast.success("Equipamento registrado com sucesso no acervo.");
      void navigate({ to: "/stock" });
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, "Erro ao registrar equipamento."));
    },
  });

  function handleSubmit(e: FormEvent) {
    e.preventDefault();

    if (notaFiscal && !isValidNotaFiscal(notaFiscal)) {
      toast.error("Nota fiscal inválida. Deve conter 9 dígitos.");
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
      date_registered: dateRegistered,
      ...specificFields,
    };
    mutation.mutate(data);
  }

  return (
    <div className="space-y-6">
      {/* Topo Editorial com Ações e Voltar */}
      <div className="border-b border-border pb-4 flex items-center justify-between">
        <div>
          <span className="text-caption text-muted-foreground font-mono">
            INVENTÁRIO · INGRESSO DE PATRIMÔNIO
          </span>
          <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
            Cadastrar Novo Ativo
          </h1>
        </div>
        <Button
          variant="outline"
          size="sm"
          onClick={() => void navigate({ to: "/stock" })}
          className="text-xs"
        >
          <ArrowLeft className="mr-1.5 size-3.5" />
          Voltar ao Estoque
        </Button>
      </div>

      {/* COMPOSIÇÃO 70% / 30% (Seção 25 do PROMPT 3.md) */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
        {/* Formulário Principal Estruturado em Fieldsets (70% - 8 colunas) */}
        <form onSubmit={handleSubmit} className="lg:col-span-8 space-y-6">
          {/* Seção 1: Equipamento e Categoria */}
          <fieldset className="bg-surface border border-border rounded-md p-5 space-y-4">
            <legend className="px-2 text-caption font-semibold text-foreground border-b border-border w-full pb-2">
              1. Dados do Equipamento
            </legend>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 pt-1">
              <div className="flex flex-col gap-1.5">
                <Label className="text-caption">Tipo de Equipamento *</Label>
                <Select value={tipo} onValueChange={setTipo} required disabled={constantsLoading}>
                  <SelectTrigger className="text-xs">
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
                <Label htmlFor="brand" className="text-caption">Marca / Fabricante *</Label>
                <Input
                  id="brand"
                  value={brand}
                  onChange={(e) => setBrand(e.target.value)}
                  placeholder="Ex.: Dell, Lenovo, Apple"
                  className="text-xs"
                  required
                />
              </div>

              <div className="flex flex-col gap-1.5">
                <Label htmlFor="model" className="text-caption">Modelo *</Label>
                <Input
                  id="model"
                  value={model}
                  onChange={(e) => setModel(e.target.value)}
                  placeholder="Ex.: Latitude 5420"
                  className="text-xs"
                  required
                />
              </div>
            </div>
          </fieldset>

          {/* Seção 2: Localização e Aquisição */}
          <fieldset className="bg-surface border border-border rounded-md p-5 space-y-4">
            <legend className="px-2 text-caption font-semibold text-foreground border-b border-border w-full pb-2">
              2. Localização e Dados de Aquisição
            </legend>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 pt-1">
              <div className="flex flex-col gap-1.5">
                <Label className="text-caption">Unidade / Filial *</Label>
                <Select value={revenda} onValueChange={setRevenda} required disabled={constantsLoading}>
                  <SelectTrigger className="text-xs">
                    <SelectValue placeholder="Selecione a unidade" />
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
                <Label htmlFor="nota_fiscal" className="text-caption">Nota Fiscal</Label>
                <Input
                  id="nota_fiscal"
                  value={notaFiscal}
                  onChange={(e) => setNotaFiscal(maskNotaFiscalInput(e.target.value))}
                  placeholder="9 dígitos numéricos"
                  inputMode="numeric"
                  className="font-mono text-xs"
                />
              </div>

              <div className="flex flex-col gap-1.5">
                <Label htmlFor="fornecedor" className="text-caption">Fornecedor / Distribuidor</Label>
                <Input
                  id="fornecedor"
                  value={fornecedor}
                  onChange={(e) => setFornecedor(e.target.value)}
                  placeholder="Razão social do parceiro"
                  className="text-xs"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 pt-2">
              <div className="flex flex-col gap-1.5">
                <Label htmlFor="date_registered" className="text-caption">Data de Entrada *</Label>
                <Input
                  id="date_registered"
                  value={dateRegistered}
                  onChange={(e) => setDateRegistered(e.target.value)}
                  placeholder="dd/mm/aaaa"
                  className="font-mono text-xs"
                  required
                />
              </div>
            </div>
          </fieldset>

          {/* Seção 3: Especificações Técnicas Conforme Tipo */}
          {tipo && (
            <fieldset className="bg-surface border border-border rounded-md p-5 space-y-4">
              <legend className="px-2 text-caption font-semibold text-foreground border-b border-border w-full pb-2">
                3. Especificações Técnicas — {tipo}
              </legend>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 pt-1">
                <TypeSpecificFields
                  tipo={tipo}
                  values={specificFields}
                  onChange={(k, v) => setSpecificFields((prev) => ({ ...prev, [k]: v }))}
                />
              </div>
            </fieldset>
          )}

          <div className="flex items-center justify-between pt-4 border-t border-border">
            <span className="text-xs text-muted-foreground">
              * Campos de preenchimento obrigatório pelo protocolo de patrimônio.
            </span>
            <Button type="submit" disabled={mutation.isPending} className="text-xs px-5">
              <PackagePlus className="mr-2 size-4" />
              {mutation.isPending ? "Registrando Ativo…" : "Efetivar Cadastro"}
            </Button>
          </div>
        </form>

        {/* Coluna Lateral de Contexto e Regras (30% - 4 colunas) */}
        <aside className="lg:col-span-4 space-y-5">
          {/* Resumo do Registro em Tempo Real */}
          <div className="bg-surface border border-border rounded-md p-5 space-y-3">
            <div className="flex items-center justify-between border-b border-border pb-2">
              <h2 className="text-caption font-semibold text-foreground">Ficha do Novo Item</h2>
              <span className="text-[10px] font-mono text-muted-foreground uppercase">Prévia</span>
            </div>

            <div className="space-y-2 text-xs">
              <div className="flex justify-between py-1 border-b border-border/50">
                <span className="text-muted-foreground">Equipamento:</span>
                <span className="font-medium text-foreground text-right truncate max-w-[150px]">
                  {brand || model ? `${brand} ${model}`.trim() : "—"}
                </span>
              </div>
              <div className="flex justify-between py-1 border-b border-border/50">
                <span className="text-muted-foreground">Categoria:</span>
                <span className="font-medium text-foreground">{tipo || "—"}</span>
              </div>
              <div className="flex justify-between py-1 border-b border-border/50">
                <span className="text-muted-foreground">Unidade Alvo:</span>
                <span className="font-medium text-foreground">{revenda || "—"}</span>
              </div>
              <div className="flex justify-between py-1 border-b border-border/50">
                <span className="text-muted-foreground">Patrimônio / Serial:</span>
                <span className="font-mono num font-semibold text-foreground">
                  {specificFields["identificador"] || "—"}
                </span>
              </div>
              <div className="flex justify-between py-1 border-b border-border/50">
                <span className="text-muted-foreground">Nota Fiscal:</span>
                <span className="font-mono num text-foreground">{notaFiscal || "—"}</span>
              </div>
              <div className="flex justify-between py-1">
                <span className="text-muted-foreground">Status Inicial:</span>
                <span className="text-caption px-1.5 py-0.5 rounded-sm border border-border text-foreground font-semibold">
                  ○ DISPONÍVEL
                </span>
              </div>
            </div>
          </div>

          {/* Governança e Diretrizes de TI */}
          <div className="bg-surface-alt/40 border border-border rounded-md p-4 space-y-2.5 text-xs text-muted-foreground">
            <h3 className="font-semibold text-foreground text-caption flex items-center gap-1.5">
              <Info className="size-3.5" />
              Diretrizes de Cadastramento
            </h3>
            <ul className="space-y-1.5 list-disc pl-4 leading-relaxed">
              <li>
                Equipamentos cadastrados ingressam imediatamente com o status <strong>Disponível</strong>.
              </li>
              <li>
                Para computadores e notebooks, o preenchimento de Hostname e MAC é mandatório para inventário de rede.
              </li>
              <li>
                Após a confirmação, o ativo receberá um ID único no banco de dados e ficará visível no Estoque.
              </li>
            </ul>
          </div>
        </aside>
      </div>
    </div>
  );
}
