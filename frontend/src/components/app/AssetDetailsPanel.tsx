import { useState, type FormEvent } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { Link } from "@tanstack/react-router";
import {
  FileDown,
  Pencil,
  Save,
  Loader2,
  ExternalLink,
  Keyboard,
  Shield,
} from "lucide-react";
import { toast } from "sonner";

import type { Item } from "@/api/items";
import { updateItem } from "@/api/items";
import { downloadSignedTerm } from "@/api/loans";
import { generateAndDownloadLoanTerm } from "@/components/app/ConfirmacaoTermo";
import { listItemPeripherals, type Peripheral } from "@/api/peripherals";
import {
  Sheet,
  SheetContent,
  SheetHeader,
  SheetTitle,
  SheetDescription,
} from "@/components/ui/sheet";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useAuth } from "@/lib/auth";
import { formatDate, maskNotaFiscalInput, isValidNotaFiscal } from "@/lib/utils";
import { getErrorMessage } from "@/lib/api-error";

interface AssetDetailsPanelProps {
  item: Item | null;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

function DetailRow({
  label,
  value,
  mono,
}: {
  label: string;
  value?: string | null | undefined;
  mono?: boolean | undefined;
}) {
  return (
    <div className="flex items-baseline justify-between py-1.5 border-b border-border/50 text-xs">
      <span className="text-muted-foreground uppercase text-[11px] font-medium tracking-wide">
        {label}
      </span>
      <span className={`text-foreground font-medium text-right ${mono ? "font-mono num" : ""}`}>
        {value || "—"}
      </span>
    </div>
  );
}

export function AssetDetailsPanel({ item, open, onOpenChange }: AssetDetailsPanelProps) {
  const { hasRole } = useAuth();
  const queryClient = useQueryClient();
  const [editing, setEditing] = useState(false);
  const [downloadingTerm, setDownloadingTerm] = useState(false);

  // Campos editáveis
  const [editBrand, setEditBrand] = useState("");
  const [editModel, setEditModel] = useState("");
  const [editRevenda, setEditRevenda] = useState("");
  const [editNotaFiscal, setEditNotaFiscal] = useState("");
  const [editFornecedor, setEditFornecedor] = useState("");
  const [editIdentificador, setEditIdentificador] = useState("");
  const [editHost, setEditHost] = useState("");
  const [editCpu, setEditCpu] = useState("");
  const [editRam, setEditRam] = useState("");
  const [editStorage, setEditStorage] = useState("");
  const [editSistema, setEditSistema] = useState("");
  const [editMac, setEditMac] = useState("");
  const [editIp, setEditIp] = useState("");

  const itemId = item?.id;

  const { data: peripherals = [], isLoading: peripheralsLoading } = useQuery({
    queryKey: ["item-peripherals", itemId],
    queryFn: () => listItemPeripherals(itemId as number),
    enabled: Boolean(itemId && open),
  });

  const updateMutation = useMutation({
    mutationFn: (data: Record<string, unknown>) => updateItem(itemId as number, data),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      toast.success("Patrimônio atualizado com sucesso!");
      setEditing(false);
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, "Erro ao atualizar patrimônio."));
    },
  });

  function startEdit() {
    if (!item) return;
    setEditBrand(item.brand ?? "");
    setEditModel(item.model ?? "");
    setEditRevenda(item.revenda ?? "");
    setEditNotaFiscal(item.nota_fiscal ?? "");
    setEditFornecedor(item.fornecedor ?? "");
    setEditIdentificador(item.identificador ?? "");
    setEditHost(item.host ?? "");
    setEditCpu(item.cpu ?? "");
    setEditRam(item.ram ?? "");
    setEditStorage(item.storage ?? "");
    setEditSistema(item.sistema ?? "");
    setEditMac(item.mac ?? "");
    setEditIp(item.ip ?? "");
    setEditing(true);
  }

  function handleSave(e: FormEvent) {
    e.preventDefault();
    if (editNotaFiscal && !isValidNotaFiscal(editNotaFiscal)) {
      toast.error("Nota fiscal inválida. Deve conter 9 dígitos.");
      return;
    }
    updateMutation.mutate({
      brand: editBrand,
      model: editModel,
      revenda: editRevenda,
      nota_fiscal: editNotaFiscal,
      fornecedor: editFornecedor,
      identificador: editIdentificador,
      host: editHost,
      cpu: editCpu,
      ram: editRam,
      storage: editStorage,
      sistema: editSistema,
      mac: editMac,
      ip: editIp,
    });
  }

  async function handleDownloadSigned() {
    if (!itemId) return;
    setDownloadingTerm(true);
    try {
      const blob = await downloadSignedTerm(itemId);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `termo_assinado_item_${itemId}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 10000);
    } catch (err) {
      toast.error(getErrorMessage(err, "Termo assinado não disponível para este item."));
    } finally {
      setDownloadingTerm(false);
    }
  }

  if (!item) return null;

  return (
    <Sheet open={open} onOpenChange={onOpenChange}>
      <SheetContent
        side="right"
        className="w-full sm:max-w-xl md:max-w-2xl overflow-y-auto bg-surface border-l border-border p-6 sm:p-8 space-y-6"
      >
        <SheetHeader className="border-b border-border pb-4 text-left">
          <div className="flex items-center justify-between">
            <span className="text-caption text-muted-foreground font-mono num">
              PATRIMÔNIO #{item.id}
            </span>
            <span className="text-caption px-2 py-0.5 rounded-sm border border-border-strong text-foreground font-semibold">
              {item.status === "Disponível" ? "● DISPONÍVEL" : item.status?.toUpperCase() ?? "● ATIVO"}
            </span>
          </div>
          <SheetTitle className="text-heading font-semibold text-foreground mt-1">
            {item.brand ?? ""} {item.model ?? "Equipamento"}
          </SheetTitle>
          <SheetDescription className="text-body-sm text-secondary">
            {item.tipo} · {item.revenda || "Unidade não associada"}
          </SheetDescription>
        </SheetHeader>

        {editing ? (
          <form onSubmit={handleSave} className="space-y-4">
            <div className="grid grid-cols-2 gap-3">
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Marca</Label>
                <Input value={editBrand} onChange={(e) => setEditBrand(e.target.value)} />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Modelo</Label>
                <Input value={editModel} onChange={(e) => setEditModel(e.target.value)} />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Patrimônio / Identificador</Label>
                <Input
                  value={editIdentificador}
                  onChange={(e) => setEditIdentificador(e.target.value)}
                  className="font-mono"
                />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Nota Fiscal</Label>
                <Input
                  value={editNotaFiscal}
                  onChange={(e) => setEditNotaFiscal(maskNotaFiscalInput(e.target.value))}
                  className="font-mono"
                />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Host / Hostname</Label>
                <Input value={editHost} onChange={(e) => setEditHost(e.target.value)} />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Sistema Operacional</Label>
                <Input value={editSistema} onChange={(e) => setEditSistema(e.target.value)} />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Processador (CPU)</Label>
                <Input value={editCpu} onChange={(e) => setEditCpu(e.target.value)} />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Memória RAM</Label>
                <Input value={editRam} onChange={(e) => setEditRam(e.target.value)} />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Armazenamento</Label>
                <Input value={editStorage} onChange={(e) => setEditStorage(e.target.value)} />
              </div>
              <div className="flex flex-col gap-1">
                <Label className="text-caption">Endereço MAC</Label>
                <Input
                  value={editMac}
                  onChange={(e) => setEditMac(e.target.value)}
                  className="font-mono"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-4 border-t border-border">
              <Button type="button" variant="outline" size="sm" onClick={() => setEditing(false)}>
                Cancelar
              </Button>
              <Button type="submit" size="sm" disabled={updateMutation.isPending}>
                <Save className="mr-1.5 size-3.5" />
                {updateMutation.isPending ? "Salvando..." : "Salvar Alterações"}
              </Button>
            </div>
          </form>
        ) : (
          <div className="space-y-6">
            {/* Seção 1: Identificação & Localização */}
            <div className="space-y-1">
              <div className="flex items-center justify-between pb-1 border-b border-border">
                <h3 className="text-caption text-foreground font-semibold">Identificação e Registro</h3>
                {hasRole("Gestor") && (
                  <Button variant="ghost" size="sm" className="h-6 text-xs px-2" onClick={startEdit}>
                    <Pencil className="mr-1 size-3" /> Editar
                  </Button>
                )}
              </div>
              <DetailRow label="ID do Sistema" value={`#${item.id}`} mono />
              <DetailRow label="Identificador / Patrimônio" value={item.identificador} mono />
              <DetailRow label="Unidade / Revenda" value={item.revenda} />
              <DetailRow label="Nota Fiscal" value={item.nota_fiscal} mono />
              <DetailRow label="Fornecedor" value={item.fornecedor} />
              <DetailRow label="Data de Cadastro" value={formatDate(item.date_registered)} mono />
            </div>

            {/* Seção 2: Especificações Técnicas / Hardware */}
            <div className="space-y-1">
              <h3 className="text-caption text-foreground font-semibold pb-1 border-b border-border">
                Especificações Técnicas
              </h3>
              <DetailRow label="Hostname / Rede" value={item.host} mono />
              <DetailRow label="Processador" value={item.cpu} />
              <DetailRow label="Memória RAM" value={item.ram} />
              <DetailRow label="Armazenamento" value={item.storage} />
              <DetailRow label="Sistema Operacional" value={item.sistema} />
              <DetailRow label="Endereço MAC" value={item.mac} mono />
              <DetailRow label="Endereço IP" value={item.ip} mono />
            </div>

            {/* Seção 3: Responsabilidade e Alocação */}
            <div className="space-y-1">
              <h3 className="text-caption text-foreground font-semibold pb-1 border-b border-border">
                Alocação e Responsável
              </h3>
              <DetailRow label="Colaborador Atribuído" value={item.assigned_to} />
              <DetailRow label="CPF" value={item.cpf} mono />
              <DetailRow label="Cargo" value={item.cargo} />
              <DetailRow label="Setor" value={item.setor} />
              <DetailRow label="Data do Empréstimo" value={formatDate(item.date_issued)} mono />
            </div>

            {/* Seção 4: Periféricos Vinculados */}
            <div className="space-y-2">
              <div className="flex items-center justify-between pb-1 border-b border-border">
                <h3 className="text-caption text-foreground font-semibold">Periféricos Vinculados</h3>
                <Link to="/peripherals" className="text-xs text-muted-foreground hover:text-foreground flex items-center gap-1">
                  Ver todos <ExternalLink className="size-3" />
                </Link>
              </div>

              {peripheralsLoading ? (
                <p className="text-body-sm text-muted-foreground py-2">Carregando periféricos...</p>
              ) : peripherals.length === 0 ? (
                <p className="text-body-sm text-muted-foreground py-1">
                  Nenhum periférico vinculado a este ativo.
                </p>
              ) : (
                <div className="divide-y divide-border/60">
                  {peripherals.map((p: Peripheral) => (
                    <div key={p.id} className="py-2 flex items-center justify-between text-xs">
                      <div className="flex items-center gap-2">
                        <Keyboard className="size-3.5 text-muted-foreground" />
                        <div>
                          <p className="font-medium text-foreground">{p.tipo}</p>
                          <p className="text-muted-foreground">{p.brand} {p.model}</p>
                        </div>
                      </div>
                      <span className="font-mono text-muted-foreground">#{p.id}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Seção 5: Documentos e Termos de Responsabilidade */}
            <div className="space-y-2 pt-2 border-t border-border">
              <h3 className="text-caption text-foreground font-semibold">
                Termos de Responsabilidade
              </h3>
              <div className="flex flex-wrap gap-2">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => generateAndDownloadLoanTerm(item.id)}
                >
                  <FileDown className="mr-1.5 size-3.5" />
                  Gerar Minuta de Termo
                </Button>
                <Button
                  variant="secondary"
                  size="sm"
                  disabled={downloadingTerm}
                  onClick={handleDownloadSigned}
                >
                  {downloadingTerm ? (
                    <Loader2 className="mr-1.5 size-3.5 animate-spin" />
                  ) : (
                    <Shield className="mr-1.5 size-3.5" />
                  )}
                  Baixar Termo Assinado
                </Button>
              </div>
            </div>
          </div>
        )}
      </SheetContent>
    </Sheet>
  );
}
