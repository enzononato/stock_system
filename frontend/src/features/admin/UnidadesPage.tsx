import { useState, type FormEvent } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { Ban, BarChart3, Building2, ChevronDown, ChevronRight, Pencil, Plus, RotateCcw } from "lucide-react";

import {
  listUnidades,
  createUnidade,
  updateUnidade,
  deactivateUnidade,
  reactivateUnidade,
  getIndicadoresUnidade,
  type Unidade,
  type UnidadeInput,
  type IndicadoresUnidade,
} from "@/api/unidades";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
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
import { LoadingState } from "@/components/app/StateBlocks";
import { getErrorMessage } from "@/lib/api-error";
import { toast } from "sonner";

// --- Validação e máscaras --------------------------------------------------

export function isValidCnpj(cnpj: string | null | undefined): boolean {
  if (!cnpj) return false;
  const digits = cnpj.replace(/\D/g, "");
  if (digits.length !== 14) return false;
  if (/^(\d)\1{13}$/.test(digits)) return false;
  const calcCheckDigit = (base: string, weights: number[]): number => {
    let total = 0;
    for (let i = 0; i < base.length; i++) total += Number(base[i]) * weights[i]!;
    const remainder = total % 11;
    return remainder < 2 ? 0 : 11 - remainder;
  };
  const weights1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
  const weights2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
  const dv1 = calcCheckDigit(digits.slice(0, 12), weights1);
  const dv2 = calcCheckDigit(digits.slice(0, 12) + String(dv1), weights2);
  return dv1 === Number(digits[12]) && dv2 === Number(digits[13]);
}

export function maskCnpjInput(value: string): string {
  return value.replace(/\D/g, "").slice(0, 14)
    .replace(/(\d{2})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1/$2")
    .replace(/(\d{4})(\d{1,2})$/, "$1-$2");
}

export function isValidCep(cep: string | null | undefined): boolean {
  if (!cep) return false;
  return /^\d{5}-\d{3}$/.test(cep);
}

export function maskCepInput(value: string): string {
  return value.replace(/\D/g, "").slice(0, 8).replace(/(\d{5})(\d{1,3})$/, "$1-$2");
}

export function isValidUf(uf: string | null | undefined): boolean {
  if (!uf) return false;
  return /^[A-Z]{2}$/.test(uf);
}

export function maskUfInput(value: string): string {
  return value.replace(/[^A-Za-z]/g, "").slice(0, 2).toUpperCase();
}

// --- Indicadores de Unidade ------------------------------------------------

const STATUS_ORDER = ["Disponível", "Indisponível", "Pendente", "Pendente Devolução"];

function IndicadoresPanel({ dados }: { dados: IndicadoresUnidade }) {
  const semMovimento =
    dados.termos_emitidos === 0 &&
    dados.termos_confirmados === 0 &&
    dados.devolucoes_concluidas === 0 &&
    dados.itens_total === 0 &&
    dados.emprestimos_ativos === 0 &&
    dados.perifericos_vinculados === 0;

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3">
        {[
          { label: "Termos Emitidos", value: dados.termos_emitidos },
          { label: "Termos Confirmados", value: dados.termos_confirmados },
          { label: "Devoluções Concluídas", value: dados.devolucoes_concluidas },
          { label: "Total de Itens", value: dados.itens_total },
          { label: "Empréstimos Ativos", value: dados.emprestimos_ativos },
          { label: "Periféricos Vinculados", value: dados.perifericos_vinculados },
        ].map((s) => (
          <div key={s.label} className="rounded-[4px] border border-border bg-surface-alt p-3">
            <p className="text-caption text-muted-foreground">{s.label}</p>
            <p className="font-mono text-body-lg font-bold text-foreground">{s.value}</p>
          </div>
        ))}
      </div>

      <div className="pt-3 border-t border-border space-y-2">
        <p className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
          Itens por Status
        </p>
        <div className="flex flex-wrap gap-2">
          {STATUS_ORDER.map((status) => (
            <div key={status} className="flex items-center gap-1.5 text-caption">
              <span className="font-mono text-muted-foreground">{dados.itens_por_status[status] ?? 0}</span>
              <span className="text-muted-foreground">{status}</span>
            </div>
          ))}
        </div>
      </div>

      {semMovimento && (
        <p className="text-caption italic text-muted-foreground">
          Esta unidade ainda não tem nenhuma movimentação registrada.
        </p>
      )}
    </div>
  );
}

// --- UnidadeRow (Tree Node) ------------------------------------------------

function UnidadeRow({
  unidade,
  onEdit,
  onShowIndicadores,
  onDeactivate,
  onReactivate,
  isPending,
}: {
  unidade: Unidade;
  onEdit: (u: Unidade) => void;
  onShowIndicadores: (id: number) => void;
  onDeactivate: (id: number) => void;
  onReactivate: (id: number) => void;
  isPending: boolean;
}) {
  const [expanded, setExpanded] = useState(false);

  return (
    <div
      className={`rounded-[4px] border transition-all ${
        unidade.is_active !== false ? "border-border" : "border-border/40 opacity-60"
      } bg-surface overflow-hidden`}
    >
      {/* Header do Nó */}
      <div className="flex items-center gap-3 px-4 py-3">
        <button
          type="button"
          onClick={() => setExpanded((p) => !p)}
          className="flex items-center gap-2 flex-1 min-w-0 text-left hover:bg-transparent"
        >
          {expanded ? (
            <ChevronDown className="size-4 text-muted-foreground shrink-0" />
          ) : (
            <ChevronRight className="size-4 text-muted-foreground shrink-0" />
          )}
          <Building2 className="size-4 text-muted-foreground shrink-0" />
          <div className="min-w-0">
            <span className={`text-body-sm font-semibold ${unidade.is_active === false ? "line-through text-muted-foreground" : "text-foreground"}`}>
              {unidade.nome}
            </span>
            {unidade.is_active === false && (
              <Badge variant="outline" className="ml-2 rounded-[2px] text-[10px] font-mono border-border text-muted-foreground">
                × INATIVA
              </Badge>
            )}
          </div>
        </button>

        <div className="flex items-center gap-1 shrink-0">
          <span className="font-mono text-caption text-muted-foreground hidden sm:inline">
            {unidade.cidade && unidade.uf ? `${unidade.cidade}/${unidade.uf}` : ""}
          </span>
          <Button
            size="sm"
            variant="ghost"
            onClick={() => onShowIndicadores(unidade.id)}
            className="size-7 p-0 rounded-[2px] text-muted-foreground hover:text-foreground"
          >
            <BarChart3 className="size-3.5" />
          </Button>
          <Button
            size="sm"
            variant="ghost"
            onClick={() => onEdit(unidade)}
            className="size-7 p-0 rounded-[2px] text-muted-foreground hover:text-foreground"
          >
            <Pencil className="size-3.5" />
          </Button>
          {unidade.is_active !== false ? (
            <AlertDialog>
              <AlertDialogTrigger asChild>
                <Button size="sm" variant="ghost" disabled={isPending} className="size-7 p-0 rounded-[2px] text-muted-foreground hover:text-foreground">
                  <Ban className="size-3.5" />
                </Button>
              </AlertDialogTrigger>
              <AlertDialogContent className="rounded-[6px] border-border bg-surface">
                <AlertDialogHeader>
                  <AlertDialogTitle className="text-body-lg font-semibold">Inativar "{unidade.nome}"?</AlertDialogTitle>
                  <AlertDialogDescription className="text-caption text-muted-foreground">
                    Unidades com itens ativos vinculados não podem ser inativadas.
                  </AlertDialogDescription>
                </AlertDialogHeader>
                <AlertDialogFooter>
                  <AlertDialogCancel className="rounded-[4px] border-border text-xs">Cancelar</AlertDialogCancel>
                  <AlertDialogAction
                    className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs"
                    onClick={() => onDeactivate(unidade.id)}
                  >
                    Inativar
                  </AlertDialogAction>
                </AlertDialogFooter>
              </AlertDialogContent>
            </AlertDialog>
          ) : (
            <Button
              size="sm"
              variant="ghost"
              disabled={isPending}
              onClick={() => onReactivate(unidade.id)}
              className="size-7 p-0 rounded-[2px] text-muted-foreground hover:text-foreground"
            >
              <RotateCcw className="size-3.5" />
            </Button>
          )}
        </div>
      </div>

      {/* Conteúdo expandido: dados jurídicos e endereço */}
      {expanded && (
        <div className="border-t border-border bg-surface-alt px-4 py-3 grid grid-cols-2 sm:grid-cols-4 gap-3 text-caption">
          <div>
            <span className="text-muted-foreground block">Razão Social</span>
            <span className="font-medium text-foreground">{unidade.razao_social}</span>
          </div>
          <div>
            <span className="text-muted-foreground block">CNPJ</span>
            <span className="font-mono text-foreground">{unidade.cnpj}</span>
          </div>
          <div>
            <span className="text-muted-foreground block">CEP</span>
            <span className="font-mono text-foreground">{unidade.cep || "—"}</span>
          </div>
          <div>
            <span className="text-muted-foreground block">Endereço</span>
            <span className="text-foreground">{unidade.endereco || "—"}</span>
          </div>
          <div>
            <span className="text-muted-foreground block">Cidade</span>
            <span className="text-foreground">{unidade.cidade || "—"}</span>
          </div>
          <div>
            <span className="text-muted-foreground block">UF</span>
            <span className="font-mono font-semibold text-foreground">{unidade.uf || "—"}</span>
          </div>
          <div>
            <span className="text-muted-foreground block">ID Interno</span>
            <span className="font-mono text-foreground">#{unidade.id}</span>
          </div>
          <div>
            <span className="text-muted-foreground block">Status</span>
            <span className="font-mono text-foreground">
              {unidade.is_active !== false ? "● ATIVA" : "○ INATIVA"}
            </span>
          </div>
        </div>
      )}
    </div>
  );
}

// --- Página Principal -------------------------------------------------------

const emptyForm = {
  nome: "",
  razaoSocial: "",
  cnpj: "",
  endereco: "",
  cep: "",
  cidade: "",
  uf: "",
};

export function UnidadesPage() {
  const queryClient = useQueryClient();

  const [editingId, setEditingId] = useState<number | null>(null);
  const [originalNome, setOriginalNome] = useState("");
  const [form, setForm] = useState(emptyForm);
  const [showForm, setShowForm] = useState(false);

  const [confirmRenameOpen, setConfirmRenameOpen] = useState(false);
  const [pendingData, setPendingData] = useState<UnidadeInput | null>(null);
  const [indicadoresUnidadeId, setIndicadoresUnidadeId] = useState<number | null>(null);
  const [showInactive, setShowInactive] = useState(false);

  const { data: unidades = [], isLoading, error, refetch } = useQuery({
    queryKey: ["unidades", { includeInactive: showInactive }],
    queryFn: () => listUnidades(showInactive ? { include_inactive: true } : undefined),
  });

  const { data: indicadores, isLoading: indicadoresLoading } = useQuery({
    queryKey: ["unidades", indicadoresUnidadeId, "indicadores"],
    queryFn: () => getIndicadoresUnidade(indicadoresUnidadeId as number),
    enabled: indicadoresUnidadeId !== null,
  });

  function invalidateAfterWrite() {
    void queryClient.invalidateQueries({ queryKey: ["unidades"] });
    void queryClient.invalidateQueries({ queryKey: ["constants"] });
  }

  function resetForm() {
    setEditingId(null);
    setOriginalNome("");
    setForm(emptyForm);
    setShowForm(false);
  }

  const createMutation = useMutation({
    mutationFn: (data: UnidadeInput) => createUnidade(data),
    onSuccess: () => { invalidateAfterWrite(); resetForm(); toast.success("Unidade criada!"); },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao criar unidade.")),
  });

  const updateMutation = useMutation({
    mutationFn: ({ id, data }: { id: number; data: UnidadeInput }) => updateUnidade(id, data),
    onSuccess: () => { invalidateAfterWrite(); resetForm(); toast.success("Unidade atualizada!"); },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao atualizar.")),
  });

  const deactivateMutation = useMutation({
    mutationFn: (id: number) => deactivateUnidade(id),
    onSuccess: () => { invalidateAfterWrite(); toast.success("Unidade inativada."); },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao inativar.")),
  });

  const reactivateMutation = useMutation({
    mutationFn: (id: number) => reactivateUnidade(id),
    onSuccess: () => { invalidateAfterWrite(); toast.success("Unidade reativada."); },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao reativar.")),
  });

  function startEdit(u: Unidade) {
    setEditingId(u.id);
    setOriginalNome(u.nome);
    setForm({
      nome: u.nome,
      razaoSocial: u.razao_social,
      cnpj: u.cnpj,
      endereco: u.endereco ?? "",
      cep: u.cep ?? "",
      cidade: u.cidade ?? "",
      uf: u.uf ?? "",
    });
    setShowForm(true);
  }

  function buildAndValidate(): UnidadeInput | null {
    if (!form.nome.trim()) { toast.error("Informe o nome da unidade."); return null; }
    if (!form.razaoSocial.trim()) { toast.error("Informe a razão social."); return null; }
    if (!isValidCnpj(form.cnpj)) { toast.error("CNPJ inválido."); return null; }
    if (form.cep.trim() && !isValidCep(form.cep)) { toast.error("CEP deve ter o formato 00000-000."); return null; }
    if (form.uf.trim() && !isValidUf(form.uf)) { toast.error("UF deve ter 2 letras."); return null; }
    const data: UnidadeInput = {
      nome: form.nome.trim(),
      razao_social: form.razaoSocial.trim(),
      cnpj: form.cnpj.trim(),
    };
    if (form.endereco.trim()) data.endereco = form.endereco.trim();
    if (form.cep.trim()) data.cep = form.cep.trim();
    if (form.cidade.trim()) data.cidade = form.cidade.trim();
    if (form.uf.trim()) data.uf = form.uf.trim();
    return data;
  }

  function submit(data: UnidadeInput) {
    if (editingId !== null) updateMutation.mutate({ id: editingId, data });
    else createMutation.mutate(data);
  }

  function handleSubmit(e: FormEvent) {
    e.preventDefault();
    const data = buildAndValidate();
    if (!data) return;
    if (editingId !== null && data.nome !== originalNome) {
      setPendingData(data);
      setConfirmRenameOpen(true);
      return;
    }
    submit(data);
  }

  const isSaving = createMutation.isPending || updateMutation.isPending;
  const ativas = unidades.filter((u) => u.is_active !== false);
  const inativas = unidades.filter((u) => u.is_active === false);
  const selectedUnidade = unidades.find((u) => u.id === indicadoresUnidadeId);

  return (
    <div className="page-container-dense space-y-6">
      {/* Header */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Administração Corporativa
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Hierarquia Organizacional</span>
          </div>
          <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">
            Unidades e Filiais
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Gestão de unidades de revenda, dados jurídicos, CNPJ, endereço e indicadores patrimoniais por filial.
          </p>
        </div>

        <div className="flex items-center gap-3">
          <label className="flex items-center gap-2 text-caption text-muted-foreground cursor-pointer select-none">
            <Switch checked={showInactive} onCheckedChange={setShowInactive} />
            <span>Mostrar inativas</span>
          </label>
          <Button
            size="sm"
            onClick={() => { setEditingId(null); setForm(emptyForm); setShowForm((p) => !p); }}
            className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs h-8"
          >
            <Plus className="mr-1.5 size-3.5" />
            Nova Unidade
          </Button>
        </div>
      </div>

      {/* Sumário */}
      <div className="grid grid-cols-3 gap-4">
        {[
          { label: "Total de Unidades", value: unidades.length },
          { label: "Unidades Ativas", value: ativas.length },
          { label: "Unidades Inativas", value: inativas.length },
        ].map((s) => (
          <div key={s.label} className="rounded-[6px] border border-border bg-surface p-4">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">{s.label}</span>
            <span className="font-mono text-heading font-bold text-foreground">{s.value}</span>
          </div>
        ))}
      </div>

      {/* Formulário de Criação/Edição (colapsável) */}
      {showForm && (
        <div className="rounded-[6px] border border-border bg-surface p-6 space-y-5">
          <div className="border-b border-border pb-3">
            <h2 className="text-body-lg font-semibold text-foreground">
              {editingId !== null ? `Editar Unidade #${editingId}` : "Criar Nova Unidade"}
            </h2>
            {editingId !== null && (
              <p className="text-caption text-muted-foreground mt-1">
                ⚠ Alterar o nome da unidade atualiza todas as referências vinculadas em equipamentos e histórico.
              </p>
            )}
          </div>

          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Nome *</Label>
                <Input value={form.nome} onChange={(e) => setForm((f) => ({ ...f, nome: e.target.value }))} required className="rounded-[4px] border-border bg-background text-body-sm h-9" />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Razão Social *</Label>
                <Input value={form.razaoSocial} onChange={(e) => setForm((f) => ({ ...f, razaoSocial: e.target.value }))} required className="rounded-[4px] border-border bg-background text-body-sm h-9" />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">CNPJ *</Label>
                <Input value={form.cnpj} onChange={(e) => setForm((f) => ({ ...f, cnpj: maskCnpjInput(e.target.value) }))} placeholder="00.000.000/0000-00" maxLength={18} required className="font-mono rounded-[4px] border-border bg-background text-body-sm h-9" />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Endereço</Label>
                <Input value={form.endereco} onChange={(e) => setForm((f) => ({ ...f, endereco: e.target.value }))} className="rounded-[4px] border-border bg-background text-body-sm h-9" />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">CEP</Label>
                <Input value={form.cep} onChange={(e) => setForm((f) => ({ ...f, cep: maskCepInput(e.target.value) }))} placeholder="00000-000" maxLength={9} className="font-mono rounded-[4px] border-border bg-background text-body-sm h-9" />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div className="space-y-1.5">
                  <Label className="text-caption font-medium">Cidade</Label>
                  <Input value={form.cidade} onChange={(e) => setForm((f) => ({ ...f, cidade: e.target.value }))} className="rounded-[4px] border-border bg-background text-body-sm h-9" />
                </div>
                <div className="space-y-1.5">
                  <Label className="text-caption font-medium">UF</Label>
                  <Input value={form.uf} onChange={(e) => setForm((f) => ({ ...f, uf: maskUfInput(e.target.value) }))} placeholder="PE" maxLength={2} className="font-mono rounded-[4px] border-border bg-background text-body-sm h-9" />
                </div>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-border">
              <Button type="button" variant="outline" size="sm" onClick={resetForm} className="rounded-[4px] border-border text-xs h-8">Cancelar</Button>
              <Button type="submit" size="sm" disabled={isSaving} className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs h-8">
                {isSaving ? "Salvando…" : editingId !== null ? "Salvar Alterações" : "Criar Unidade"}
              </Button>
            </div>
          </form>
        </div>
      )}

      {/* Árvore Hierárquica de Unidades */}
      <div className="rounded-[6px] border border-border bg-surface p-5 space-y-2">
        <div className="border-b border-border pb-3 mb-4">
          <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
            Hierarquia Organizacional — {unidades.length} filial{unidades.length !== 1 ? "is" : ""}
          </span>
          <p className="text-caption text-muted-foreground mt-0.5">
            Clique em uma unidade para expandir os dados jurídicos. Use os ícones para editar, ver indicadores ou inativar.
          </p>
        </div>

        {isLoading ? (
          <LoadingState label="Carregando unidades…" />
        ) : unidades.length === 0 ? (
          <div className="py-8 text-center text-caption text-muted-foreground">
            Nenhuma unidade cadastrada.
          </div>
        ) : (
          <div className="space-y-2">
            {unidades.map((u) => (
              <UnidadeRow
                key={u.id}
                unidade={u}
                onEdit={startEdit}
                onShowIndicadores={setIndicadoresUnidadeId}
                onDeactivate={(id) => deactivateMutation.mutate(id)}
                onReactivate={(id) => reactivateMutation.mutate(id)}
                isPending={deactivateMutation.isPending || reactivateMutation.isPending}
              />
            ))}
          </div>
        )}
      </div>

      {/* Dialog de Indicadores */}
      <Dialog open={indicadoresUnidadeId !== null} onOpenChange={(open) => !open && setIndicadoresUnidadeId(null)}>
        <DialogContent className="max-w-xl rounded-[6px] border border-border bg-surface p-6">
          <DialogHeader>
            <DialogTitle className="text-body-lg font-semibold text-foreground">
              Indicadores — {selectedUnidade?.nome ?? ""}
            </DialogTitle>
            <DialogDescription className="text-caption text-muted-foreground">
              Movimentação e status patrimonial desta filial.
            </DialogDescription>
          </DialogHeader>
          {indicadoresLoading ? (
            <LoadingState label="Carregando indicadores…" />
          ) : indicadores ? (
            <IndicadoresPanel dados={indicadores} />
          ) : null}
        </DialogContent>
      </Dialog>

      {/* Dialog de Confirmação de Renomeação */}
      <AlertDialog open={confirmRenameOpen} onOpenChange={(open) => !open && setConfirmRenameOpen(false)}>
        <AlertDialogContent className="rounded-[6px] border-border bg-surface">
          <AlertDialogHeader>
            <AlertDialogTitle className="text-body-lg font-semibold">Confirmar renomeação?</AlertDialogTitle>
            <AlertDialogDescription className="text-caption text-muted-foreground">
              Renomear "{originalNome}" para "{pendingData?.nome}" atualiza automaticamente todas as referências vinculadas em equipamentos e histórico.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel className="rounded-[4px] border-border text-xs" onClick={() => { setConfirmRenameOpen(false); setPendingData(null); }}>Cancelar</AlertDialogCancel>
            <AlertDialogAction
              className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs"
              onClick={() => { if (pendingData) submit(pendingData); setConfirmRenameOpen(false); setPendingData(null); }}
            >
              Confirmar Renomeação
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
