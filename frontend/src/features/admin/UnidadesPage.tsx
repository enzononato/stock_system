import { useState, type FormEvent } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { Ban, BarChart3, Building2, Pencil, RotateCcw } from "lucide-react";

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
import { DataTable, type Column } from "@/components/app/DataTable";
import { KpiCard } from "@/components/app/KpiCard";
import { PageHeader, Section } from "@/components/app/PageHeader";
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

// --- Validação e máscaras -----------------------------------------------------

/**
 * Valida um CNPJ conferindo os dois dígitos verificadores (módulo 11).
 * Rejeita sequências com os 14 dígitos repetidos.
 */
export function isValidCnpj(cnpj: string | null | undefined): boolean {
  if (!cnpj) return false;
  const digits = cnpj.replace(/\D/g, "");
  if (digits.length !== 14) return false;
  if (/^(\d)\1{13}$/.test(digits)) return false;

  const calcCheckDigit = (base: string, weights: number[]): number => {
    let total = 0;
    for (let i = 0; i < base.length; i++) {
      total += Number(base[i]) * weights[i]!;
    }
    const remainder = total % 11;
    return remainder < 2 ? 0 : 11 - remainder;
  };

  const weights1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
  const weights2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];

  const dv1 = calcCheckDigit(digits.slice(0, 12), weights1);
  const dv2 = calcCheckDigit(digits.slice(0, 12) + String(dv1), weights2);

  return dv1 === Number(digits[12]) && dv2 === Number(digits[13]);
}

/** Máscara progressiva de CNPJ (00.000.000/0000-00). */
export function maskCnpjInput(value: string): string {
  return value
    .replace(/\D/g, "")
    .slice(0, 14)
    .replace(/(\d{2})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1/$2")
    .replace(/(\d{4})(\d{1,2})$/, "$1-$2");
}

/** CEP no formato 00000-000 — validação usada só quando o campo está preenchido. */
export function isValidCep(cep: string | null | undefined): boolean {
  if (!cep) return false;
  return /^\d{5}-\d{3}$/.test(cep);
}

/** Máscara progressiva de CEP (00000-000). */
export function maskCepInput(value: string): string {
  return value
    .replace(/\D/g, "")
    .slice(0, 8)
    .replace(/(\d{5})(\d{1,3})$/, "$1-$2");
}

/** UF: exatamente 2 letras. */
export function isValidUf(uf: string | null | undefined): boolean {
  if (!uf) return false;
  return /^[A-Z]{2}$/.test(uf);
}

/** Restringe a digitação da UF a no máximo 2 letras, sempre maiúsculas. */
export function maskUfInput(value: string): string {
  return value
    .replace(/[^A-Za-z]/g, "")
    .slice(0, 2)
    .toUpperCase();
}

// --- Indicadores ---------------------------------------------------------------

const STATUS_ORDER = ["Disponível", "Indisponível", "Pendente", "Pendente Devolução"];

function Stat({ label, value }: { label: string; value: number }) {
  return (
    <div className="rounded-lg border border-border bg-muted/40 p-3">
      <p className="text-xs text-muted-foreground">{label}</p>
      <p className="text-lg font-semibold text-foreground">{value}</p>
    </div>
  );
}

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
        <Stat label="Termos Emitidos" value={dados.termos_emitidos} />
        <Stat label="Termos Confirmados" value={dados.termos_confirmados} />
        <Stat label="Devoluções Concluídas" value={dados.devolucoes_concluidas} />
        <Stat label="Total de Itens" value={dados.itens_total} />
        <Stat label="Empréstimos Ativos" value={dados.emprestimos_ativos} />
        <Stat label="Periféricos Vinculados" value={dados.perifericos_vinculados} />
      </div>

      <div>
        <p className="mb-2 text-sm font-medium text-muted-foreground">Itens por Status</p>
        <div className="flex flex-wrap gap-2">
          {STATUS_ORDER.map((status) => (
            <Badge key={status} variant="outline">
              {status}: {dados.itens_por_status[status] ?? 0}
            </Badge>
          ))}
        </div>
      </div>

      {semMovimento && (
        <p className="text-sm italic text-muted-foreground">
          Esta unidade ainda não tem nenhuma movimentação registrada.
        </p>
      )}
    </div>
  );
}

// --- Página ---------------------------------------------------------------------

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

  const [confirmRenameOpen, setConfirmRenameOpen] = useState(false);
  const [pendingData, setPendingData] = useState<UnidadeInput | null>(null);

  const [indicadoresUnidadeId, setIndicadoresUnidadeId] = useState<number | null>(null);

  const [showInactive, setShowInactive] = useState(false);

  const {
    data: unidades = [],
    isLoading,
    error,
    refetch,
  } = useQuery({
    queryKey: ["unidades", { includeInactive: showInactive }],
    queryFn: () => listUnidades(showInactive ? { include_inactive: true } : undefined),
  });

  const { data: indicadores, isLoading: indicadoresLoading } = useQuery({
    queryKey: ["unidades", indicadoresUnidadeId, "indicadores"],
    queryFn: () => getIndicadoresUnidade(indicadoresUnidadeId as number),
    enabled: indicadoresUnidadeId !== null,
  });

  function invalidateAfterWrite() {
    queryClient.invalidateQueries({ queryKey: ["unidades"] });
    queryClient.invalidateQueries({ queryKey: ["constants"] });
  }

  function resetForm() {
    setEditingId(null);
    setOriginalNome("");
    setForm(emptyForm);
  }

  const createMutation = useMutation({
    mutationFn: (data: UnidadeInput) => createUnidade(data),
    onSuccess: () => {
      invalidateAfterWrite();
      resetForm();
      toast.success("Unidade criada com sucesso!");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao criar unidade.")),
  });

  const updateMutation = useMutation({
    mutationFn: ({ id, data }: { id: number; data: UnidadeInput }) => updateUnidade(id, data),
    onSuccess: () => {
      invalidateAfterWrite();
      resetForm();
      toast.success("Unidade atualizada com sucesso!");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao atualizar unidade.")),
  });

  const deactivateMutation = useMutation({
    mutationFn: (id: number) => deactivateUnidade(id),
    onSuccess: () => {
      invalidateAfterWrite();
      toast.success("Unidade inativada.");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao inativar unidade.")),
  });

  const reactivateMutation = useMutation({
    mutationFn: (id: number) => reactivateUnidade(id),
    onSuccess: () => {
      invalidateAfterWrite();
      toast.success("Unidade reativada.");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao reativar unidade.")),
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
  }

  function buildAndValidate(): UnidadeInput | null {
    if (!form.nome.trim()) {
      toast.error("Informe o nome da unidade.");
      return null;
    }
    if (!form.razaoSocial.trim()) {
      toast.error("Informe a razão social.");
      return null;
    }
    if (!isValidCnpj(form.cnpj)) {
      toast.error("CNPJ inválido. Confira os dígitos digitados.");
      return null;
    }
    if (form.cep.trim() && !isValidCep(form.cep)) {
      toast.error("CEP deve estar no formato 00000-000.");
      return null;
    }
    if (form.uf.trim() && !isValidUf(form.uf)) {
      toast.error("UF deve ter exatamente 2 letras.");
      return null;
    }
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
    if (editingId !== null) {
      updateMutation.mutate({ id: editingId, data });
    } else {
      createMutation.mutate(data);
    }
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

  function confirmRename() {
    if (pendingData) submit(pendingData);
    setConfirmRenameOpen(false);
    setPendingData(null);
  }

  function cancelRename() {
    setConfirmRenameOpen(false);
    setPendingData(null);
  }

  const isSaving = createMutation.isPending || updateMutation.isPending;

  const columns: Column<Unidade>[] = [
    {
      key: "nome",
      header: "Nome",
      primary: true,
      cell: (u) => (
        <div className="flex items-center gap-2">
          <span className={u.is_active ? "" : "text-muted-foreground line-through"}>{u.nome}</span>
          {!u.is_active && <Badge variant="destructive">Inativa</Badge>}
        </div>
      ),
    },
    { key: "razao_social", header: "Razão Social", cell: (u) => u.razao_social, hideBelow: "md" },
    { key: "cnpj", header: "CNPJ", cell: (u) => u.cnpj, hideBelow: "lg" },
    {
      key: "localizacao",
      header: "Cidade/UF",
      hideBelow: "lg",
      cell: (u) => {
        if (!u.cidade && !u.uf) return <span className="text-muted-foreground">—</span>;
        return (
          <span>
            {u.cidade || "—"}/{u.uf || "—"}
          </span>
        );
      },
    },
    {
      key: "actions",
      header: "",
      cell: (u) => (
        <div className="flex gap-2">
          <Button
            size="sm"
            variant="ghost"
            aria-label={`Indicadores da unidade ${u.nome}`}
            onClick={() => setIndicadoresUnidadeId(u.id)}
          >
            <BarChart3 className="size-3.5" aria-hidden />
          </Button>
          <Button
            size="sm"
            variant="ghost"
            aria-label={`Editar unidade ${u.nome}`}
            onClick={() => startEdit(u)}
          >
            <Pencil className="size-3.5" aria-hidden />
          </Button>
          {u.is_active && (
            <AlertDialog>
              <AlertDialogTrigger asChild>
                <Button
                  size="sm"
                  variant="ghost"
                  aria-label={`Inativar unidade ${u.nome}`}
                  className="text-destructive hover:text-destructive/80"
                  disabled={deactivateMutation.isPending}
                >
                  <Ban className="size-3.5" aria-hidden />
                </Button>
              </AlertDialogTrigger>
              <AlertDialogContent>
                <AlertDialogHeader>
                  <AlertDialogTitle>Inativar unidade &quot;{u.nome}&quot;?</AlertDialogTitle>
                  <AlertDialogDescription>
                    Unidades com itens ativos vinculados não podem ser inativadas. Esta ação não
                    pode ser desfeita.
                  </AlertDialogDescription>
                </AlertDialogHeader>
                <AlertDialogFooter>
                  <AlertDialogCancel>Cancelar</AlertDialogCancel>
                  <AlertDialogAction onClick={() => deactivateMutation.mutate(u.id)}>
                    Inativar
                  </AlertDialogAction>
                </AlertDialogFooter>
              </AlertDialogContent>
            </AlertDialog>
          )}
          {!u.is_active && (
            <Button
              size="sm"
              variant="ghost"
              aria-label={`Reativar unidade ${u.nome}`}
              className="text-emerald-600 hover:text-emerald-700"
              disabled={reactivateMutation.isPending}
              onClick={() => reactivateMutation.mutate(u.id)}
            >
              <RotateCcw className="size-3.5" aria-hidden />
            </Button>
          )}
        </div>
      ),
    },
  ];

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Administração"
        title="Unidades de Revenda"
        description='Dados jurídicos e fiscais das unidades (antigas "revendas") e indicadores por unidade'
      />

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
        <KpiCard label="Unidades Cadastradas" value={unidades.length} />
        <KpiCard label="Ativas" value={unidades.filter((u) => u.is_active !== false).length} />
        <KpiCard label="Inativas" value={unidades.filter((u) => u.is_active === false).length} />
      </div>

      <Section title={editingId !== null ? `Editar Unidade #${editingId}` : "Nova Unidade"}>
        <form onSubmit={handleSubmit} className="space-y-4">
          <div className="flex items-center gap-2 text-sm font-medium text-foreground">
            <Building2 className="size-4" aria-hidden />
            {editingId !== null ? `Editar Unidade #${editingId}` : "Nova Unidade"}
          </div>

          {editingId !== null && (
            <p className="rounded-md border border-amber-500/30 bg-amber-500/10 px-3 py-2 text-xs text-amber-700 dark:text-amber-400">
              O nome da unidade liga os equipamentos e todo o histórico a ela. Alterá-lo atualiza
              automaticamente essas referências — confirme antes de salvar se for esse o caso.
            </p>
          )}

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="unidade-nome">Nome *</Label>
              <Input
                id="unidade-nome"
                value={form.nome}
                onChange={(e) => setForm((f) => ({ ...f, nome: e.target.value }))}
                required
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="unidade-razao-social">Razão Social *</Label>
              <Input
                id="unidade-razao-social"
                value={form.razaoSocial}
                onChange={(e) => setForm((f) => ({ ...f, razaoSocial: e.target.value }))}
                required
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="unidade-cnpj">CNPJ *</Label>
              <Input
                id="unidade-cnpj"
                value={form.cnpj}
                onChange={(e) => setForm((f) => ({ ...f, cnpj: maskCnpjInput(e.target.value) }))}
                placeholder="00.000.000/0000-00"
                maxLength={18}
                required
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="unidade-endereco">Endereço</Label>
              <Input
                id="unidade-endereco"
                value={form.endereco}
                onChange={(e) => setForm((f) => ({ ...f, endereco: e.target.value }))}
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <div className="flex items-center gap-2">
                <Label htmlFor="unidade-cep">CEP</Label>
                {editingId !== null && !form.cep && (
                  <Badge variant="outline" className="text-[10px]">
                    Não informado
                  </Badge>
                )}
              </div>
              <Input
                id="unidade-cep"
                value={form.cep}
                onChange={(e) => setForm((f) => ({ ...f, cep: maskCepInput(e.target.value) }))}
                placeholder="00000-000 (opcional)"
                maxLength={9}
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="unidade-cidade">Cidade</Label>
              <Input
                id="unidade-cidade"
                value={form.cidade}
                onChange={(e) => setForm((f) => ({ ...f, cidade: e.target.value }))}
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor="unidade-uf">UF</Label>
              <Input
                id="unidade-uf"
                value={form.uf}
                onChange={(e) => setForm((f) => ({ ...f, uf: maskUfInput(e.target.value) }))}
                placeholder="Ex.: PE"
                maxLength={2}
              />
            </div>
          </div>

          <div className="flex justify-end gap-2">
            {editingId !== null && (
              <Button type="button" variant="ghost" onClick={resetForm}>
                Cancelar edição
              </Button>
            )}
            <Button type="submit" disabled={isSaving}>
              {isSaving
                ? "Salvando..."
                : editingId !== null
                  ? "Salvar Alterações"
                  : "Criar Unidade"}
            </Button>
          </div>
        </form>
      </Section>

      <Section
        title="Unidades"
        actions={
          <label className="flex items-center gap-2 text-sm text-muted-foreground">
            <Switch checked={showInactive} onCheckedChange={setShowInactive} id="show-inactive" />
            Mostrar inativas
          </label>
        }
      >
        <DataTable
          data={unidades}
          columns={columns}
          rowKey={(u) => u.id}
          isLoading={isLoading}
          error={error}
          onRetry={() => void refetch()}
          clientPageSize={7}
          emptyTitle="Nenhuma unidade cadastrada"
        />
      </Section>

      <AlertDialog open={confirmRenameOpen} onOpenChange={(open) => !open && cancelRename()}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Confirmar renomeação da unidade?</AlertDialogTitle>
            <AlertDialogDescription>
              Renomear &quot;{originalNome}&quot; para &quot;{pendingData?.nome}&quot; atualiza
              automaticamente todas as referências a esta unidade em equipamentos e histórico.
              Deseja continuar?
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel onClick={cancelRename}>Cancelar</AlertDialogCancel>
            <AlertDialogAction onClick={confirmRename}>Confirmar</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      <Dialog
        open={indicadoresUnidadeId !== null}
        onOpenChange={(open) => !open && setIndicadoresUnidadeId(null)}
      >
        <DialogContent className="max-w-xl">
          <DialogHeader>
            <DialogTitle>
              Indicadores — {unidades.find((u) => u.id === indicadoresUnidadeId)?.nome ?? ""}
            </DialogTitle>
            <DialogDescription>Movimentação e status de itens desta unidade.</DialogDescription>
          </DialogHeader>
          {indicadoresLoading ? (
            <LoadingState label="Carregando indicadores…" />
          ) : indicadores ? (
            <IndicadoresPanel dados={indicadores} />
          ) : null}
        </DialogContent>
      </Dialog>
    </div>
  );
}
