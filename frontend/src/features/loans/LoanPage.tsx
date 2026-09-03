import { useState, type FormEvent } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { FileDown, PackageCheck, UserRound, ClipboardCheck } from "lucide-react";

import { listItemsPaginated, type Item } from "@/api/items";
import { initiateLoan } from "@/api/loans";
import { useConstants } from "@/hooks/useConstants";
import { getErrorMessage } from "@/lib/api-error";
import { formatDate, isValidCpf, maskCpfInput } from "@/lib/utils";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { DataTable, type Column } from "@/components/app/DataTable";
import { SearchableSelect } from "@/components/app/SearchableSelect";
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from "@/components/app/ConfirmacaoTermo";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

const FETCH_ALL_LIMIT = 500;

export function LoanPage() {
  const queryClient = useQueryClient();
  const { centerCosts, setores, revendas, isLoading: constantsLoading } = useConstants();
  const [selectedItemId, setSelectedItemId] = useState("");
  const [usuario, setUsuario] = useState("");
  const [cpf, setCpf] = useState("");
  const [centerCost, setCenterCost] = useState("");
  const [setor, setSetor] = useState("");
  const [cargo, setCargo] = useState("");
  const [revenda, setRevenda] = useState("");
  const [pessoaJuridica, setPessoaJuridica] = useState(false);
  const [dateIssue, setDateIssue] = useState(() => {
    const d = new Date();
    return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
  });
  const [pendingItemId, setPendingItemId] = useState<number | null>(null);

  const { data, isLoading, error, refetch } = useQuery({
    queryKey: ["items"],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  });
  const items = data?.items ?? [];
  const disponivel = items.filter((i) => i.status === "Disponível");
  const pendentes = items.filter((i) => i.status === "Pendente");

  const loanMutation = useMutation({
    mutationFn: initiateLoan,
    onSuccess: (_, vars) => {
      queryClient.invalidateQueries({ queryKey: ["items"] });
      setPendingItemId(vars.item_id);
      toast.success("Empréstimo iniciado", {
        description: "Agora gere e confirme o termo de responsabilidade.",
      });
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao iniciar empréstimo.")),
  });

  function handleLoanSubmit(e: FormEvent) {
    e.preventDefault();
    if (!selectedItemId) {
      toast.error("Selecione um equipamento.");
      return;
    }
    if (!isValidCpf(cpf)) {
      toast.error("CPF inválido.");
      return;
    }
    loanMutation.mutate({
      item_id: Number(selectedItemId),
      usuario,
      cpf,
      center_cost: centerCost,
      cargo,
      setor,
      revenda,
      date_issue: dateIssue,
      pessoa_juridica: pessoaJuridica,
    });
  }

  const selectedItem = disponivel.find((i) => String(i.id) === selectedItemId);
  const pendingColumns: Column<Item>[] = [
    { key: "id", header: "ID", cell: (i) => `#${i.id}`, primary: true },
    { key: "tipo", header: "Tipo", cell: (i) => i.tipo ?? "-" },
    { key: "brand", header: "Marca", cell: (i) => i.brand ?? "-" },
    { key: "assigned_to", header: "Usuário", cell: (i) => i.assigned_to ?? "-" },
    { key: "revenda", header: "Revenda", cell: (i) => i.revenda ?? "-", hideBelow: "md" },
    { key: "date_issued", header: "Data", cell: (i) => formatDate(i.date_issued), hideBelow: "lg" },
    {
      key: "actions",
      header: "Ações",
      hideOnMobile: true,
      cell: (i) => (
        <Button size="sm" variant="outline" onClick={() => generateAndDownloadLoanTerm(i.id)}>
          <FileDown className="mr-1.5 size-3.5" aria-hidden />
          Termo
        </Button>
      ),
    },
  ];

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Operação"
        title="Empréstimo de equipamento"
        description="Registre a entrega de um patrimônio e conclua o termo de responsabilidade."
      />

      <div className="grid gap-4 md:grid-cols-3">
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-primary/10 text-primary">
            <PackageCheck className="size-4" aria-hidden />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Disponíveis
            </p>
            <p className="mt-0.5 text-xl font-bold num">{disponivel.length}</p>
          </div>
        </div>
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-muted text-foreground">
            <ClipboardCheck className="size-4" aria-hidden />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Pendentes
            </p>
            <p className="mt-0.5 text-xl font-bold num">{pendentes.length}</p>
          </div>
        </div>
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-muted text-foreground">
            <UserRound className="size-4" aria-hidden />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Etapa atual
            </p>
            <p className="mt-0.5 text-sm font-bold">Dados do responsável</p>
          </div>
        </div>
      </div>

      <Section
        title="Dados do empréstimo"
        description="Selecione o equipamento e informe o responsável pela retirada."
      >
        <form onSubmit={handleLoanSubmit} className="space-y-6">
          <div className="rounded-lg border border-border bg-muted/20 p-4">
            <Label>Equipamento *</Label>
            <div className="mt-1.5">
              <SearchableSelect
                options={disponivel.map((i) => ({
                  value: String(i.id),
                  label: `#${i.id} — ${i.tipo ?? ""} ${i.brand ?? ""} ${i.model ?? ""}`,
                  subtitle: [i.revenda, i.identificador].filter(Boolean).join(" • "),
                }))}
                value={selectedItemId}
                onValueChange={setSelectedItemId}
                placeholder="Selecione ou busque um equipamento disponível…"
                searchPlaceholder="Buscar por ID, tipo, marca, modelo, patrimônio…"
              />
            </div>
            {selectedItem && (
              <div className="mt-3 flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground">
                <span>
                  <strong className="text-foreground">Tipo:</strong> {selectedItem.tipo || "-"}
                </span>
                <span>
                  <strong className="text-foreground">Modelo:</strong>{" "}
                  {[selectedItem.brand, selectedItem.model].filter(Boolean).join(" ") || "-"}
                </span>
                <span>
                  <strong className="text-foreground">Unidade:</strong>{" "}
                  {selectedItem.revenda || "-"}
                </span>
              </div>
            )}
          </div>

          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
            <div className="flex flex-col gap-1.5">
              <Label>Funcionário *</Label>
              <Input
                value={usuario}
                onChange={(e) => setUsuario(e.target.value)}
                placeholder="Nome completo"
                required
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>CPF *</Label>
              <Input
                value={cpf}
                onChange={(e) => setCpf(maskCpfInput(e.target.value))}
                placeholder="000.000.000-00"
                maxLength={14}
                required
              />
              <label className="mt-1 flex select-none items-center gap-2 text-xs text-muted-foreground">
                <Checkbox
                  checked={pessoaJuridica}
                  onCheckedChange={(v) => setPessoaJuridica(v === true)}
                />
                É pessoa jurídica
              </label>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Cargo *</Label>
              <Input
                value={cargo}
                onChange={(e) => setCargo(e.target.value)}
                placeholder="Cargo"
                required
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Centro de custo *</Label>
              <Select
                value={centerCost}
                onValueChange={setCenterCost}
                required
                disabled={constantsLoading}
              >
                <SelectTrigger>
                  <SelectValue placeholder="Selecione" />
                </SelectTrigger>
                <SelectContent>
                  {centerCosts.map((c) => (
                    <SelectItem key={c} value={c}>
                      {c}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Setor *</Label>
              <Select value={setor} onValueChange={setSetor} required disabled={constantsLoading}>
                <SelectTrigger>
                  <SelectValue placeholder="Selecione" />
                </SelectTrigger>
                <SelectContent>
                  {setores.map((s) => (
                    <SelectItem key={s} value={s}>
                      {s}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Revenda *</Label>
              <Select
                value={revenda}
                onValueChange={setRevenda}
                required
                disabled={constantsLoading}
              >
                <SelectTrigger>
                  <SelectValue placeholder="Selecione" />
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
            <div className="flex flex-col gap-1.5 sm:max-w-[calc(50%-0.5rem)] lg:max-w-none">
              <Label>Data do empréstimo *</Label>
              <Input
                value={dateIssue}
                onChange={(e) => setDateIssue(e.target.value)}
                placeholder="dd/mm/aaaa"
                required
              />
            </div>
          </div>

          <div className="flex flex-col-reverse gap-3 border-t border-border pt-5 sm:flex-row sm:items-center sm:justify-end">
            <p className="mr-auto text-xs text-muted-foreground">
              Após iniciar, o sistema solicitará a confirmação do termo.
            </p>
            <Button type="submit" disabled={loanMutation.isPending || isLoading}>
              {loanMutation.isPending ? "Iniciando…" : "Iniciar empréstimo"}
            </Button>
          </div>
        </form>
      </Section>

      {pendingItemId && (
        <ConfirmacaoTermo
          itemId={pendingItemId}
          description={
            <>
              <span>1. Gere o termo de responsabilidade.</span>
              <br />
              <span>2. Imprima e assine o documento.</span>
              <br />
              <span>3. Envie o PDF assinado para confirmar.</span>
            </>
          }
          showGenerateButton
          onConfirmed={() => setPendingItemId(null)}
        />
      )}

      <Section
        title="Empréstimos pendentes de confirmação"
        description="Processos iniciados que ainda aguardam a confirmação do termo."
      >
        <DataTable
          data={pendentes}
          columns={pendingColumns}
          rowKey={(i) => i.id}
          isLoading={isLoading}
          error={error}
          onRetry={() => void refetch()}
          emptyTitle="Nenhum empréstimo pendente"
          emptyDescription="Os empréstimos iniciados aparecem aqui até a confirmação."
        />
      </Section>
    </div>
  );
}
