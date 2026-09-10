import { useState, type FormEvent } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import {
  FileDown,
  PackageCheck,
  ClipboardCheck,
  User,
  ArrowRight,
  ArrowLeft,
  CheckCircle2,
} from "lucide-react";

import { listItemsPaginated, type Item } from "@/api/items";
import { initiateLoan } from "@/api/loans";
import { useConstants } from "@/hooks/useConstants";
import { getErrorMessage } from "@/lib/api-error";
import { formatDate, isValidCpf, maskCpfInput } from "@/lib/utils";
import { SearchableSelect } from "@/components/app/SearchableSelect";
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from "@/components/app/ConfirmacaoTermo";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import { Badge } from "@/components/ui/badge";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Stepper } from "@/components/patterns/stepper";

const FETCH_ALL_LIMIT = 500;

const STEPS = [
  { id: "equipamento", label: "Equipamento", description: "Ativo disponível" },
  { id: "usuario", label: "Colaborador", description: "Responsável e CPF" },
  { id: "condicoes", label: "Condições", description: "Setor e centro de custo" },
  { id: "confirmacao", label: "Confirmação", description: "Emissão de termo" },
];

export function LoanPage() {
  const queryClient = useQueryClient();
  const { centerCosts, setores, revendas, isLoading: constantsLoading } = useConstants();

  const [currentStep, setCurrentStep] = useState(0);

  // Form State
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
  const [pendentesPage, setPendentesPage] = useState(0);
  const PENDENTES_PAGE_SIZE = 7;

  const { data, isLoading } = useQuery({
    queryKey: ["items"],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  });
  const items = data?.items ?? [];
  const disponivel = items.filter((i) => i.status === "Disponível");
  const pendentes = items.filter((i) => i.status === "Pendente");

  const selectedItem = disponivel.find((i) => String(i.id) === selectedItemId);

  const loanMutation = useMutation({
    mutationFn: initiateLoan,
    onSuccess: (_, vars) => {
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      setPendingItemId(vars.item_id);
      toast.success("Empréstimo iniciado com sucesso!", {
        description: "Agora emita e colete a assinatura do termo de responsabilidade.",
      });
      // Reset
      setSelectedItemId("");
      setUsuario("");
      setCpf("");
      setCenterCost("");
      setSetor("");
      setCargo("");
      setRevenda("");
      setPessoaJuridica(false);
      setCurrentStep(0);
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, "Erro ao iniciar empréstimo."));
    },
  });

  function handleNext() {
    if (currentStep === 0) {
      if (!selectedItemId) {
        toast.error("Selecione um equipamento disponível.");
        return;
      }
      // Auto-preenche revenda do item se disponível
      if (selectedItem?.revenda && !revenda) {
        setRevenda(selectedItem.revenda);
      }
      setCurrentStep(1);
    } else if (currentStep === 1) {
      if (!usuario.trim()) {
        toast.error("Informe o nome do colaborador.");
        return;
      }
      if (!isValidCpf(cpf)) {
        toast.error("CPF informado é inválido.");
        return;
      }
      setCurrentStep(2);
    } else if (currentStep === 2) {
      if (!revenda) {
        toast.error("Selecione a revenda.");
        return;
      }
      if (!setor) {
        toast.error("Selecione o setor.");
        return;
      }
      setCurrentStep(3);
    }
  }

  function handlePrev() {
    if (currentStep > 0) setCurrentStep(currentStep - 1);
  }

  function handleFinalSubmit(e: FormEvent) {
    e.preventDefault();
    if (!selectedItemId || !usuario || !isValidCpf(cpf) || !revenda || !setor) {
      toast.error("Verifique os campos obrigatórios.");
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

  return (
    <div className="page-container-reading space-y-6">
      {/* Header */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Operações de Campo
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Fluxo Guiado</span>
          </div>
          <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">
            Empréstimo de Equipamento
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Alocação de ativo com validação de colaborador e emissão de termo de responsabilidade.
          </p>
        </div>

        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-[4px] border border-border bg-surface text-caption">
            <PackageCheck className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Disponíveis:</span>
            <span className="font-mono font-semibold text-foreground">{disponivel.length}</span>
          </div>
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-[4px] border border-border bg-surface text-caption">
            <ClipboardCheck className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Pendentes:</span>
            <span className="font-mono font-semibold text-foreground">{pendentes.length}</span>
          </div>
        </div>
      </div>

      {/* Modal de confirmação pós-empréstimo */}
      {pendingItemId && (
        <ConfirmacaoTermo
          itemId={pendingItemId}
          description="Faça o upload do termo de responsabilidade assinado pelo colaborador (PDF)."
          uploadLabel="Arraste ou clique para anexar o termo assinado"
          errorMessage="Erro ao confirmar termo de empréstimo."
          onConfirmed={() => setPendingItemId(null)}
          onCancel={() => setPendingItemId(null)}
        />
      )}

      {/* Stepper Superior */}
      <div className="rounded-[6px] border border-border bg-surface p-5">
        <Stepper steps={STEPS} currentStep={currentStep} onStepClick={(step) => {
          if (step < currentStep) setCurrentStep(step);
        }} />
      </div>

      {/* Layout Grid 2 Colunas: Formulário Ativo + Contexto Fixo */}
      <div className="grid grid-cols-1 lg:grid-cols-[1fr_340px] gap-6 items-start">
        {/* Painel do Stepper (Formulário Atual) */}
        <div className="rounded-[6px] border border-border bg-surface p-6 space-y-6">
          {/* FASE 1: EQUIPAMENTO */}
          {currentStep === 0 && (
            <div className="space-y-4">
              <div>
                <h2 className="text-body-lg font-semibold text-foreground">
                  1. Seleção do Equipamento
                </h2>
                <p className="text-caption text-muted-foreground mt-0.5">
                  Escolha o ativo patrimonial disponível para alocação.
                </p>
              </div>

              <div className="space-y-2 pt-2 border-t border-border">
                <Label className="text-caption font-medium text-foreground">
                  Equipamento em Estoque *
                </Label>
                <SearchableSelect
                  options={disponivel.map((i) => ({
                    value: String(i.id),
                    label: `#${i.id} — ${i.tipo ?? ""} ${i.brand ?? ""} ${i.model ?? ""}`.trim(),
                    subtitle: [i.revenda, i.identificador].filter(Boolean).join(" • "),
                  }))}
                  value={selectedItemId}
                  onValueChange={setSelectedItemId}
                  placeholder="Buscar por ID, modelo, tipo ou serial…"
                  searchPlaceholder="Filtrar equipamentos disponíveis…"
                />
              </div>

              {selectedItem && (
                <div className="rounded-[4px] border border-border bg-surface-alt p-4 space-y-2 text-caption">
                  <div className="flex justify-between border-b border-border pb-1.5">
                    <span className="text-muted-foreground">Tipo / Modelo</span>
                    <span className="font-semibold text-foreground">
                      {selectedItem.tipo} {selectedItem.brand} {selectedItem.model}
                    </span>
                  </div>
                  <div className="flex justify-between border-b border-border pb-1.5">
                    <span className="text-muted-foreground">Serial / Identificador</span>
                    <span className="font-mono text-foreground">{selectedItem.identificador || "—"}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted-foreground">Unidade de Origem</span>
                    <span className="text-foreground">{selectedItem.revenda || "—"}</span>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* FASE 2: COLABORADOR */}
          {currentStep === 1 && (
            <div className="space-y-4">
              <div>
                <h2 className="text-body-lg font-semibold text-foreground">
                  2. Dados do Colaborador Responsável
                </h2>
                <p className="text-caption text-muted-foreground mt-0.5">
                  Informe quem receberá a custódia do equipamento.
                </p>
              </div>

              <div className="space-y-4 pt-2 border-t border-border">
                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">
                    Nome Completo do Colaborador *
                  </Label>
                  <Input
                    value={usuario}
                    onChange={(e) => setUsuario(e.target.value)}
                    placeholder="Ex: João da Silva"
                    required
                    className="rounded-[4px] border-border bg-background text-body-sm h-9"
                  />
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <Label className="text-caption font-medium text-foreground">
                      CPF do Colaborador *
                    </Label>
                    <Input
                      value={cpf}
                      onChange={(e) => setCpf(maskCpfInput(e.target.value))}
                      placeholder="000.000.000-00"
                      maxLength={14}
                      required
                      className="font-mono rounded-[4px] border-border bg-background text-body-sm h-9"
                    />
                  </div>

                  <div className="flex items-center gap-2 pt-6">
                    <Checkbox
                      id="pj"
                      checked={pessoaJuridica}
                      onCheckedChange={(c) => setPessoaJuridica(Boolean(c))}
                      className="rounded-[2px] border-border"
                    />
                    <label htmlFor="pj" className="text-caption text-foreground cursor-pointer select-none">
                      Colaborador Terceirizado / Pessoa Jurídica
                    </label>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* FASE 3: CONDIÇÕES */}
          {currentStep === 2 && (
            <div className="space-y-4">
              <div>
                <h2 className="text-body-lg font-semibold text-foreground">
                  3. Condições Operacionais e Alocação
                </h2>
                <p className="text-caption text-muted-foreground mt-0.5">
                  Unidade de lotação, centro de custo e cargo do colaborador.
                </p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-2 border-t border-border">
                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">Revenda / Filial *</Label>
                  <Select value={revenda} onValueChange={setRevenda} required disabled={constantsLoading}>
                    <SelectTrigger className="rounded-[4px] border-border bg-background text-body-sm h-9">
                      <SelectValue placeholder="Selecione a filial" />
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
                  <Label className="text-caption font-medium text-foreground">Setor *</Label>
                  <Select value={setor} onValueChange={setSetor} required disabled={constantsLoading}>
                    <SelectTrigger className="rounded-[4px] border-border bg-background text-body-sm h-9">
                      <SelectValue placeholder="Selecione o setor" />
                    </SelectTrigger>
                    <SelectContent className="rounded-[4px] border-border bg-surface">
                      {setores.map((s) => (
                        <SelectItem key={s} value={s} className="text-body-sm rounded-[2px]">
                          {s}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">Cargo / Função</Label>
                  <Input
                    value={cargo}
                    onChange={(e) => setCargo(e.target.value)}
                    placeholder="Ex: Analista Comercial"
                    className="rounded-[4px] border-border bg-background text-body-sm h-9"
                  />
                </div>

                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">Centro de Custo</Label>
                  <Select value={centerCost} onValueChange={setCenterCost} disabled={constantsLoading}>
                    <SelectTrigger className="rounded-[4px] border-border bg-background text-body-sm h-9">
                      <SelectValue placeholder="Selecione o centro de custo" />
                    </SelectTrigger>
                    <SelectContent className="rounded-[4px] border-border bg-surface">
                      {centerCosts.map((c) => (
                        <SelectItem key={c} value={c} className="text-body-sm rounded-[2px]">
                          {c}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">Data da Entrega</Label>
                  <Input
                    value={dateIssue}
                    onChange={(e) => setDateIssue(e.target.value)}
                    placeholder="DD/MM/AAAA"
                    className="font-mono rounded-[4px] border-border bg-background text-body-sm h-9"
                  />
                </div>
              </div>
            </div>
          )}

          {/* FASE 4: CONFIRMAÇÃO & REVISÃO */}
          {currentStep === 3 && (
            <div className="space-y-4">
              <div>
                <h2 className="text-body-lg font-semibold text-foreground">
                  4. Revisão e Emissão do Termo
                </h2>
                <p className="text-caption text-muted-foreground mt-0.5">
                  Confira as informações antes de formalizar a entrega patrimonial.
                </p>
              </div>

              <div className="rounded-[4px] border border-border bg-surface-alt p-4 space-y-3 text-caption">
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">Equipamento</span>
                  <span className="font-semibold text-foreground">
                    #{selectedItem?.id} — {selectedItem?.tipo} {selectedItem?.brand} {selectedItem?.model}
                  </span>
                </div>
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">Colaborador</span>
                  <span className="font-semibold text-foreground">{usuario}</span>
                </div>
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">CPF</span>
                  <span className="font-mono text-foreground">{cpf}</span>
                </div>
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">Revenda / Setor</span>
                  <span className="text-foreground">{revenda} • {setor}</span>
                </div>
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">Data Empréstimo</span>
                  <span className="font-mono text-foreground">{dateIssue}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-muted-foreground">Status Pós-Registro</span>
                  <Badge variant="outline" className="rounded-[2px] font-mono text-[10px] border-border text-foreground">
                    ! PENDENTE TERMO
                  </Badge>
                </div>
              </div>
            </div>
          )}

          {/* Controles de Navegação do Stepper */}
          <div className="pt-4 border-t border-border flex items-center justify-between">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={handlePrev}
              disabled={currentStep === 0 || loanMutation.isPending}
              className="rounded-[4px] border-border"
            >
              <ArrowLeft className="mr-1.5 size-3.5" />
              Anterior
            </Button>

            {currentStep < 3 ? (
              <Button
                type="button"
                size="sm"
                onClick={handleNext}
                className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 font-medium"
              >
                Próximo
                <ArrowRight className="ml-1.5 size-3.5" />
              </Button>
            ) : (
              <Button
                type="button"
                size="sm"
                onClick={handleFinalSubmit}
                disabled={loanMutation.isPending}
                className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 font-medium"
              >
                <CheckCircle2 className="mr-1.5 size-3.5" />
                {loanMutation.isPending ? "Processando…" : "Confirmar Empréstimo"}
              </Button>
            )}
          </div>
        </div>

        {/* Coluna Lateral de Contexto Persistente */}
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-4">
          <div className="border-b border-border pb-2.5">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Contexto do Ativo
            </span>
            <p className="text-body-sm font-semibold text-foreground mt-0.5">
              {selectedItem ? `#${selectedItem.id} — ${selectedItem.tipo}` : "Nenhum ativo selecionado"}
            </p>
          </div>

          {selectedItem ? (
            <div className="space-y-2.5 text-caption">
              <div>
                <span className="text-muted-foreground block">Marca / Modelo</span>
                <span className="font-medium text-foreground">{selectedItem.brand} {selectedItem.model}</span>
              </div>
              <div>
                <span className="text-muted-foreground block">Serial</span>
                <span className="font-mono text-foreground">{selectedItem.identificador || "—"}</span>
              </div>
              <div>
                <span className="text-muted-foreground block">Filial Atual</span>
                <span className="text-foreground">{selectedItem.revenda || "—"}</span>
              </div>
              <div>
                <span className="text-muted-foreground block">Nota Fiscal</span>
                <span className="font-mono text-foreground">{selectedItem.nota_fiscal || "—"}</span>
              </div>
            </div>
          ) : (
            <p className="text-caption text-muted-foreground italic">
              Selecione um equipamento na etapa 1 para manter seus dados visíveis durante todo o processo.
            </p>
          )}

          <div className="pt-3 border-t border-border">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block mb-2">
              Regras do Fluxo
            </span>
            <ul className="space-y-1.5 text-caption text-muted-foreground leading-relaxed">
              <li>• O status passará para <strong>Pendente</strong> até que o termo assinado seja anexado.</li>
              <li>• O CPF é validado pelos dígitos verificadores oficiais.</li>
              <li>• O termo de responsabilidade pode ser impresso ou baixado em PDF imediatamente após a confirmação.</li>
            </ul>
          </div>
        </div>
      </div>

      {/* Empréstimos Pendentes Recentes */}
      {pendentes.length > 0 && (
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-3">
          <div className="flex items-center justify-between border-b border-border pb-2">
            <div>
              <h3 className="text-body font-semibold text-foreground">
                Empréstimos Pendentes de Termo ({pendentes.length})
              </h3>
              <p className="text-caption text-muted-foreground">
                Equipamentos entregues que ainda aguardam a anexação do PDF assinado.
              </p>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-body-sm">
              <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground">
                <tr>
                  <th className="py-2 px-3">ID</th>
                  <th className="py-2 px-3">Tipo</th>
                  <th className="py-2 px-3">Colaborador</th>
                  <th className="py-2 px-3">Revenda</th>
                  <th className="py-2 px-3 text-right">Ação</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {pendentes
                  .slice(pendentesPage * PENDENTES_PAGE_SIZE, (pendentesPage + 1) * PENDENTES_PAGE_SIZE)
                  .map((item) => (
                    <tr key={item.id} className="hover:bg-muted/30">
                      <td className="py-2 px-3 font-mono font-medium">#{item.id}</td>
                      <td className="py-2 px-3">{item.tipo} {item.brand}</td>
                      <td className="py-2 px-3 font-medium">{item.assigned_to}</td>
                      <td className="py-2 px-3 text-muted-foreground">{item.revenda || "—"}</td>
                      <td className="py-2 px-3 text-right">
                        <div className="flex justify-end gap-2">
                          <Button
                            size="sm"
                            variant="outline"
                            onClick={() => generateAndDownloadLoanTerm(item.id)}
                            className="rounded-[4px] border-border text-xs h-7 px-2.5"
                          >
                            <FileDown className="mr-1 size-3" />
                            Termo
                          </Button>
                          <Button
                            size="sm"
                            onClick={() => setPendingItemId(item.id)}
                            className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs h-7 px-2.5"
                          >
                            Confirmar
                          </Button>
                        </div>
                      </td>
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>

          {Math.ceil(pendentes.length / PENDENTES_PAGE_SIZE) > 1 && (
            <div className="flex items-center justify-between border-t border-border pt-3">
              <span className="text-caption text-muted-foreground">
                Página {pendentesPage + 1} de {Math.ceil(pendentes.length / PENDENTES_PAGE_SIZE)} ({pendentes.length} pendentes)
              </span>
              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pendentesPage === 0}
                  onClick={() => setPendentesPage((p) => Math.max(0, p - 1))}
                  className="rounded-[4px] border-border text-xs h-7 px-2.5"
                >
                  ← Anterior
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pendentesPage >= Math.ceil(pendentes.length / PENDENTES_PAGE_SIZE) - 1}
                  onClick={() => setPendentesPage((p) => p + 1)}
                  className="rounded-[4px] border-border text-xs h-7 px-2.5"
                >
                  Próxima →
                </Button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
