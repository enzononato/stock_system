import { Landmark, ReceiptText } from "lucide-react";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { InfoChecklist, ProcessSummary } from "./components";
import { setChecklist, useOffboardingProcesses } from "./store";

const CHECKLIST_FINANCEIRO = [
  { key: "cartao_corporativo", label: "Cancelamento / devolução do cartão corporativo" },
  { key: "adiantamentos", label: "Acerto de adiantamentos e reembolsos de despesas pendentes" },
  { key: "acessos_bancarios", label: "Bloqueio e revogação de acessos bancários e assinaturas eletrônicas" },
];

const CHECKLIST_CONTABILIDADE = [
  { key: "centro_custo", label: "Transferência ou encerramento de responsabilidade de centro de custo" },
  { key: "baixa_patrimonial", label: "Conferência de baixas de ativos sob custódia contábil" },
  { key: "provisoes_encargos", label: "Apuração e conciliação contábil das rescisões e provisões" },
];

export function OffboardingFinanceiroContabilidadePage() {
  const all = useOffboardingProcesses();
  const ativos = all.filter((p) => !p.rejected && p.currentStep < 10);

  return (
    <div className="space-y-7">
      <PageHeader
        eyebrow="Desligamento · Financeiro e Contabilidade"
        title="Checklists Administrativos"
        description="Acompanhamento e registro das etapas financeiras e contábeis do desligamento. Os checklists são informativos para controle de quitação interna."
      />

      <div className="grid gap-3 sm:grid-cols-2">
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-primary/10 text-primary">
            <Landmark className="size-4" />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Processos Ativos</p>
            <p className="mt-1 text-2xl font-bold num">{ativos.length}</p>
          </div>
        </div>
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-muted">
            <ReceiptText className="size-4" />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Escopo de Controle</p>
            <p className="mt-1 text-sm font-bold">Cartões, Adiantamentos e Ativos</p>
          </div>
        </div>
      </div>

      <Section
        title={`Checklist Financeiro e Contábil (${ativos.length})`}
        description="Marque os itens concluídos para cada processo em andamento."
      >
        {ativos.length === 0 ? (
          <p className="text-sm text-muted-foreground">Nenhum processo de desligamento em andamento.</p>
        ) : (
          <div className="space-y-6">
            {ativos.map((p) => (
              <ProcessSummary key={p.id} process={p}>
                <div className="grid gap-6 md:grid-cols-2">
                  <div className="space-y-2 rounded-lg border border-border bg-muted/10 p-4">
                    <p className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                      <Landmark className="size-3.5 text-primary" /> Financeiro
                    </p>
                    <InfoChecklist
                      items={CHECKLIST_FINANCEIRO}
                      values={p.financeiro}
                      onChange={(key, state) => setChecklist(p.id, "financeiro", key, state)}
                    />
                  </div>

                  <div className="space-y-2 rounded-lg border border-border bg-muted/10 p-4">
                    <p className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                      <ReceiptText className="size-3.5 text-primary" /> Contabilidade
                    </p>
                    <InfoChecklist
                      items={CHECKLIST_CONTABILIDADE}
                      values={p.contabilidade}
                      onChange={(key, state) => setChecklist(p.id, "contabilidade", key, state)}
                    />
                  </div>
                </div>
              </ProcessSummary>
            ))}
          </div>
        )}
      </Section>
    </div>
  );
}
