import { ClipboardCheck, Clock3, LockKeyhole, ShieldCheck } from "lucide-react";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { EmptyState } from "@/components/app/StateBlocks";
import { useAuth } from "@/lib/auth";
import { toast } from "sonner";
import { ProcessSummary, StepAction } from "./components";
import { completeStep, useOffboardingProcesses } from "./store";

export function OffboardingTiPage() {
  const { user } = useAuth();
  const all = useOffboardingProcesses();
  const queue = all.filter((p) => !p.rejected && (p.currentStep === 4 || p.currentStep === 6));
  const backup = queue.filter((p) => p.currentStep === 4).length;
  const bloqueio = queue.filter((p) => p.currentStep === 6).length;
  const aguardando = all.filter((p) => !p.rejected && p.currentStep < 4);

  return (
    <div className="space-y-7">
      <PageHeader
        eyebrow="Desligamento · TI"
        title="Solicitações de TI"
        description="Acompanhe os chamados de backup e bloqueio de acessos relacionados aos desligamentos. O sistema registra a conclusão informada pelo TI; não executa o backup automaticamente."
      />

      <div className="grid gap-3 sm:grid-cols-3">
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-primary/10 text-primary">
            <ClipboardCheck className="size-4" />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Na fila
            </p>
            <p className="mt-1 text-2xl font-bold num">{queue.length}</p>
          </div>
        </div>
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-muted">
            <ShieldCheck className="size-4" />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Backup
            </p>
            <p className="mt-1 text-2xl font-bold num">{backup}</p>
          </div>
        </div>
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-muted">
            <LockKeyhole className="size-4" />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Bloqueios
            </p>
            <p className="mt-1 text-2xl font-bold num">{bloqueio}</p>
          </div>
        </div>
      </div>

      <Section
        title={`Chamados na fila (${queue.length})`}
        description="Cada solicitação precisa ser atendida pelo TI antes de liberar a próxima etapa do processo."
      >
        {queue.length === 0 ? (
          <EmptyState
            title="Nenhuma solicitação pendente"
            description="Quando um processo chegar ao TI, ele aparecerá nesta fila."
          />
        ) : (
          <div className="space-y-4">
            {queue.map((p) => (
              <ProcessSummary key={p.id} process={p}>
                {p.currentStep === 4 ? (
                  <StepAction
                    label="Backup e preservação das informações concluídos"
                    confirmLabel="Registrar conclusão do backup"
                    noteLabel="Registro do atendimento (local, escopo e responsável)"
                    onConfirm={(note) => {
                      completeStep(p.id, 4, user?.username ?? "TI", note);
                      toast.success(
                        "Atendimento de backup registrado. Etapa 5 liberada para o gestor.",
                      );
                    }}
                  />
                ) : (
                  <StepAction
                    label="Acessos bloqueados (e-mail, VPN, sistemas e redes)"
                    confirmLabel="Registrar conclusão do bloqueio"
                    noteLabel="Registro do atendimento (sistemas afetados, data/hora)"
                    requireNote
                    onConfirm={(note) => {
                      completeStep(p.id, 6, user?.username ?? "TI", note);
                      toast.success("Atendimento de bloqueio registrado. Próxima etapa liberada.");
                    }}
                  />
                )}
              </ProcessSummary>
            ))}
          </div>
        )}
      </Section>

      <Section
        title={`Aguardando etapas anteriores (${aguardando.length})`}
        description="Somente leitura — estas solicitações ainda não chegaram à fila do TI."
      >
        {aguardando.length === 0 ? (
          <p className="text-sm text-muted-foreground">
            Nenhum processo aguardando etapas anteriores.
          </p>
        ) : (
          <div className="space-y-4">
            {aguardando.map((p) => (
              <ProcessSummary key={p.id} process={p} />
            ))}
          </div>
        )}
      </Section>

      <div className="flex items-start gap-3 rounded-lg border border-border bg-muted/30 p-4 text-sm text-muted-foreground">
        <Clock3 className="mt-0.5 size-4 shrink-0" />
        <p>
          <strong className="text-foreground">Importante:</strong> o sistema funciona como controle
          da solicitação. A execução do backup, bloqueio e demais atividades continuam sendo
          responsabilidade do time de TI.
        </p>
      </div>
    </div>
  );
}
