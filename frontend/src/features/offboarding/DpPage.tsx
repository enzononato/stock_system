import { useMemo } from "react";
import { toast } from "sonner";
import { CheckCircle2, ClipboardList, UserCheck } from "lucide-react";

import { PageHeader, Section } from "@/components/app/PageHeader";
import { Button } from "@/components/ui/button";
import { useAuth } from "@/lib/auth";
import { ProcessSummary } from "./components";
import { completeStep, setReturnedItems, useOffboardingProcesses } from "./store";
import { useCollaborators } from "./useCollaborators";

export function OffboardingDpPage() {
  const { user } = useAuth();
  const all = useOffboardingProcesses();
  const { collaborators } = useCollaborators();
  const paraAtivar = all.filter((p) => !p.rejected && p.currentStep === 3);
  const paraEncerrar = all.filter((p) => !p.rejected && p.currentStep === 9);
  const andamento = useMemo(
    () => all.filter((p) => !p.rejected && p.currentStep > 0 && p.currentStep < 10).length,
    [all],
  );

  return (
    <div className="space-y-7">
      <PageHeader
        eyebrow="Desligamento · DP"
        title="Ativação e encerramento"
        description="Ative os processos validados pelo RH, acompanhe o patrimônio e finalize o desligamento após todas as etapas."
      />
      <div className="grid gap-3 sm:grid-cols-3">
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-primary/10 text-primary">
            <UserCheck className="size-4" />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Para ativar
            </p>
            <p className="mt-1 text-2xl font-bold num">{paraAtivar.length}</p>
          </div>
        </div>
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-muted">
            <ClipboardList className="size-4" />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Em andamento
            </p>
            <p className="mt-1 text-2xl font-bold num">{andamento}</p>
          </div>
        </div>
        <div className="surface-panel flex items-center gap-3 p-4">
          <div className="grid size-9 place-items-center rounded-lg bg-muted">
            <CheckCircle2 className="size-4" />
          </div>
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
              Para encerrar
            </p>
            <p className="mt-1 text-2xl font-bold num">{paraEncerrar.length}</p>
          </div>
        </div>
      </div>

      <Section
        title={`Aguardando ativação (${paraAtivar.length})`}
        description="Após a validação do RH, o DP inicia formalmente o fluxo e registra o levantamento de patrimônio."
      >
        {paraAtivar.length === 0 ? (
          <p className="text-sm text-muted-foreground">Nenhum processo aguardando ativação.</p>
        ) : (
          <div className="space-y-4">
            {paraAtivar.map((p) => {
              const vinculados = collaborators.find((c) => c.nome === p.colaborador)?.items ?? [];
              return (
                <ProcessSummary key={p.id} process={p}>
                  <div className="rounded-lg border border-border bg-muted/20 p-4">
                    <p className="text-sm font-medium">Patrimônio vinculado atualmente</p>
                    <p className="mt-1 text-sm text-muted-foreground">
                      <strong className="text-foreground">{vinculados.length}</strong> item(ns)
                      localizado(s) para {p.colaborador}.
                    </p>
                  </div>
                  <Button
                    className="mt-4"
                    size="sm"
                    onClick={() => {
                      setReturnedItems(p.id, []);
                      completeStep(
                        p.id,
                        3,
                        user?.username ?? "DP",
                        `${vinculados.length} item(ns) levantados na ativação`,
                      );
                      toast.success("Fluxo ativado. Solicitação enviada ao TI.");
                    }}
                  >
                    Ativar fluxo e levantar patrimônio
                  </Button>
                </ProcessSummary>
              );
            })}
          </div>
        )}
      </Section>

      <Section
        title={`Prontos para encerramento (${paraEncerrar.length})`}
        description="Somente encerre depois de confirmar que as etapas anteriores foram concluídas."
      >
        {paraEncerrar.length === 0 ? (
          <p className="text-sm text-muted-foreground">Nenhum processo pronto para encerramento.</p>
        ) : (
          <div className="space-y-4">
            {paraEncerrar.map((p) => (
              <ProcessSummary key={p.id} process={p}>
                <div className="rounded-lg border border-emerald-500/20 bg-emerald-500/5 p-4 text-sm text-muted-foreground">
                  <p className="font-medium text-foreground">Processo liberado para encerramento</p>
                  <p className="mt-1">As etapas anteriores foram registradas no fluxo.</p>
                </div>
                <Button
                  className="mt-4"
                  size="sm"
                  onClick={() => {
                    completeStep(p.id, 9, user?.username ?? "DP");
                    toast.success("Desligamento concluído.");
                  }}
                >
                  Marcar desligamento como concluído
                </Button>
              </ProcessSummary>
            ))}
          </div>
        )}
      </Section>

      <Section
        title={`Panorama geral (${all.length})`}
        description="Visão consolidada para acompanhamento operacional do DP."
      >
        {all.length === 0 ? (
          <p className="text-sm text-muted-foreground">Nenhum processo registrado.</p>
        ) : (
          <div className="grid gap-3 md:grid-cols-2">
            {all.map((p) => (
              <ProcessSummary key={p.id} process={p} />
            ))}
          </div>
        )}
      </Section>
    </div>
  );
}
