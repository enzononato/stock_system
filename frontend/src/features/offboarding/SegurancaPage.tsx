import { useState } from "react";
import { Download } from "lucide-react";
import { toast } from "sonner";

import { PageHeader, Section } from "@/components/app/PageHeader";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { FileUpload } from "@/components/app/FileUpload";
import { StatusBadge } from "@/components/app/StatusBadge";
import { LoadingState } from "@/components/app/StateBlocks";
import { confirmReturn, downloadReturnTerm } from "@/api/loans";
import { getErrorMessage } from "@/lib/api-error";
import { useAuth } from "@/lib/auth";

import { ProcessSummary } from "./components";
import { completeStep, markItemReturned, setTermoGerado, useOffboardingProcesses } from "./store";
import { useCollaborators } from "./useCollaborators";

/**
 * Etapa 7 — devolução do patrimônio. Reaproveita integralmente o fluxo de
 * devolução já existente na API (`/loans/{id}/return/initiate` para gerar o
 * termo e `/loans/{id}/return/confirm` para confirmar com o PDF assinado); o
 * status do item volta a "disponível" pelo próprio backend.
 */
export function OffboardingSegurancaPage() {
  const { user } = useAuth();
  const all = useOffboardingProcesses();
  const { collaborators, isLoading, refetch } = useCollaborators();
  const [pdfs, setPdfs] = useState<Record<number, File | null>>({});
  const [obs, setObs] = useState<Record<number, string>>({});
  const [busy, setBusy] = useState<number | null>(null);

  const fila = all.filter((p) => !p.rejected && p.currentStep === 7);

  async function baixarTermo(itemId: number) {
    try {
      setBusy(itemId);
      await downloadReturnTerm(itemId);
      toast.success("Termo de devolução gerado.");
    } catch (e) {
      toast.error(getErrorMessage(e));
    } finally {
      setBusy(null);
    }
  }

  async function confirmar(processId: string, itemId: number) {
    const pdf = pdfs[itemId];
    if (!pdf) {
      toast.error("Anexe o termo assinado para confirmar a devolução.");
      return;
    }
    try {
      setBusy(itemId);
      await confirmReturn(itemId, pdf);
      markItemReturned(processId, itemId, obs[itemId]);
      setPdfs((s) => ({ ...s, [itemId]: null }));
      await refetch();
      toast.success("Devolução confirmada. Item liberado no estoque.");
    } catch (e) {
      toast.error(getErrorMessage(e));
    } finally {
      setBusy(null);
    }
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Desligamento"
        title="Devolução de patrimônio (Admin/Segurança)"
        description="Liberada após o bloqueio de acessos pelo TI. Cada item devolvido usa o mesmo fluxo de devolução do sistema, com termo assinado."
      />

      <Section title={`Na sua fila (${fila.length})`}>
        {isLoading ? (
          <LoadingState label="Carregando patrimônio vinculado…" />
        ) : fila.length === 0 ? (
          <p className="text-sm text-muted-foreground">Nenhum processo aguardando devolução de patrimônio.</p>
        ) : (
          <div className="space-y-4">
            {fila.map((p) => {
              const itens = collaborators.find((c) => c.nome === p.colaborador)?.items ?? [];
              const pendentes = itens.filter((i) => !p.returnedItemIds.includes(i.id));
              return (
                <ProcessSummary key={p.id} process={p}>
                  {itens.length === 0 && p.returnedItemIds.length === 0 ? (
                    <p className="text-sm text-muted-foreground">
                      Nenhum item de patrimônio vinculado a este colaborador na API.
                    </p>
                  ) : (
                    <ul className="space-y-4">
                      {pendentes.map((item) => (
                        <li key={item.id} className="rounded-md border border-border p-3">
                          <div className="flex flex-wrap items-center justify-between gap-2">
                            <div className="min-w-0">
                              <p className="truncate text-sm font-medium text-foreground">
                                {item.tipo} · {item.brand} {item.model}
                              </p>
                              <p className="text-xs text-muted-foreground">
                                Identificador: {item.identificador ?? "—"}
                              </p>
                            </div>
                            <StatusBadge status={item.status} />
                          </div>

                          <div className="mt-3 grid gap-3 sm:grid-cols-2">
                            <div className="space-y-1.5">
                              <Label htmlFor={`obs-${item.id}`}>Estado de conservação / observação</Label>
                              <Input
                                id={`obs-${item.id}`}
                                value={obs[item.id] ?? ""}
                                onChange={(e) => setObs((s) => ({ ...s, [item.id]: e.target.value }))}
                              />
                            </div>
                            <FileUpload
                              id={`pdf-${item.id}`}
                              label="Termo de devolução assinado (PDF)"
                              onFile={(f) => setPdfs((s) => ({ ...s, [item.id]: f }))}
                            />
                          </div>

                          <div className="mt-3 flex flex-wrap gap-2">
                            <Button
                              size="sm"
                              variant="outline"
                              disabled={busy === item.id}
                              onClick={() => void baixarTermo(item.id)}
                            >
                              <Download className="size-4" aria-hidden /> Gerar termo
                            </Button>
                            <Button
                              size="sm"
                              disabled={busy === item.id || !pdfs[item.id]}
                              onClick={() => void confirmar(p.id, item.id)}
                            >
                              Confirmar devolução
                            </Button>
                          </div>
                        </li>
                      ))}
                    </ul>
                  )}

                  <Button
                    className="mt-4"
                    size="sm"
                    disabled={pendentes.length > 0}
                    onClick={() => {
                      setTermoGerado(p.id);
                      completeStep(
                        p.id,
                        7,
                        user?.username ?? "Admin/Segurança",
                        `${p.returnedItemIds.length} item(ns) devolvido(s)`,
                      );
                      toast.success("Devolução concluída. Etapa 8 liberada para o RH.");
                    }}
                  >
                    Concluir etapa 7
                  </Button>
                </ProcessSummary>
              );
            })}
          </div>
        )}
      </Section>
    </div>
  );
}
