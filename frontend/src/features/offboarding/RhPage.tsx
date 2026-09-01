import { useState } from "react";
import { toast } from "sonner";
import { CheckCircle2, ClipboardCheck, FileCheck2, XCircle } from "lucide-react";

import { PageHeader, Section } from "@/components/app/PageHeader";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useAuth } from "@/lib/auth";

import { InfoChecklist, ProcessSummary } from "./components";
import { completeStep, rejectProcess, setChecklist, useOffboardingProcesses } from "./store";

const CONFERENCIA_RH = [
  { key: "exame_demissional", label: "Exame demissional / toxicológico realizado" },
  { key: "pesquisa_desligamento", label: "Pesquisa de desligamento aplicada" },
  { key: "espelho_ponto", label: "Espelho de ponto conferido" },
];

export function OffboardingRhPage() {
  const { user } = useAuth();
  const all = useOffboardingProcesses();
  const [rejeicao, setRejeicao] = useState<Record<string, string>>({});
  const paraValidar = all.filter((p) => !p.rejected && p.currentStep === 2);
  const paraConferir = all.filter((p) => !p.rejected && p.currentStep === 8);

  return <div className="space-y-7">
    <PageHeader eyebrow="Desligamento · RH" title="Validação e conferência" description="Valide as solicitações dos gestores e conclua a conferência documental no final do processo." />
    <div className="grid gap-3 sm:grid-cols-3">
      <div className="surface-panel flex items-center gap-3 p-4"><div className="grid size-9 place-items-center rounded-lg bg-primary/10 text-primary"><ClipboardCheck className="size-4" /></div><div><p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Para validar</p><p className="mt-1 text-2xl font-bold num">{paraValidar.length}</p></div></div>
      <div className="surface-panel flex items-center gap-3 p-4"><div className="grid size-9 place-items-center rounded-lg bg-muted"><FileCheck2 className="size-4" /></div><div><p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Conferência final</p><p className="mt-1 text-2xl font-bold num">{paraConferir.length}</p></div></div>
      <div className="surface-panel flex items-center gap-3 p-4"><div className="grid size-9 place-items-center rounded-lg bg-muted"><CheckCircle2 className="size-4" /></div><div><p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Visão</p><p className="mt-1 text-sm font-bold">Pendências do RH</p></div></div>
    </div>

    <Section title={`Pendentes de validação (${paraValidar.length})`} description="Aprovação libera o processo para o Departamento Pessoal. Rejeições exigem justificativa.">
      {paraValidar.length === 0 ? <p className="text-sm text-muted-foreground">Nenhuma solicitação pendente de validação.</p> : <div className="space-y-4">{paraValidar.map((p) => <ProcessSummary key={p.id} process={p}><div className="rounded-lg border border-border bg-muted/20 p-4"><div className="flex items-center gap-2 text-sm font-medium"><ClipboardCheck className="size-4 text-primary" />Revisão da solicitação</div><p className="mt-1 text-xs text-muted-foreground">Confira os dados do colaborador e a justificativa antes de liberar o processo.</p></div><div className="mt-4 flex flex-col gap-3 sm:flex-row sm:items-end"><div className="flex-1 space-y-1.5"><Label htmlFor={`rej-${p.id}`}>Justificativa para rejeição</Label><Input id={`rej-${p.id}`} placeholder="Obrigatória somente para rejeitar" value={rejeicao[p.id] ?? ""} onChange={(e) => setRejeicao((r) => ({ ...r, [p.id]: e.target.value }))} /></div><div className="flex gap-2"><Button onClick={() => { completeStep(p.id, 2, user?.username ?? "RH"); toast.success("Desligamento validado. Etapa 3 liberada para o DP."); }}><CheckCircle2 className="mr-1.5 size-4" />Aprovar</Button><Button variant="destructive" disabled={!(rejeicao[p.id] ?? "").trim()} onClick={() => { rejectProcess(p.id, (rejeicao[p.id] ?? "").trim(), user?.username ?? "RH"); toast.success("Solicitação rejeitada."); }}><XCircle className="mr-1.5 size-4" />Rejeitar</Button></div></div></ProcessSummary>)}</div>}
    </Section>

    <Section title={`Conferência documental final (${paraConferir.length})`} description="Checklist informativo — não há integração real com exames, pesquisa e ponto neste sistema.">
      {paraConferir.length === 0 ? <p className="text-sm text-muted-foreground">Nenhum processo aguardando conferência do RH.</p> : <div className="space-y-4">{paraConferir.map((p) => { const completo = CONFERENCIA_RH.every((c) => p.financeiro[`rh_${c.key}`]?.done); return <ProcessSummary key={p.id} process={p}><InfoChecklist items={CONFERENCIA_RH.map((c) => ({ key: `rh_${c.key}`, label: c.label }))} values={p.financeiro} onChange={(key, state) => setChecklist(p.id, "financeiro", key, state)} /><Button className="mt-4" size="sm" disabled={!completo} onClick={() => { completeStep(p.id, 8, user?.username ?? "RH"); toast.success("Conferência concluída. Etapa 9 liberada para o DP."); }}>Concluir conferência</Button></ProcessSummary>; })}</div>}
    </Section>
  </div>;
}
