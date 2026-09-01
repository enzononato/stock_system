import { useState } from "react";
import { toast } from "sonner";

import { PageHeader, Section } from "@/components/app/PageHeader";
import { EmptyState, ErrorState, LoadingState } from "@/components/app/StateBlocks";
import { SearchableSelect } from "@/components/app/SearchableSelect";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { useAuth } from "@/lib/auth";

import { ProcessSummary, StepAction } from "./components";
import { completeStep, createProcess, useOffboardingProcesses } from "./store";
import { useCollaborators } from "./useCollaborators";

export function OffboardingGestorPage() {
  const { user } = useAuth();
  const { collaborators, isLoading, error, refetch } = useCollaborators();
  const all = useOffboardingProcesses();
  const [colaborador, setColaborador] = useState("");
  const [dataPrevista, setDataPrevista] = useState("");
  const [motivo, setMotivo] = useState("");
  const meus = all.filter((p) => p.gestor === (user?.username ?? ""));
  const paraValidar = meus.filter((p) => !p.rejected && p.currentStep === 5);
  const selecionado = collaborators.find((c) => c.nome === colaborador);
  const jaAberto = all.some((p) => p.colaborador === colaborador && !p.rejected && p.currentStep <= 9);

  function submit(e: React.FormEvent) {
    e.preventDefault();
    if (!colaborador || !dataPrevista || !motivo.trim()) return toast.error("Preencha colaborador, data prevista e motivo.");
    if (jaAberto) return toast.error("Já existe um processo de desligamento em andamento para este colaborador.");
    createProcess({ colaborador, cpf: selecionado?.cpf, setor: selecionado?.setor, revenda: selecionado?.revenda, gestor: user?.username ?? "Gestor", dataPrevista, motivo: motivo.trim() });
    toast.success("Solicitação enviada ao RH para validação.");
    setColaborador(""); setDataPrevista(""); setMotivo("");
  }

  return (
    <div className="space-y-7">
      <PageHeader eyebrow="Desligamento" title="Solicitar desligamento" description="Abra uma solicitação e acompanhe o processo até a validação da transferência de atividades." />
      <div className="grid gap-3 sm:grid-cols-3">
        <div className="surface-panel p-4"><p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Em andamento</p><p className="mt-1 text-2xl font-bold num">{meus.filter((p) => !p.rejected && p.currentStep < 9).length}</p></div>
        <div className="surface-panel p-4"><p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Aguardando validação</p><p className="mt-1 text-2xl font-bold num">{paraValidar.length}</p></div>
        <div className="surface-panel p-4"><p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Solicitações abertas</p><p className="mt-1 text-2xl font-bold num">{meus.length}</p></div>
      </div>

      <Section title="Nova solicitação" description="Informe quem será desligado, a data prevista e o motivo da solicitação.">
        {isLoading ? <LoadingState label="Carregando colaboradores com patrimônio vinculado…" /> : error ? <ErrorState error={error} onRetry={() => void refetch()} /> : collaborators.length === 0 ? <EmptyState title="Nenhum colaborador com patrimônio vinculado" description="A lista é derivada dos itens emprestados retornados pela API." /> : (
          <form onSubmit={submit} className="space-y-5">
            <div className="grid gap-4 sm:grid-cols-2">
              <div className="space-y-1.5 sm:col-span-2"><Label htmlFor="colaborador">Colaborador *</Label><SearchableSelect id="colaborador" options={collaborators.map((c) => ({ value: c.nome, label: c.nome, subtitle: `${c.items.length} item(ns) vinculado(s)${c.setor ? ` · ${c.setor}` : ""}` }))} value={colaborador} onValueChange={setColaborador} placeholder="Selecione o colaborador" /></div>
              <div className="space-y-1.5"><Label htmlFor="data-prevista">Data prevista *</Label><Input id="data-prevista" type="date" value={dataPrevista} onChange={(e) => setDataPrevista(e.target.value)} /></div>
              <div className="space-y-1.5 sm:col-span-2"><Label htmlFor="motivo">Motivo *</Label><Textarea id="motivo" rows={4} placeholder="Descreva o motivo da solicitação…" value={motivo} onChange={(e) => setMotivo(e.target.value)} /></div>
            </div>
            {selecionado && <div className="rounded-lg border border-border bg-muted/30 p-4 text-sm"><p className="font-semibold">Patrimônio vinculado</p><p className="mt-1 text-muted-foreground">{selecionado.items.length} item(ns) · {selecionado.setor || "Setor não informado"} · {selecionado.revenda || "Revenda não informada"}</p></div>}
            <div className="flex justify-end border-t border-border pt-5"><Button type="submit">Enviar solicitação ao RH</Button></div>
          </form>
        )}
      </Section>

      <Section title={`Aguardando sua validação (${paraValidar.length})`} description="Processos que chegaram à etapa de transferência de atividades.">
        {paraValidar.length === 0 ? <p className="text-sm text-muted-foreground">Nenhum processo aguardando você na etapa 5.</p> : <div className="space-y-4">{paraValidar.map((p) => <ProcessSummary key={p.id} process={p}><StepAction label="Transferência de atividades concluída (tarefas, arquivos, senhas corporativas e conhecimento pendente)" confirmLabel="Concluir etapa 5" noteLabel="Observação (quem assumiu cada frente)" requireNote onConfirm={(note) => { completeStep(p.id, 5, user?.username ?? "Gestor", note); toast.success("Transferência validada. Etapa 6 liberada para o TI."); }} /></ProcessSummary>)}</div>}
      </Section>

      <Section title={`Minhas solicitações (${meus.length})`} description="Acompanhe o andamento das solicitações abertas por você.">
        {meus.length === 0 ? <p className="text-sm text-muted-foreground">Você ainda não abriu solicitações de desligamento.</p> : <div className="space-y-4">{meus.map((p) => <ProcessSummary key={p.id} process={p} />)}</div>}
      </Section>
    </div>
  );
}
