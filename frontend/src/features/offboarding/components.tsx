import { useState, type ReactNode } from "react";
import { Check, Clock, Lock } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { cn, formatDate } from "@/lib/utils";
import { STEPS, isConcluded, statusOf, type OffboardingProcess, type StepState } from "./store";

export function ProcessStatusBadge({ process }: { process: OffboardingProcess }) {
  const status = statusOf(process);
  return <Badge variant="outline" className={cn("font-medium", process.rejected && "border-destructive/25 bg-destructive/10 text-destructive", isConcluded(process) && "border-emerald-500/25 bg-emerald-500/12 text-emerald-700 dark:text-emerald-400")}>{status}</Badge>;
}

export function ProcessSummary({ process, children }: { process: OffboardingProcess; children?: ReactNode }) {
  const current = STEPS.find((s) => s.n === process.currentStep);
  const progress = Math.min(100, Math.max(0, ((process.currentStep - 1) / STEPS.length) * 100));
  return <article className="rounded-xl border border-border bg-card p-4 shadow-sm sm:p-5">
    <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
      <div className="min-w-0"><div className="flex flex-wrap items-center gap-2"><h3 className="truncate text-sm font-semibold text-foreground">{process.colaborador}</h3><ProcessStatusBadge process={process} /></div><p className="mt-1 text-xs text-muted-foreground">Gestor: {process.gestor} · Desligamento previsto: {formatDate(process.dataPrevista)}{process.setor ? ` · ${process.setor}` : ""}</p><p className="mt-1 text-xs text-muted-foreground">Motivo: {process.motivo}</p></div>
      <div className="shrink-0 text-left sm:text-right"><p className="text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">Etapa {process.currentStep} de {STEPS.length}</p><p className="mt-0.5 text-xs font-medium text-foreground">{current?.label ?? "Concluído"}</p></div>
    </div>
    <div className="mt-4"><div className="mb-2 h-1.5 overflow-hidden rounded-full bg-muted"><div className="h-full rounded-full bg-primary transition-all" style={{ width: `${progress}%` }} /></div><StepTimeline process={process} /></div>
    {process.rejected ? <p className="mt-3 rounded-md border border-destructive/25 bg-destructive/5 p-2.5 text-xs text-destructive">Rejeitado por {process.rejected.by}: {process.rejected.reason}</p> : null}
    {children ? <div className="mt-4 border-t border-border pt-4">{children}</div> : null}
  </article>;
}

export function StepTimeline({ process }: { process: OffboardingProcess }) {
  return <ol className="grid grid-cols-2 gap-1.5 sm:flex sm:flex-wrap" aria-label="Etapas do desligamento">{STEPS.map((s) => { const done = process.currentStep > s.n; const current = process.currentStep === s.n && !process.rejected; return <li key={s.key}><span title={`${s.n}. ${s.label} — ${s.owner}`} className={cn("inline-flex w-full items-center gap-1 rounded-full border px-2 py-1 text-[11px] font-medium sm:w-auto", done && "border-emerald-500/25 bg-emerald-500/10 text-emerald-700 dark:text-emerald-400", current && "border-primary/30 bg-primary/10 text-primary", !done && !current && "border-border text-muted-foreground")}>{done ? <Check className="size-3 shrink-0" /> : current ? <Clock className="size-3 shrink-0" /> : <Lock className="size-3 shrink-0" />}{s.n}. {s.owner}</span></li>; })}</ol>;
}

export function StepAction({ label, confirmLabel, noteLabel = "Observação", requireNote = false, disabled = false, onConfirm }: { label: string; confirmLabel: string; noteLabel?: string; requireNote?: boolean; disabled?: boolean; onConfirm: (note: string) => void }) {
  const [note, setNote] = useState(""); const [checked, setChecked] = useState(false); const id = `step-${label.replace(/\W+/g, "-").toLowerCase()}`;
  return <div className="space-y-3"><div className="flex items-start gap-2"><Checkbox id={id} checked={checked} onCheckedChange={(v) => setChecked(v === true)} disabled={disabled} /><Label htmlFor={id} className="text-sm font-medium leading-snug">{label}</Label></div><div className="space-y-1.5"><Label htmlFor={`${id}-note`} className="text-xs text-muted-foreground">{noteLabel}{requireNote ? " (obrigatória)" : " (opcional)"}</Label><Textarea id={`${id}-note`} value={note} onChange={(e) => setNote(e.target.value)} rows={2} disabled={disabled} /></div><Button size="sm" disabled={disabled || !checked || (requireNote && !note.trim())} onClick={() => { onConfirm(note.trim()); setNote(""); setChecked(false); }}>{confirmLabel}</Button></div>;
}

export function InfoChecklist({ items, values, onChange, disabled = false }: { items: { key: string; label: string }[]; values: Partial<Record<string, StepState>>; onChange: (key: string, state: StepState) => void; disabled?: boolean }) {
  return <ul className="space-y-2.5">{items.map((item) => { const state = values[item.key]; return <li key={item.key} className="flex items-start gap-2"><Checkbox id={item.key} checked={state?.done ?? false} disabled={disabled} onCheckedChange={(v) => onChange(item.key, { done: v === true, at: new Date().toISOString() })} /><Label htmlFor={item.key} className="text-sm font-normal leading-snug">{item.label}</Label></li>; })}</ul>;
}
