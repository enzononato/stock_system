import { useCallback, useSyncExternalStore } from "react";

/**
 * Módulo de Desligamento de Colaboradores — SOMENTE FRONTEND.
 *
 * O backend atual não expõe nenhum endpoint de processo de desligamento
 * (não há `/offboarding`, nem campos de etapa/status de desligamento nos
 * recursos existentes). Como esta tarefa é exclusivamente de frontend e não é
 * permitido criar endpoints, o estado do processo é mantido no navegador
 * (localStorage), enquanto TODO o dado de patrimônio (itens vinculados ao
 * colaborador, devolução, termos) continua vindo exclusivamente da API real.
 *
 * Quando o time de backend expuser o recurso, basta trocar esta camada por
 * chamadas HTTP: a UI consome apenas as funções exportadas aqui.
 */

export const STEPS = [
  { n: 1, key: "solicitacao", label: "Solicitação do gestor", owner: "Gestor" },
  { n: 2, key: "validacao_rh", label: "Validação do RH", owner: "RH" },
  { n: 3, key: "ativacao_dp", label: "Ativação pelo DP", owner: "DP" },
  { n: 4, key: "backup_ti", label: "Backup e preservação (TI)", owner: "TI" },
  { n: 5, key: "transferencia_gestor", label: "Transferência de atividades", owner: "Gestor" },
  { n: 6, key: "bloqueio_ti", label: "Bloqueio de acessos (TI)", owner: "TI" },
  { n: 7, key: "devolucao_patrimonio", label: "Devolução de patrimônio", owner: "Admin/Segurança" },
  { n: 8, key: "conferencia_rh", label: "Conferência documental (RH)", owner: "RH" },
  { n: 9, key: "encerramento_dp", label: "Encerramento (DP)", owner: "DP" },
] as const;

export type StepKey = (typeof STEPS)[number]["key"];

export interface StepState {
  done: boolean;
  note?: string;
  at?: string;
  by?: string;
}

export interface OffboardingProcess {
  id: string;
  colaborador: string;
  cpf?: string;
  setor?: string;
  revenda?: string;
  gestor: string;
  dataPrevista: string;
  motivo: string;
  createdAt: string;
  /** Etapa atual (1..9). 10 = concluído. */
  currentStep: number;
  rejected?: { reason: string; at: string; by: string };
  steps: Partial<Record<StepKey, StepState>>;
  /** Itens de patrimônio marcados como devolvidos nesta etapa 7 (ids da API). */
  returnedItemIds: number[];
  /** Checklists puramente informativos. */
  financeiro: Partial<Record<string, StepState>>;
  contabilidade: Partial<Record<string, StepState>>;
  termoGerado?: boolean;
}

const STORAGE_KEY = "controle-patrimonio:offboarding:v1";

let cache: OffboardingProcess[] | null = null;
const listeners = new Set<() => void>();

function read(): OffboardingProcess[] {
  if (cache) return cache;
  if (typeof window === "undefined") return (cache = []);
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    cache = raw ? (JSON.parse(raw) as OffboardingProcess[]) : [];
  } catch {
    cache = [];
  }
  return cache;
}

function write(next: OffboardingProcess[]) {
  cache = next;
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  } catch {
    /* quota/private mode: mantém só em memória */
  }
  listeners.forEach((l) => l());
}

function subscribe(listener: () => void) {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

export function useOffboardingProcesses(): OffboardingProcess[] {
  return useSyncExternalStore(subscribe, read, () => [] as OffboardingProcess[]);
}

export function stepLabel(n: number): string {
  if (n > 9) return "Concluído";
  return STEPS.find((s) => s.n === n)?.label ?? "—";
}

export function stepOwner(n: number): string {
  return STEPS.find((s) => s.n === n)?.owner ?? "—";
}

export function isConcluded(p: OffboardingProcess): boolean {
  return p.currentStep > 9;
}

export function statusOf(p: OffboardingProcess): string {
  if (p.rejected) return "Rejeitado";
  if (isConcluded(p)) return "Concluído";
  return `Etapa ${p.currentStep} · ${stepOwner(p.currentStep)}`;
}

function update(id: string, fn: (p: OffboardingProcess) => OffboardingProcess) {
  write(read().map((p) => (p.id === id ? fn(p) : p)));
}

export function createProcess(input: {
  colaborador: string;
  cpf?: string | undefined;
  setor?: string | undefined;
  revenda?: string | undefined;
  gestor: string;
  dataPrevista: string;
  motivo: string;
}): OffboardingProcess {
  const process: OffboardingProcess = {
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    colaborador: input.colaborador,
    ...(input.cpf ? { cpf: input.cpf } : {}),
    ...(input.setor ? { setor: input.setor } : {}),
    ...(input.revenda ? { revenda: input.revenda } : {}),
    gestor: input.gestor,
    dataPrevista: input.dataPrevista,
    motivo: input.motivo,
    createdAt: new Date().toISOString(),
    currentStep: 2,
    steps: { solicitacao: { done: true, at: new Date().toISOString(), by: input.gestor } },
    returnedItemIds: [],
    financeiro: {},
    contabilidade: {},
  };
  write([process, ...read()]);
  return process;
}

/** Conclui a etapa `step` de um processo, avançando o fluxo sequencialmente. */
export function completeStep(id: string, step: number, by: string, note?: string) {
  update(id, (p) => {
    if (p.currentStep !== step || p.rejected) return p;
    const key = STEPS.find((s) => s.n === step)!.key;
    return {
      ...p,
      currentStep: step + 1,
      steps: {
        ...p.steps,
        [key]: { done: true, at: new Date().toISOString(), by, ...(note ? { note } : {}) },
      },
    };
  });
}

export function rejectProcess(id: string, reason: string, by: string) {
  update(id, (p) => ({ ...p, rejected: { reason, at: new Date().toISOString(), by } }));
}

export function setReturnedItems(id: string, itemIds: number[]) {
  update(id, (p) => ({ ...p, returnedItemIds: itemIds }));
}

export function markItemReturned(id: string, itemId: number, note?: string) {
  update(id, (p) => ({
    ...p,
    returnedItemIds: p.returnedItemIds.includes(itemId) ? p.returnedItemIds : [...p.returnedItemIds, itemId],
    steps: note
      ? { ...p.steps, devolucao_patrimonio: { done: false, note, at: new Date().toISOString() } }
      : p.steps,
  }));
}

export function setTermoGerado(id: string) {
  update(id, (p) => ({ ...p, termoGerado: true }));
}

export function setChecklist(
  id: string,
  area: "financeiro" | "contabilidade",
  key: string,
  state: StepState,
) {
  update(id, (p) => ({ ...p, [area]: { ...p[area], [key]: state } }));
}

/** Processos cuja etapa atual pertence a um dos setores informados. */
export function useQueue(steps: number[]): OffboardingProcess[] {
  const all = useOffboardingProcesses();
  const filter = useCallback(
    (p: OffboardingProcess) => !p.rejected && steps.includes(p.currentStep),
    [steps],
  );
  return all.filter(filter);
}
