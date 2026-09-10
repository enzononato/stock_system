import { useState } from "react";
import {
  Building2,
  CheckCircle2,
  CircleDashed,
  ClipboardList,
  HardDrive,
  Landmark,
  ShieldAlert,
  UserCheck,
  UserMinus,
} from "lucide-react";

import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { useAuth } from "@/lib/auth";

import { OffboardingGestorPage } from "./GestorPage";
import { OffboardingRhPage } from "./RhPage";
import { OffboardingDpPage } from "./DpPage";
import { OffboardingTiPage } from "./TiPage";
import { OffboardingSegurancaPage } from "./SegurancaPage";
import { OffboardingFinanceiroContabilidadePage } from "./FinanceiroContabilidadePage";
import { useOffboardingProcesses } from "./store";

const DEPARTMENTS = [
  { id: "panorama", label: "Panorama / DP", icon: ClipboardList, queueKey: "dp" },
  { id: "gestor", label: "Gestor", icon: UserMinus, queueKey: "gestor" },
  { id: "rh", label: "RH", icon: UserCheck, queueKey: "rh" },
  { id: "ti", label: "TI", icon: HardDrive, queueKey: "ti" },
  { id: "patrimonio", label: "Patrimônio", icon: ShieldAlert, queueKey: "patrimonio" },
  { id: "adm", label: "Financeiro & Contábil", icon: Landmark, queueKey: null },
] as const;

type DeptId = (typeof DEPARTMENTS)[number]["id"];

export function OffboardingHubPage() {
  const { user } = useAuth();
  const all = useOffboardingProcesses();
  const [activeTab, setActiveTab] = useState<DeptId>("panorama");

  const queueTi = all.filter(
    (p) => !p.rejected && (p.currentStep === 4 || p.currentStep === 6),
  ).length;
  const queueRh = all.filter(
    (p) => !p.rejected && (p.currentStep === 2 || p.currentStep === 8),
  ).length;
  const queueDp = all.filter(
    (p) => !p.rejected && (p.currentStep === 3 || p.currentStep === 9),
  ).length;
  const queuePatrimonio = all.filter((p) => !p.rejected && p.currentStep === 7).length;

  const queues: Record<string, number> = {
    dp: queueDp,
    rh: queueRh,
    ti: queueTi,
    patrimonio: queuePatrimonio,
    gestor: 0,
  };

  // Progresso do processo selecionado (panorama geral)
  const activeProcesses = all.filter((p) => !p.rejected && p.currentStep < 10);
  const completedProcesses = all.filter((p) => !p.rejected && p.currentStep >= 10);
  const rejectedProcesses = all.filter((p) => p.rejected);

  const totalQueue = queueDp + queueRh + queueTi + queuePatrimonio;

  return (
    <div className="page-container-dense space-y-0">
      {/* Header Operacional */}
      <div className="border-b border-border pb-4 mb-6">
        <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <div className="flex items-center gap-2">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
                Operações de Saída
              </span>
              <span className="text-muted-foreground">•</span>
              <span className="text-caption text-foreground font-medium">Hub Interdepartamental</span>
            </div>
            <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">
              Offboarding de Colaboradores
            </h1>
            <p className="text-body-sm text-muted-foreground mt-1">
              Gerenciamento integrado de processos de desligamento — fluxo multidepartamental com checklist operacional por setor.
            </p>
          </div>

          {/* Status Geral */}
          <div className="flex items-center gap-4 text-caption shrink-0">
            <div className="text-center">
              <span className="font-mono text-body-lg font-bold text-foreground block">{activeProcesses.length}</span>
              <span className="text-muted-foreground">Em andamento</span>
            </div>
            <div className="w-px h-8 bg-border" />
            <div className="text-center">
              <span className="font-mono text-body-lg font-bold text-foreground block">{completedProcesses.length}</span>
              <span className="text-muted-foreground">Concluídos</span>
            </div>
            <div className="w-px h-8 bg-border" />
            <div className="text-center">
              <span className="font-mono text-body-lg font-bold text-foreground block">{totalQueue}</span>
              <span className="text-muted-foreground">Pendentes</span>
            </div>
          </div>
        </div>

        {/* Barra de Status de Departamentos */}
        {totalQueue > 0 && (
          <div className="mt-4 flex flex-wrap gap-2">
            {DEPARTMENTS.filter((d) => d.queueKey && (queues[d.queueKey] ?? 0) > 0).map((dept) => (
              <button
                key={dept.id}
                type="button"
                onClick={() => setActiveTab(dept.id)}
                className={`flex items-center gap-1.5 text-caption rounded-[4px] border px-2.5 py-1 transition-colors ${
                  activeTab === dept.id
                    ? "border-foreground bg-foreground text-background"
                    : "border-border bg-surface text-foreground hover:bg-muted/40"
                }`}
              >
                <dept.icon className="size-3" />
                <span className="font-medium">{dept.label}</span>
                <span className={`font-mono font-bold rounded-full px-1 text-[10px] ${
                  activeTab === dept.id ? "bg-background/20" : "bg-foreground/10"
                }`}>
                  {dept.queueKey ? (queues[dept.queueKey] ?? 0) : ""}
                </span>
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Tabs com Navegação por Setor */}
      <Tabs value={activeTab} onValueChange={(v) => setActiveTab(v as DeptId)} className="w-full">
        {/* Barra de Abas — Strip Lateral */}
        <div className="grid grid-cols-1 lg:grid-cols-[200px_1fr] gap-6 items-start">
          {/* Sidebar de Departamentos */}
          <div className="rounded-[6px] border border-border bg-surface p-3 space-y-1">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground px-2 block mb-2">
              Departamentos
            </span>
            <TabsList className="flex flex-col h-auto w-full bg-transparent p-0 gap-0.5">
              {DEPARTMENTS.map((dept) => {
                const q = dept.queueKey ? (queues[dept.queueKey] ?? 0) : 0;
                const isActive = activeTab === dept.id;
                return (
                  <TabsTrigger
                    key={dept.id}
                    value={dept.id}
                    className={`flex items-center justify-between gap-2 w-full rounded-[4px] px-3 py-2.5 text-left text-body-sm font-medium transition-all
                      data-[state=active]:bg-foreground data-[state=active]:text-background
                      data-[state=inactive]:bg-transparent data-[state=inactive]:text-muted-foreground data-[state=inactive]:hover:bg-muted/30 data-[state=inactive]:hover:text-foreground`}
                  >
                    <div className="flex items-center gap-2">
                      <dept.icon className="size-3.5 shrink-0" />
                      <span>{dept.label}</span>
                    </div>
                    {q > 0 && (
                      <span className={`font-mono text-[10px] font-bold rounded-[2px] px-1.5 py-0.5 ${
                        isActive ? "bg-background/20 text-background" : "bg-foreground/10 text-foreground"
                      }`}>
                        {q}
                      </span>
                    )}
                  </TabsTrigger>
                );
              })}
            </TabsList>

            {/* Legenda de Status */}
            <div className="pt-3 mt-3 border-t border-border space-y-1.5 text-caption text-muted-foreground">
              <div className="flex items-center gap-2">
                <CircleDashed className="size-3 text-muted-foreground" />
                <span>Pendente de ação</span>
              </div>
              <div className="flex items-center gap-2">
                <CheckCircle2 className="size-3 text-foreground" />
                <span className="text-foreground">Concluído</span>
              </div>
            </div>
          </div>

          {/* Conteúdo do Departamento */}
          <div className="min-w-0">
            <TabsContent value="panorama" className="m-0 focus-visible:outline-none">
              <OffboardingDpPage />
            </TabsContent>

            <TabsContent value="gestor" className="m-0 focus-visible:outline-none">
              <OffboardingGestorPage />
            </TabsContent>

            <TabsContent value="rh" className="m-0 focus-visible:outline-none">
              <OffboardingRhPage />
            </TabsContent>

            <TabsContent value="ti" className="m-0 focus-visible:outline-none">
              <OffboardingTiPage />
            </TabsContent>

            <TabsContent value="patrimonio" className="m-0 focus-visible:outline-none">
              <OffboardingSegurancaPage />
            </TabsContent>

            <TabsContent value="adm" className="m-0 focus-visible:outline-none">
              <OffboardingFinanceiroContabilidadePage />
            </TabsContent>
          </div>
        </div>
      </Tabs>
    </div>
  );
}
