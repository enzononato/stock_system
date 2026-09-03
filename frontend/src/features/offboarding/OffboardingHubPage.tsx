import { useState } from "react";
import {
  Building2,
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

export function OffboardingHubPage() {
  const { user } = useAuth();
  const all = useOffboardingProcesses();
  const [activeTab, setActiveTab] = useState("panorama");

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

  return (
    <div className="space-y-6">
      <div className="border-b border-border pb-2">
        <Tabs value={activeTab} onValueChange={setActiveTab} className="w-full">
          <TabsList className="flex h-auto w-full flex-wrap justify-start gap-1 bg-transparent p-0">
            <TabsTrigger
              value="panorama"
              className="flex items-center gap-2 rounded-md border border-transparent px-3 py-2 text-xs font-semibold data-[state=active]:border-border data-[state=active]:bg-muted/60 data-[state=active]:text-foreground"
            >
              <ClipboardList className="size-3.5" />
              Panorama / DP
              {queueDp > 0 && (
                <span className="rounded-full bg-primary/20 px-1.5 py-0.5 text-[10px] font-bold text-primary">
                  {queueDp}
                </span>
              )}
            </TabsTrigger>

            <TabsTrigger
              value="gestor"
              className="flex items-center gap-2 rounded-md border border-transparent px-3 py-2 text-xs font-semibold data-[state=active]:border-border data-[state=active]:bg-muted/60 data-[state=active]:text-foreground"
            >
              <UserMinus className="size-3.5" />
              Gestor
            </TabsTrigger>

            <TabsTrigger
              value="rh"
              className="flex items-center gap-2 rounded-md border border-transparent px-3 py-2 text-xs font-semibold data-[state=active]:border-border data-[state=active]:bg-muted/60 data-[state=active]:text-foreground"
            >
              <UserCheck className="size-3.5" />
              RH
              {queueRh > 0 && (
                <span className="rounded-full bg-primary/20 px-1.5 py-0.5 text-[10px] font-bold text-primary">
                  {queueRh}
                </span>
              )}
            </TabsTrigger>

            <TabsTrigger
              value="ti"
              className="flex items-center gap-2 rounded-md border border-transparent px-3 py-2 text-xs font-semibold data-[state=active]:border-border data-[state=active]:bg-muted/60 data-[state=active]:text-foreground"
            >
              <HardDrive className="size-3.5" />
              TI
              {queueTi > 0 && (
                <span className="rounded-full bg-primary/20 px-1.5 py-0.5 text-[10px] font-bold text-primary">
                  {queueTi}
                </span>
              )}
            </TabsTrigger>

            <TabsTrigger
              value="patrimonio"
              className="flex items-center gap-2 rounded-md border border-transparent px-3 py-2 text-xs font-semibold data-[state=active]:border-border data-[state=active]:bg-muted/60 data-[state=active]:text-foreground"
            >
              <ShieldAlert className="size-3.5" />
              Patrimônio
              {queuePatrimonio > 0 && (
                <span className="rounded-full bg-primary/20 px-1.5 py-0.5 text-[10px] font-bold text-primary">
                  {queuePatrimonio}
                </span>
              )}
            </TabsTrigger>

            <TabsTrigger
              value="adm"
              className="flex items-center gap-2 rounded-md border border-transparent px-3 py-2 text-xs font-semibold data-[state=active]:border-border data-[state=active]:bg-muted/60 data-[state=active]:text-foreground"
            >
              <Landmark className="size-3.5" />
              Financeiro & Contábil
            </TabsTrigger>
          </TabsList>

          <div className="mt-6">
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
        </Tabs>
      </div>
    </div>
  );
}
