import {
  BellRing,
  CheckCircle2,
  Clock3,
  FileText,
  Headphones,
  ShieldCheck,
  UserRoundX,
} from "lucide-react";

import { PageHeader, Section } from "@/components/app/PageHeader";
import { Button } from "@/components/ui/button";

const requestSteps = [
  {
    title: "Solicitação de backup",
    description:
      "Enviar ao time de TI a solicitação para preservação das informações do colaborador.",
  },
  {
    title: "Bloqueio de acessos",
    description:
      "O time de TI recebe a demanda para executar os bloqueios previstos no desligamento.",
  },
  {
    title: "Retorno do chamado",
    description: "Registrar o retorno do TI quando o atendimento da solicitação estiver concluído.",
  },
];

export function TiOffboardingPage() {
  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Desligamentos · TI"
        title="Solicitações para o time de TI"
        description="A etapa de TI funciona como uma demanda interna: o desligamento gera solicitações para que o time responsável faça o atendimento."
      />

      <div className="grid gap-4 sm:grid-cols-3">
        <div className="surface-panel p-4">
          <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
            Tipo
          </p>
          <p className="mt-1 font-semibold">Chamado interno</p>
        </div>
        <div className="surface-panel p-4">
          <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
            Responsável
          </p>
          <p className="mt-1 font-semibold">Time de TI</p>
        </div>
        <div className="surface-panel p-4">
          <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
            Execução
          </p>
          <p className="mt-1 font-semibold">Fora do sistema</p>
        </div>
      </div>

      <div className="grid gap-4 lg:grid-cols-[minmax(0,1fr)_320px]">
        <section className="surface-panel p-5 sm:p-6">
          <div className="flex items-start gap-3">
            <div className="grid size-10 shrink-0 place-items-center rounded-lg bg-primary/10 text-primary">
              <Headphones className="size-5" aria-hidden />
            </div>
            <div>
              <p className="font-semibold">Chamado de TI</p>
              <p className="mt-1 text-sm leading-relaxed text-muted-foreground">
                Esta tela representa a solicitação e o acompanhamento do atendimento, não a execução
                técnica do backup ou dos bloqueios.
              </p>
            </div>
          </div>

          <div className="mt-6">
            <Section
              title="Fluxo da solicitação"
              description="O processo fica preparado para receber o retorno do time responsável quando a integração de chamados estiver disponível."
            >
              <div className="space-y-3">
                {requestSteps.map((step, index) => (
                  <div
                    key={step.title}
                    className="flex items-start gap-4 rounded-lg border border-border p-4 transition-colors hover:bg-muted/20"
                  >
                    <div className="grid size-8 shrink-0 place-items-center rounded-full bg-muted text-xs font-bold num">
                      {index + 1}
                    </div>
                    <div className="min-w-0 flex-1">
                      <p className="text-sm font-semibold">{step.title}</p>
                      <p className="mt-1 text-sm leading-relaxed text-muted-foreground">
                        {step.description}
                      </p>
                    </div>
                    <Clock3
                      className="mt-0.5 size-4 shrink-0 text-muted-foreground"
                      aria-label="Aguardando atendimento"
                    />
                  </div>
                ))}
              </div>
            </Section>
          </div>

          <div className="mt-5 rounded-lg border border-border bg-muted/20 p-4">
            <div className="flex items-center gap-2 text-sm font-semibold">
              <BellRing className="size-4 text-primary" aria-hidden />
              Solicitação de backup
            </div>
            <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
              Quando um desligamento chegar à etapa de TI, o sistema deve informar o time
              responsável para realizar o backup necessário e devolver o status do atendimento.
            </p>
            <div className="mt-4 flex flex-wrap items-center gap-3">
              <Button disabled>
                <FileText className="mr-2 size-4" aria-hidden />
                Solicitar backup
              </Button>
              <span className="text-xs text-muted-foreground">
                Integração de chamados ainda não disponível na API atual.
              </span>
            </div>
          </div>
        </section>

        <aside className="surface-panel p-5">
          <div className="flex items-center gap-2">
            <UserRoundX className="size-4 text-primary" aria-hidden />
            <h2 className="text-sm font-semibold">Fila de solicitações</h2>
          </div>
          <div className="mt-5 rounded-lg border border-dashed border-border p-6 text-center">
            <ShieldCheck className="mx-auto size-7 text-muted-foreground" aria-hidden />
            <p className="mt-3 text-sm font-semibold">Nenhum chamado disponível</p>
            <p className="mt-1 text-xs leading-relaxed text-muted-foreground">
              A API atual não fornece um recurso de chamados de desligamento. A interface não cria
              nem simula uma solicitação real.
            </p>
          </div>
          <div className="mt-5 rounded-lg bg-muted/30 p-4">
            <div className="flex items-center gap-2 text-sm font-semibold">
              <CheckCircle2 className="size-4 text-primary" aria-hidden />
              Regra funcional
            </div>
            <p className="mt-2 text-xs leading-relaxed text-muted-foreground">
              Solicitação → atendimento pelo TI → retorno do chamado.
            </p>
          </div>
        </aside>
      </div>
    </div>
  );
}
