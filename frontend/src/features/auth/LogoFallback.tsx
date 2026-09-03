import { ShieldCheck, Cpu } from "lucide-react";

export function LogoFallback() {
  return (
    <div className="relative flex flex-col items-center justify-center p-8 w-full max-w-[420px]">
      {/* Glow e anéis concêntricos decorativos estilo Solvd */}
      <div className="absolute inset-0 -z-10 flex items-center justify-center">
        <div className="size-64 rounded-full bg-primary/10 blur-3xl" />
        <div className="absolute size-80 rounded-full border border-primary/10 opacity-40 animate-pulse" />
        <div className="absolute size-[380px] rounded-full border border-dashed border-primary/15 opacity-30" />
      </div>

      {/* Card/Insígnia geométrica Solvd */}
      <div className="relative flex flex-col items-center rounded-2xl border border-primary/20 bg-surface/80 p-8 backdrop-blur-xl shadow-2xl transition-transform duration-300 hover:border-primary/40">
        <div className="absolute -top-3 left-1/2 -translate-x-1/2 flex items-center gap-1.5 rounded-full border border-primary/30 bg-background/90 px-3 py-0.5 text-[10px] font-semibold tracking-wider text-primary uppercase">
          <Cpu className="size-3" />
          <span>Controle de Patrimônio</span>
        </div>

        {/* Emblema Revalle vetorizado */}
        <div className="relative my-4 flex size-28 items-center justify-center rounded-xl bg-gradient-to-br from-primary/20 via-background to-sidebar p-1 shadow-inner border border-primary/25">
          <img
            src="/logo-revalle.jpg"
            alt="Logo Revalle"
            className="size-full rounded-lg object-cover contrast-105 brightness-95"
          />
        </div>

        <h3 className="mt-2 text-base font-bold tracking-tight text-foreground text-center">
          Tecnologia & Governança de Ativos
        </h3>
        <p className="mt-1 text-xs text-muted-foreground text-center max-w-[260px]">
          Plataforma corporativa centralizada para auditoria, controle de estoque e ciclo de vida de
          TI.
        </p>

        <div className="mt-5 flex items-center gap-2 text-[11px] font-medium text-primary">
          <ShieldCheck className="size-3.5" />
          <span>Sessão protegida por criptografia</span>
        </div>
      </div>
    </div>
  );
}
