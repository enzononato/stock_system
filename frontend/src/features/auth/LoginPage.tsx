import { useEffect, useState, lazy, Suspense, type FormEvent } from "react";
import { useNavigate } from "@tanstack/react-router";
import {
  Eye,
  EyeOff,
  Lock,
  ShieldCheck,
  User,
  Sparkles,
  Server,
  FileCheck2,
  History,
} from "lucide-react";

import { useAuth } from "@/lib/auth";
import { getErrorMessage } from "@/lib/api-error";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { LogoFallback } from "./LogoFallback";

// Carregamento sob demanda do WebGL Three.js (isolado do bundle principal)
const ThreeLogoCanvas = lazy(() => import("./ThreeLogoCanvas"));

export function LoginPage() {
  const { login, user, isLoading } = useAuth();
  const navigate = useNavigate();

  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [error, setError] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [isDesktop, setIsDesktop] = useState(false);

  useEffect(() => {
    if (!isLoading && user) void navigate({ to: "/", replace: true });
  }, [isLoading, user, navigate]);

  useEffect(() => {
    // Detecta viewport desktop para renderização seletiva do 3D
    const checkDesktop = () => {
      setIsDesktop(window.innerWidth >= 1024);
    };
    checkDesktop();
    window.addEventListener("resize", checkDesktop);
    return () => window.removeEventListener("resize", checkDesktop);
  }, []);

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setError("");
    setSubmitting(true);
    try {
      await login(username, password);
      void navigate({ to: "/", replace: true });
    } catch (err) {
      setError(getErrorMessage(err, "Usuário ou senha inválidos. Verifique suas credenciais."));
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="relative min-h-screen bg-background overflow-hidden flex flex-col lg:grid lg:grid-cols-[1.2fr_minmax(420px,520px)]">
      {/* Background radial e malha tecnológica (Solvd Style) */}
      <div
        className="pointer-events-none absolute inset-0 -z-10 bg-grid-tech opacity-40"
        aria-hidden
      />
      <div
        className="pointer-events-none absolute -top-40 -left-40 size-[600px] rounded-full bg-primary/10 blur-[120px] -z-10"
        aria-hidden
      />
      <div
        className="pointer-events-none absolute bottom-0 right-1/3 size-[500px] rounded-full bg-primary/5 blur-[100px] -z-10"
        aria-hidden
      />

      {/* PAINEL VISUAL E MOTION (Spotify Experience + Solvd Tech Identity) */}
      <section className="relative hidden lg:flex flex-col justify-between p-10 xl:p-14 border-r border-border/60 bg-gradient-to-br from-sidebar/95 via-background to-sidebar/90 overflow-hidden">
        {/* Topo institucional */}
        <div className="relative z-10 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex size-9 items-center justify-center rounded-lg border border-primary/30 bg-primary/10 shadow-sm">
              <img src="/logo-revalle.jpg" alt="Revalle" className="size-7 rounded object-cover" />
            </div>
            <div>
              <span className="text-sm font-bold tracking-tight text-foreground">Revalle</span>
              <span className="ml-2 rounded border border-primary/20 bg-primary/10 px-1.5 py-0.2 text-[10px] font-semibold text-primary uppercase">
                Enterprise Core
              </span>
            </div>
          </div>

          <div className="flex items-center gap-2 rounded-full border border-border/80 bg-surface/60 px-3 py-1 text-xs text-muted-foreground backdrop-blur-md">
            <span className="size-2 rounded-full bg-emerald-500 animate-pulse" />
            <span>Sistema Operacional</span>
          </div>
        </div>

        {/* Centro: Elemento 3D interativo ou Fallback */}
        <div className="relative z-10 my-auto flex h-[460px] w-full items-center justify-center">
          {/* Anéis decorativos Solvd */}
          <div className="pointer-events-none absolute inset-0 flex items-center justify-center -z-10">
            <div className="size-[420px] rounded-full border border-primary/10 opacity-60" />
            <div className="size-[520px] rounded-full border border-dashed border-primary/15 opacity-40" />
          </div>

          {isDesktop ? (
            <Suspense fallback={<LogoFallback />}>
              <ThreeLogoCanvas />
            </Suspense>
          ) : (
            <LogoFallback />
          )}
        </div>

        {/* Rodapé: Pilares de Gestão Corporativa */}
        <div className="relative z-10 space-y-4">
          <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wider text-primary">
            <Sparkles className="size-3.5" />
            <span>Controle Patrimonial de Alto Desempenho</span>
          </div>
          <h2 className="text-2xl font-bold tracking-tight text-foreground xl:text-3xl max-w-lg leading-snug">
            Gestão unificada de ativos de TI, vínculos e termos com rastreabilidade total.
          </h2>

          <div className="grid grid-cols-3 gap-3 pt-2 border-t border-border/70">
            <div className="rounded-lg border border-border/60 bg-surface/40 p-3 backdrop-blur-sm">
              <div className="flex items-center gap-1.5 text-[11px] font-medium text-muted-foreground">
                <Server className="size-3.5 text-primary" />
                <span>Rastreio</span>
              </div>
              <p className="mt-1 text-xs font-semibold text-foreground">Serial, MAC & IP</p>
            </div>

            <div className="rounded-lg border border-border/60 bg-surface/40 p-3 backdrop-blur-sm">
              <div className="flex items-center gap-1.5 text-[11px] font-medium text-muted-foreground">
                <FileCheck2 className="size-3.5 text-primary" />
                <span>Termos</span>
              </div>
              <p className="mt-1 text-xs font-semibold text-foreground">Emissão em PDF</p>
            </div>

            <div className="rounded-lg border border-border/60 bg-surface/40 p-3 backdrop-blur-sm">
              <div className="flex items-center gap-1.5 text-[11px] font-medium text-muted-foreground">
                <History className="size-3.5 text-primary" />
                <span>Auditoria</span>
              </div>
              <p className="mt-1 text-xs font-semibold text-foreground">Histórico e Estorno</p>
            </div>
          </div>
        </div>
      </section>

      {/* PAINEL DE ACESSO CORPORATIVO (Formulário Elegante e Acessível) */}
      <section className="relative z-10 flex flex-1 items-center justify-center p-6 sm:p-10 lg:p-12">
        <div className="w-full max-w-[400px] space-y-6">
          {/* Header Mobile com insígnia da marca */}
          <div className="lg:hidden flex flex-col items-center text-center mb-6">
            <div className="flex size-12 items-center justify-center rounded-xl border border-primary/30 bg-primary/10 mb-3 shadow-md">
              <img
                src="/logo-revalle.jpg"
                alt="Revalle"
                className="size-9 rounded-lg object-cover"
              />
            </div>
            <h1 className="text-xl font-bold tracking-tight text-foreground">
              Controle de Estoque de TI
            </h1>
            <p className="text-xs text-muted-foreground mt-0.5">
              Portal corporativo de controle de ativos
            </p>
          </div>

          {/* Card principal do formulário com elevação suave */}
          <div className="rounded-2xl border border-border/80 bg-surface/85 p-6 sm:p-8 backdrop-blur-xl shadow-xl transition-all duration-300 hover:border-primary/40">
            <div>
              <div className="flex items-center gap-2">
                <span className="size-2 rounded-full bg-primary" />
                <p className="text-[11px] font-bold uppercase tracking-wider text-muted-foreground">
                  Acesso Restrito
                </p>
              </div>
              <h2 className="mt-2 text-2xl font-bold tracking-tight text-foreground">
                Entrar no Sistema
              </h2>
              <p className="mt-1 text-xs text-muted-foreground leading-relaxed">
                Utilize suas credenciais corporativas autorizadas para acessar a plataforma.
              </p>
            </div>

            <form onSubmit={handleSubmit} className="mt-6 space-y-4" noValidate>
              <div className="space-y-1.5">
                <Label htmlFor="username" className="text-xs font-medium text-foreground">
                  Identificação do Usuário
                </Label>
                <div className="relative transition-all duration-200 focus-within:ring-2 focus-within:ring-primary/20 focus-within:border-primary/70 rounded-md">
                  <User
                    className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground"
                    aria-hidden
                  />
                  <Input
                    id="username"
                    name="username"
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    placeholder="Digite seu usuário corporativo"
                    autoComplete="username"
                    className="h-11 pl-9 bg-background/60 border-input text-sm"
                    maxLength={64}
                    autoFocus
                    required
                  />
                </div>
              </div>

              <div className="space-y-1.5">
                <div className="flex items-center justify-between">
                  <Label htmlFor="password" className="text-xs font-medium text-foreground">
                    Senha de Acesso
                  </Label>
                </div>
                <div className="relative transition-all duration-200 focus-within:ring-2 focus-within:ring-primary/20 focus-within:border-primary/70 rounded-md">
                  <Lock
                    className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground"
                    aria-hidden
                  />
                  <Input
                    id="password"
                    name="password"
                    type={showPass ? "text" : "password"}
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    placeholder="Digite sua senha de acesso"
                    autoComplete="current-password"
                    className="h-11 pl-9 pr-11 bg-background/60 border-input text-sm"
                    maxLength={128}
                    required
                  />
                  <button
                    type="button"
                    onClick={() => setShowPass((s) => !s)}
                    aria-label={showPass ? "Ocultar senha" : "Exibir senha"}
                    className="absolute right-1.5 top-1/2 grid size-8 -translate-y-1/2 place-items-center rounded-md text-muted-foreground transition-colors hover:bg-accent hover:text-foreground"
                  >
                    {showPass ? <EyeOff className="size-4" /> : <Eye className="size-4" />}
                  </button>
                </div>
              </div>

              {error && (
                <div
                  role="alert"
                  className="rounded-lg border border-destructive/30 bg-destructive/10 px-3.5 py-2.5 text-xs font-medium text-destructive animate-in fade-in-50 duration-200"
                >
                  {error}
                </div>
              )}

              <Button
                type="submit"
                size="lg"
                className="mt-2 h-11 w-full font-semibold transition-all duration-200 shadow-md hover:shadow-primary/20 cursor-pointer"
                disabled={submitting}
              >
                {submitting ? (
                  <span className="flex items-center gap-2">
                    <span className="size-4 rounded-full border-2 border-primary-foreground border-t-transparent animate-spin" />
                    <span>Autenticando…</span>
                  </span>
                ) : (
                  "Entrar na Plataforma"
                )}
              </Button>
            </form>

            <div className="mt-6 pt-5 border-t border-border/70 flex items-center justify-center gap-1.5 text-[11px] text-muted-foreground">
              <ShieldCheck className="size-3.5 text-emerald-500 shrink-0" aria-hidden />
              <span>Autenticação criptografada por token JWT</span>
            </div>
          </div>

          <p className="text-center text-[11px] text-muted-foreground/80">
            Revalle Controle de Patrimônio • TI & Operações
          </p>
        </div>
      </section>
    </div>
  );
}
