import { useEffect, useState, type FormEvent } from "react";
import { useNavigate } from "@tanstack/react-router";
import { Eye, EyeOff, Lock, ShieldCheck, User } from "lucide-react";

const logoUrl = "/logo-revalle.jpg";
import { useAuth } from "@/lib/auth";
import { getErrorMessage } from "@/lib/api-error";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";

export function LoginPage() {
  const { login, user, isLoading } = useAuth();
  const navigate = useNavigate();

  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [error, setError] = useState("");
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    if (!isLoading && user) void navigate({ to: "/", replace: true });
  }, [isLoading, user, navigate]);

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setError("");
    setSubmitting(true);
    try {
      await login(username, password);
      void navigate({ to: "/", replace: true });
    } catch (err) {
      setError(getErrorMessage(err, "Usuário ou senha inválidos."));
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="relative grid min-h-screen bg-background lg:grid-cols-[1.1fr_minmax(0,520px)]">
      {/* Painel institucional */}
      <section className="relative hidden overflow-hidden border-r border-border bg-sidebar lg:flex lg:flex-col lg:justify-between lg:p-12">
        <div className="bg-grid pointer-events-none absolute inset-0 opacity-60" aria-hidden />
        <div className="relative flex items-center gap-3">
          <img src={logoUrl} alt="Revalle" className="size-10 rounded-lg object-cover" />
          <span className="text-sm font-semibold">Controle de Estoque de TI</span>
        </div>

        <div className="relative max-w-lg">
          <p className="text-eyebrow">Portal corporativo</p>
          <h1 className="mt-3 text-4xl font-bold leading-[1.1] tracking-tight text-foreground">
            Patrimônio, empréstimos e periféricos sob controle.
          </h1>
          <p className="mt-4 text-sm leading-relaxed text-muted-foreground">
            Cadastro de ativos, vínculo de periféricos, termos de responsabilidade em PDF,
            histórico auditável com estorno e indicadores por unidade — em um único fluxo.
          </p>
          <dl className="mt-10 grid grid-cols-3 gap-4 border-t border-border pt-6">
            {[
              ["Rastreio", "Serial, MAC e IP"],
              ["Termos", "Geração em PDF"],
              ["Auditoria", "Histórico com estorno"],
            ].map(([k, v]) => (
              <div key={k}>
                <dt className="text-eyebrow">{k}</dt>
                <dd className="mt-1 text-sm font-medium text-foreground">{v}</dd>
              </div>
            ))}
          </dl>
        </div>

        <p className="relative text-xs text-muted-foreground">Uso restrito a colaboradores autorizados.</p>
      </section>

      {/* Formulário */}
      <section className="flex items-center justify-center px-4 py-10 sm:px-8">
        <div className="w-full max-w-sm">
          <div className="mb-8 lg:hidden">
            <img src={logoUrl} alt="Revalle" className="mb-4 size-10 rounded-lg object-cover" />
            <h1 className="text-2xl font-bold tracking-tight">Controle de Estoque de TI</h1>
          </div>

          <h2 className="text-xl font-semibold tracking-tight">Entrar na sua conta</h2>
          <p className="mt-1 text-sm text-muted-foreground">Use as credenciais fornecidas pela equipe de TI.</p>

          <form onSubmit={handleSubmit} className="mt-8 flex flex-col gap-5" noValidate>
            <div className="flex flex-col gap-2">
              <Label htmlFor="username">Usuário</Label>
              <div className="relative">
                <User
                  className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground"
                  aria-hidden
                />
                <Input
                  id="username"
                  name="username"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="Digite seu usuário"
                  autoComplete="username"
                  className="h-11 pl-9"
                  maxLength={64}
                  autoFocus
                  required
                />
              </div>
            </div>

            <div className="flex flex-col gap-2">
              <Label htmlFor="password">Senha</Label>
              <div className="relative">
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
                  placeholder="Digite sua senha"
                  autoComplete="current-password"
                  className="h-11 pl-9 pr-11"
                  maxLength={128}
                  required
                />
                <button
                  type="button"
                  onClick={() => setShowPass((s) => !s)}
                  aria-label={showPass ? "Ocultar senha" : "Mostrar senha"}
                  className="absolute right-1.5 top-1/2 grid size-8 -translate-y-1/2 place-items-center rounded-md text-muted-foreground transition-colors hover:bg-accent hover:text-foreground"
                >
                  {showPass ? <EyeOff className="size-4" /> : <Eye className="size-4" />}
                </button>
              </div>
            </div>

            {error && (
              <p
                role="alert"
                className="rounded-md border border-destructive/30 bg-destructive/10 px-3 py-2 text-xs font-medium text-destructive"
              >
                {error}
              </p>
            )}

            <Button type="submit" size="lg" className="mt-1 h-11 w-full" disabled={submitting}>
              {submitting ? "Entrando…" : "Entrar"}
            </Button>
          </form>

          <p className="mt-8 flex items-center justify-center gap-1.5 text-xs text-muted-foreground">
            <ShieldCheck className="size-3.5 text-success" aria-hidden />
            Autenticação segura com token JWT
          </p>
        </div>
      </section>
    </div>
  );
}
