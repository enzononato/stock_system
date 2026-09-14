import { useEffect, useState, lazy, Suspense } from 'react'
import { useNavigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { LogoFallback } from '@/components/auth/LogoFallback'
import { Eye, EyeOff, Lock, ShieldCheck, User, Sparkles, Loader2 } from 'lucide-react'

// Carregamento sob demanda do WebGL three.js (isolado do bundle principal) —
// só é importado quando o painel institucional de fato precisa montá-lo.
const ThreeLogoCanvas = lazy(() => import('@/components/auth/ThreeLogoCanvas'))

export default function LoginPage() {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [showPass, setShowPass] = useState(false)
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)
  const [isDesktop, setIsDesktop] = useState(false)
  const { login } = useAuth()
  const navigate = useNavigate()

  useEffect(() => {
    // Detecta viewport desktop para renderização seletiva do 3D — mesmo
    // abaixo de lg o painel institucional fica com `hidden`, mas sem esta
    // checagem o React ainda inicializaria o WebGLRenderer escondido.
    const checkDesktop = () => setIsDesktop(window.innerWidth >= 1024)
    checkDesktop()
    window.addEventListener('resize', checkDesktop)
    return () => window.removeEventListener('resize', checkDesktop)
  }, [])

  useEffect(() => {
    // Bloqueia qualquer possibilidade de scroll no body durante a rota de login
    const originalOverflow = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    return () => {
      document.body.style.overflow = originalOverflow
    }
  }, [])

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setError('')
    setLoading(true)
    try {
      await login(username, password)
      navigate('/', { replace: true })
    } catch {
      setError('Usuário ou senha inválidos.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="relative h-screen h-[100dvh] max-h-[100dvh] w-full max-w-[100vw] bg-background overflow-hidden flex flex-col lg:grid lg:grid-cols-[1.2fr_minmax(420px,520px)] text-foreground">
      {/* Malha técnica e glows radiais de fundo (Enterprise Core) */}
      <div className="pointer-events-none absolute inset-0 -z-10 bg-grid-tech opacity-40" aria-hidden />
      <div
        className="pointer-events-none absolute -top-40 -left-40 size-[600px] rounded-full bg-[var(--login-wash)] blur-[120px] -z-10"
        aria-hidden
      />
      <div
        className="pointer-events-none absolute bottom-0 right-1/3 size-[500px] rounded-full bg-[var(--login-wash-soft)] blur-[100px] -z-10"
        aria-hidden
      />

      {/* PAINEL INSTITUCIONAL — some abaixo de lg, só o formulário aparece */}
      <section className="relative hidden lg:flex flex-col justify-between p-8 xl:p-12 border-r border-[var(--login-divider)] bg-gradient-to-br from-sidebar via-background to-sidebar overflow-hidden h-full max-h-[100dvh]">
        {/* Topo institucional */}
        <div className="relative z-10 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="flex size-9 items-center justify-center rounded-lg border border-[var(--login-border-strong)] bg-[var(--login-wash)] shadow-sm">
              <img src="/logo-revalle.jpg" alt="Revalle" className="size-7 rounded object-cover" />
            </div>
            <div>
              <span className="text-sm font-bold tracking-tight text-foreground">Revalle</span>
              <span className="ml-2 rounded border border-[var(--login-border)] bg-[var(--login-wash)] px-1.5 py-0.5 text-[10px] font-semibold text-primary uppercase">
                Enterprise Core
              </span>
            </div>
          </div>

          <div className="flex items-center gap-2 rounded-full border border-[var(--login-divider)] bg-[var(--login-glass)] px-3 py-1 text-xs text-muted-foreground backdrop-blur-md">
            <span className="size-2 rounded-full bg-[var(--status-available)] animate-pulse" />
            <span>Sistema Operacional</span>
          </div>
        </div>

        {/* Centro: logo 3D interativo ou fallback vetorial */}
        <div className="relative z-10 my-auto flex flex-1 min-h-[260px] max-h-[460px] w-full items-center justify-center">
          {/* Anéis decorativos concêntricos */}
          <div className="pointer-events-none absolute inset-0 flex items-center justify-center -z-10">
            <div className="size-[420px] rounded-full border border-[var(--login-wash)] opacity-60" />
            <div className="size-[520px] rounded-full border border-dashed border-[var(--login-border)] opacity-40" />
          </div>

          {isDesktop ? (
            <Suspense fallback={<LogoFallback />}>
              <ThreeLogoCanvas />
            </Suspense>
          ) : (
            <LogoFallback />
          )}
        </div>

        {/* Rodapé institucional */}
        <div className="relative z-10 space-y-3 shrink-0">
          <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wider text-primary">
            <Sparkles className="size-3.5" />
            <span>Controle Patrimonial de Alto Desempenho</span>
          </div>
          <h2 className="text-2xl font-bold tracking-tight text-foreground xl:text-3xl max-w-lg leading-snug">
            Gestão unificada de ativos de TI, vínculos e termos com rastreabilidade total.
          </h2>
        </div>
      </section>

      {/* PAINEL DE ACESSO — formulário */}
      <section className="relative z-10 flex flex-1 items-center justify-center p-6 sm:p-8 lg:p-10 overflow-hidden h-full max-h-[100dvh]">
        <div className="w-full max-w-[400px] space-y-5 my-auto shrink-0">
          {/* Identidade compacta — some a partir de lg (o painel institucional
              ao lado já mostra a marca); abaixo de lg é o painel que some, e
              sem isto a tela ficaria sem nenhuma identidade visual. */}
          <div className="lg:hidden flex flex-col items-center text-center mb-6">
            <div className="flex size-12 items-center justify-center rounded-xl border border-[var(--login-border-strong)] bg-[var(--login-wash)] mb-3 shadow-md">
              <img src="/logo-revalle.jpg" alt="Revalle" className="size-9 rounded-lg object-cover" />
            </div>
            <h1 className="text-xl font-bold tracking-tight text-foreground">Controle de Estoque</h1>
          </div>

          {/* Card do formulário */}
          <div className="rounded-2xl border border-[var(--login-divider)] bg-[var(--login-glass)] p-6 sm:p-8 backdrop-blur-xl shadow-xl transition-all duration-300 hover:border-[var(--login-border-strong)]">
            <div>
              <div className="flex items-center gap-2">
                <span className="size-2 rounded-full bg-primary" />
                <p className="text-[11px] font-bold uppercase tracking-wider text-muted-foreground">
                  Acesso Restrito
                </p>
              </div>
              <h2 className="mt-2 text-2xl font-bold tracking-tight text-foreground">Entrar no Sistema</h2>
              <p className="mt-1 text-xs text-muted-foreground leading-relaxed">
                Utilize suas credenciais corporativas autorizadas para acessar a plataforma.
              </p>
            </div>

            <form onSubmit={handleSubmit} className="mt-6 space-y-4" noValidate>
              <div className="space-y-1.5">
                <Label htmlFor="username" className="text-xs font-medium text-foreground">
                  Usuário
                </Label>
                <div className="relative transition-all duration-200 focus-within:ring-2 focus-within:ring-[var(--login-wash)] focus-within:border-[var(--login-border-strong)] rounded-md">
                  <User
                    size={18}
                    className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
                    aria-hidden
                  />
                  <Input
                    id="username"
                    name="username"
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    placeholder="Digite seu usuário"
                    autoComplete="username"
                    className="h-11 pl-10 bg-[var(--login-glass)] border-input text-sm"
                    autoFocus
                    required
                  />
                </div>
              </div>

              <div className="space-y-1.5">
                <Label htmlFor="password" className="text-xs font-medium text-foreground">
                  Senha
                </Label>
                <div className="relative transition-all duration-200 focus-within:ring-2 focus-within:ring-[var(--login-wash)] focus-within:border-[var(--login-border-strong)] rounded-md">
                  <Lock
                    size={18}
                    className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
                    aria-hidden
                  />
                  <Input
                    id="password"
                    name="password"
                    type={showPass ? 'text' : 'password'}
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    placeholder="Digite sua senha"
                    autoComplete="current-password"
                    className="h-11 pl-10 pr-10 bg-[var(--login-glass)] border-input text-sm"
                    required
                  />
                  <button
                    type="button"
                    onClick={() => setShowPass((s) => !s)}
                    aria-label={showPass ? 'Ocultar senha' : 'Exibir senha'}
                    className="absolute right-1.5 top-1/2 grid size-8 -translate-y-1/2 place-items-center rounded-md text-muted-foreground transition-colors duration-micro hover:bg-surface-alt hover:text-foreground"
                  >
                    {showPass ? <EyeOff size={18} /> : <Eye size={18} />}
                  </button>
                </div>
              </div>

              {error && (
                <div
                  role="alert"
                  className="rounded-lg border border-[var(--login-danger-border)] bg-[var(--login-danger-wash)] px-3.5 py-2.5 text-xs font-medium text-destructive text-center"
                >
                  {error}
                </div>
              )}

              <Button
                type="submit"
                size="lg"
                className="mt-2 h-11 w-full font-semibold shadow-md hover:shadow-lg transition-all duration-200"
                disabled={loading}
              >
                {loading ? (
                  <span className="flex items-center justify-center gap-2">
                    <Loader2 className="h-4 w-4 animate-spin" />
                    <span>Entrando...</span>
                  </span>
                ) : (
                  'Entrar'
                )}
              </Button>
            </form>

            <div className="mt-6 pt-5 border-t border-[var(--login-divider)] flex items-center justify-center gap-1.5 text-[11px] text-muted-foreground">
              <ShieldCheck className="size-3.5 text-[var(--status-available)] shrink-0" aria-hidden />
              <span>Autenticação Segura JWT</span>
            </div>
          </div>

          <p className="text-center text-[11px] text-muted-foreground">
            Revalle Controle de Patrimônio • TI & Operações
          </p>
        </div>
      </section>
    </div>
  )
}
