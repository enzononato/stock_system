import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import SpecularButton from '@/components/effects/SpecularButton'
import { Boxes, Eye, EyeOff, Lock, User, ShieldCheck } from 'lucide-react'

export default function LoginPage() {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [showPass, setShowPass] = useState(false)
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)
  const { login } = useAuth()
  const navigate = useNavigate()

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
    <div className="relative min-h-screen flex items-center justify-center overflow-hidden bg-background p-4 select-none">
      {/* Static subtle radial glow — replaced the animated Galaxy background (too heavy) */}
      <div
        className="absolute inset-0 pointer-events-none"
        style={{
          background:
            'radial-gradient(ellipse 60% 50% at 50% 20%, rgba(56,189,248,0.10), transparent), radial-gradient(ellipse 50% 40% at 80% 80%, rgba(168,85,247,0.08), transparent)',
        }}
      />
      <div className="absolute inset-0 bg-gradient-to-t from-background via-background/60 to-transparent pointer-events-none" />

      <div className="relative w-full max-w-md">
        <div className="rounded-xl surface shadow-2xl p-8 sm:p-10 backdrop-blur-sm">
          {/* Logo & Header */}
          <div className="flex flex-col items-center mb-8 text-center">
            <div className="h-14 w-14 rounded-lg bg-primary text-primary-foreground flex items-center justify-center shadow-lg mb-4">
              <Boxes size={30} className="stroke-[2.2]" />
            </div>
            <h1 className="text-xl font-semibold tracking-tight text-foreground">
              Controle de Estoque
            </h1>
            <p className="text-sm text-muted-foreground mt-1 flex items-center gap-1.5">
              <span>Revalle TI</span>
              <span className="h-1 w-1 rounded-full bg-muted-foreground/50" />
              <span className="text-primary font-medium">Portal Corporativo</span>
            </p>
          </div>

          <form onSubmit={handleSubmit} className="flex flex-col gap-5">
            <div className="flex flex-col gap-2">
              <Label htmlFor="username" className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                Usuário
              </Label>
              <div className="relative">
                <User size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none" />
                <Input
                  id="username"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="Digite seu usuário"
                  className="pl-10 h-10"
                  autoFocus
                  required
                />
              </div>
            </div>

            <div className="flex flex-col gap-2">
              <Label htmlFor="password" className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                Senha
              </Label>
              <div className="relative">
                <Lock size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none" />
                <Input
                  id="password"
                  type={showPass ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Digite sua senha"
                  className="pl-10 pr-10 h-10"
                  required
                />
                <button
                  type="button"
                  onClick={() => setShowPass(!showPass)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground p-1 rounded-md hover:bg-accent transition-colors"
                >
                  {showPass ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </div>

            {error && (
              <div className="p-3 rounded-md bg-destructive/10 border border-destructive/30 text-destructive text-xs font-medium text-center">
                {error}
              </div>
            )}

            <SpecularButton
              type="submit"
              size="lg"
              radius={8}
              lineColor="#38bdf8"
              baseColor="#0b1020"
              className="mt-1 w-full h-10"
              disabled={loading}
            >
              {loading ? (
                <div className="flex items-center justify-center gap-2">
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-white/30 border-t-white" />
                  <span>Entrando...</span>
                </div>
              ) : (
                'Entrar'
              )}
            </SpecularButton>
          </form>

          {/* Footer badge */}
          <div className="mt-8 pt-6 border-t border-border text-center">
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-muted text-[11px] font-medium text-muted-foreground">
              <ShieldCheck size={13} className="text-success" />
              <span>Autenticação Segura JWT</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
