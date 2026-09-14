import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Boxes, Eye, EyeOff, Lock, User, ShieldCheck, Loader2 } from 'lucide-react'

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
    <div className="min-h-screen grid lg:grid-cols-2 bg-canvas text-foreground">
      {/* Painel institucional — some abaixo de lg, só o formulário aparece */}
      <div className="hidden lg:flex bg-primary text-primary-foreground flex-col justify-between p-10">
        <div className="flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-md border border-border">
            <Boxes size={22} />
          </div>
          <h1 className="text-heading-lg font-semibold">Controle de Estoque</h1>
        </div>
        <p className="flex items-center gap-1.5 text-body text-primary-foreground">
          <span>Revalle TI</span>
          <span className="h-1 w-1 rounded-full bg-primary-foreground" />
          <span>Portal Corporativo</span>
        </p>
      </div>

      {/* Formulário */}
      <div className="flex items-center justify-center p-6">
        <div className="w-full max-w-sm space-y-5">
          {/* Identidade compacta — some a partir de lg (o painel institucional
              ao lado já mostra a marca); abaixo de lg é o painel que some, e
              sem isto a tela ficaria sem nenhuma identidade visual. */}
          <div className="flex flex-col gap-1.5 lg:hidden">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Acesso Restrito
            </span>
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-md border border-border bg-primary text-primary-foreground">
                <Boxes size={20} />
              </div>
              <h1 className="text-heading-sm font-semibold text-foreground">Controle de Estoque</h1>
            </div>
          </div>

          <form onSubmit={handleSubmit} className="space-y-5">
          <div className="flex flex-col gap-2">
            <Label htmlFor="username">Usuário</Label>
            <div className="relative">
              <User
                size={18}
                className="absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none"
              />
              <Input
                id="username"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="Digite seu usuário"
                className="pl-10"
                autoFocus
                required
              />
            </div>
          </div>

          <div className="flex flex-col gap-2">
            <Label htmlFor="password">Senha</Label>
            <div className="relative">
              <Lock
                size={18}
                className="absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none"
              />
              <Input
                id="password"
                type={showPass ? 'text' : 'password'}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Digite sua senha"
                className="pl-10 pr-10"
                required
              />
              <button
                type="button"
                onClick={() => setShowPass(!showPass)}
                className="absolute right-3 top-1/2 -translate-y-1/2 rounded p-1 text-muted-foreground transition-colors duration-micro hover:bg-surface-alt hover:text-foreground"
              >
                {showPass ? <EyeOff size={18} /> : <Eye size={18} />}
              </button>
            </div>
          </div>

          {error && <p className="text-body-sm text-destructive text-center">{error}</p>}

          <Button type="submit" size="lg" className="w-full" disabled={loading}>
            {loading ? (
              <span className="flex items-center justify-center gap-2">
                <Loader2 className="h-4 w-4 animate-spin" />
                <span>Entrando...</span>
              </span>
            ) : (
              'Entrar'
            )}
          </Button>

          <div className="border-t border-border pt-4 text-center">
            <span className="inline-flex items-center gap-1.5 rounded-sm bg-surface-alt px-3 py-1 text-caption text-muted-foreground">
              <ShieldCheck size={13} />
              <span>Autenticação Segura JWT</span>
            </span>
          </div>
          </form>
        </div>
      </div>
    </div>
  )
}
