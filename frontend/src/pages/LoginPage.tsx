import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
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
    <div className="relative min-h-screen flex items-center justify-center bg-slate-950 overflow-hidden p-4 select-none">
      {/* Background Decorative Glow Blobs */}
      <div className="absolute -top-40 -left-40 w-96 h-96 bg-indigo-600/30 rounded-full blur-[120px] pointer-events-none" />
      <div className="absolute -bottom-40 -right-40 w-96 h-96 bg-violet-600/25 rounded-full blur-[120px] pointer-events-none" />

      <div className="relative w-full max-w-md animate-scale-in">
        <div className="rounded-3xl bg-white/95 backdrop-blur-xl shadow-2xl p-8 sm:p-10 border border-white/20">
          {/* Logo & Header */}
          <div className="flex flex-col items-center mb-8 text-center">
            <div className="h-16 w-16 rounded-2xl bg-gradient-to-tr from-indigo-600 via-indigo-500 to-violet-500 flex items-center justify-center text-white shadow-xl shadow-indigo-500/35 ring-4 ring-indigo-50 mb-4">
              <Boxes size={34} className="stroke-[2.2]" />
            </div>
            <h1 className="text-2xl font-bold tracking-tight text-slate-900 font-heading">
              Controle de Estoque
            </h1>
            <p className="text-sm font-medium text-slate-500 mt-1 flex items-center gap-1">
              <span>Revalle TI</span>
              <span className="h-1 w-1 rounded-full bg-slate-300" />
              <span className="text-indigo-600 font-bold">Portal Corporativo</span>
            </p>
          </div>

          <form onSubmit={handleSubmit} className="flex flex-col gap-5">
            <div className="flex flex-col gap-2">
              <Label htmlFor="username" className="text-xs font-bold uppercase tracking-wider text-slate-600">
                Usuário
              </Label>
              <div className="relative">
                <User size={18} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                <Input
                  id="username"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="Digite seu usuário"
                  className="pl-11 h-11 bg-slate-50/70 border-slate-200 focus-visible:ring-indigo-500/30 rounded-xl font-medium"
                  autoFocus
                  required
                />
              </div>
            </div>

            <div className="flex flex-col gap-2">
              <Label htmlFor="password" className="text-xs font-bold uppercase tracking-wider text-slate-600">
                Senha
              </Label>
              <div className="relative">
                <Lock size={18} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                <Input
                  id="password"
                  type={showPass ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Digite sua senha"
                  className="pl-11 pr-11 h-11 bg-slate-50/70 border-slate-200 focus-visible:ring-indigo-500/30 rounded-xl font-medium"
                  required
                />
                <button
                  type="button"
                  onClick={() => setShowPass(!showPass)}
                  className="absolute right-3.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 rounded-lg hover:bg-slate-100 transition-colors"
                >
                  {showPass ? <EyeOff size={18} /> : <Eye size={18} />}
                </button>
              </div>
            </div>

            {error && (
              <div className="p-3 rounded-xl bg-rose-50 border border-rose-200/80 text-rose-700 text-xs font-semibold text-center animate-fade-in">
                {error}
              </div>
            )}

            <Button
              type="submit"
              variant="gradient"
              size="lg"
              className="mt-2 w-full h-11 rounded-xl text-base shadow-lg shadow-indigo-500/30 font-bold"
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
            </Button>
          </form>

          {/* Footer badge */}
          <div className="mt-8 pt-6 border-t border-slate-100 text-center">
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-slate-100 text-[11px] font-semibold text-slate-500">
              <ShieldCheck size={13} className="text-emerald-600" />
              <span>Autenticação Segura JWT</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}

