import { useRef, useState } from 'react'
import { Link, useNavigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import { confirmarCadastroMfa, iniciarCadastroMfa, verificarMfa } from '@/api/auth'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { toast } from '@/components/ui/toast'
import AuthShell from '@/components/auth/AuthShell'
import { CadastroAutenticador } from '@/components/auth/CadastroAutenticador'
import { CodigosRecuperacao } from '@/components/auth/CodigosRecuperacao'
import { mensagemErroAuth } from '@/lib/auth-errors'
import { Eye, EyeOff, Lock, User, KeyRound } from 'lucide-react'

type Etapa =
  | { tipo: 'senha' }
  | { tipo: 'mfa'; mfaToken: string }
  | { tipo: 'cadastro'; mfaToken: string }
  | { tipo: 'codigos'; codigos: string[]; accessToken: string }

const AVISO_CODIGOS_RESTANTES = 3

const CABECALHOS: Record<Etapa['tipo'], { titulo: string; descricao?: string }> = {
  senha: {
    titulo: 'Entrar no Sistema',
    descricao: 'Utilize suas credenciais corporativas autorizadas para acessar a plataforma.',
  },
  mfa: {
    titulo: 'Verificação em duas etapas',
    descricao: 'Digite o código de 6 dígitos do aplicativo autenticador ou um código de recuperação.',
  },
  cadastro: {
    titulo: 'Cadastre o autenticador',
    descricao:
      'Gestores precisam da verificação em duas etapas. Instale um aplicativo autenticador (Google Authenticator, Microsoft Authenticator ou similar), leia o QR code e digite o código gerado.',
  },
  codigos: { titulo: 'Códigos de recuperação' },
}

function MensagemErro({ texto }: { texto: string }) {
  return (
    <div
      role="alert"
      className="rounded-lg border border-[var(--login-danger-border)] bg-[var(--login-danger-wash)] px-3.5 py-2.5 text-xs font-medium text-destructive text-center"
    >
      {texto}
    </div>
  )
}

export default function LoginPage() {
  const [etapa, setEtapa] = useState<Etapa>({ tipo: 'senha' })
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const { login, concluirLogin, avisoLogin, limparAviso } = useAuth()
  const navigate = useNavigate()
  const accessTokenDoCadastro = useRef('')

  function voltarParaSenha() {
    setPassword('')
    setEtapa({ tipo: 'senha' })
  }

  const cabecalho = CABECALHOS[etapa.tipo]

  return (
    <AuthShell titulo={cabecalho.titulo} descricao={cabecalho.descricao}>
      {etapa.tipo === 'senha' && (
        <SenhaEtapa
          aviso={avisoLogin}
          username={username}
          password={password}
          onUsername={setUsername}
          onPassword={setPassword}
          onEntrar={async () => {
            limparAviso()
            const resposta = await login(username, password)
            if (resposta.etapa === 'ok') {
              navigate('/', { replace: true })
            } else if (resposta.mfa_token) {
              setPassword('')
              setEtapa({
                tipo: resposta.etapa === 'mfa' ? 'mfa' : 'cadastro',
                mfaToken: resposta.mfa_token,
              })
            }
          }}
        />
      )}

      {etapa.tipo === 'mfa' && (
        <MfaEtapa
          onVoltar={voltarParaSenha}
          onVerificar={async (codigo) => {
            const resposta = await verificarMfa(etapa.mfaToken, codigo)
            if (!resposta.access_token) throw new Error('Resposta de verificação sem token.')
            await concluirLogin(resposta.access_token)
            const restantes = resposta.codigos_recuperacao_restantes
            if (restantes != null && restantes <= AVISO_CODIGOS_RESTANTES) {
              toast(`Poucos códigos de recuperação restantes (${restantes}). Gere novos em Meu perfil.`)
            }
            navigate('/', { replace: true })
          }}
        />
      )}

      {etapa.tipo === 'cadastro' && (
        <CadastroAutenticador
          iniciar={() => iniciarCadastroMfa(etapa.mfaToken)}
          confirmar={async (codigo) => {
            const resposta = await confirmarCadastroMfa(etapa.mfaToken, codigo)
            accessTokenDoCadastro.current = resposta.access_token ?? ''
            return resposta.codigos_recuperacao ?? []
          }}
          onConcluido={(codigos) =>
            setEtapa({ tipo: 'codigos', codigos, accessToken: accessTokenDoCadastro.current })
          }
        />
      )}

      {etapa.tipo === 'codigos' && (
        <CodigosEtapa
          codigos={etapa.codigos}
          onEntrar={async () => {
            await concluirLogin(etapa.accessToken)
            navigate('/', { replace: true })
          }}
        />
      )}
    </AuthShell>
  )
}

interface SenhaEtapaProps {
  aviso: string | null
  username: string
  password: string
  onUsername: (v: string) => void
  onPassword: (v: string) => void
  onEntrar: () => Promise<void>
}

function SenhaEtapa({ aviso, username, password, onUsername, onPassword, onEntrar }: SenhaEtapaProps) {
  const [showPass, setShowPass] = useState(false)
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setError('')
    setLoading(true)
    try {
      await onEntrar()
    } catch (err) {
      setError(mensagemErroAuth(err, 'Usuário ou senha inválidos.'))
    } finally {
      setLoading(false)
    }
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-4" noValidate>
      {aviso && (
        <p
          role="status"
          className="rounded-lg border border-border bg-surface-alt px-3.5 py-2.5 text-xs font-medium text-foreground leading-relaxed"
        >
          {aviso}
        </p>
      )}
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
            onChange={(e) => onUsername(e.target.value)}
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
            onChange={(e) => onPassword(e.target.value)}
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
        <div className="flex justify-end">
          <Link
            to="/esqueci-senha"
            className="text-xs font-medium text-brand-text underline-offset-4 hover:underline"
          >
            Esqueci minha senha
          </Link>
        </div>
      </div>

      {error && <MensagemErro texto={error} />}

      <Button
        type="submit"
        size="lg"
        className="mt-2 h-11 w-full font-semibold shadow-md hover:shadow-lg transition-all duration-200"
        loading={loading}
      >
        {loading ? 'Entrando...' : 'Entrar'}
      </Button>
    </form>
  )
}

interface MfaEtapaProps {
  onVerificar: (codigo: string) => Promise<void>
  onVoltar: () => void
}

function MfaEtapa({ onVerificar, onVoltar }: MfaEtapaProps) {
  const [codigo, setCodigo] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setError('')
    setLoading(true)
    try {
      await onVerificar(codigo.trim())
    } catch (err) {
      setError(mensagemErroAuth(err, 'Não foi possível verificar o código.'))
      setLoading(false)
    }
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-4" noValidate>
      <div className="space-y-1.5">
        <Label htmlFor="codigo-mfa" className="text-xs font-medium text-foreground">
          Código
        </Label>
        <div className="relative transition-all duration-200 focus-within:ring-2 focus-within:ring-[var(--login-wash)] focus-within:border-[var(--login-border-strong)] rounded-md">
          <KeyRound
            size={18}
            className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
            aria-hidden
          />
          <Input
            id="codigo-mfa"
            name="codigo"
            value={codigo}
            onChange={(e) => setCodigo(e.target.value)}
            placeholder="000000 ou código de recuperação"
            inputMode="text"
            autoComplete="one-time-code"
            className="h-11 pl-10 bg-[var(--login-glass)] border-input text-sm"
            autoFocus
            required
          />
        </div>
      </div>

      {error && <MensagemErro texto={error} />}

      <Button
        type="submit"
        size="lg"
        className="h-11 w-full font-semibold"
        loading={loading}
        disabled={!codigo.trim()}
      >
        Verificar
      </Button>
      <Button type="button" variant="ghost" className="w-full" onClick={onVoltar} disabled={loading}>
        Voltar
      </Button>
    </form>
  )
}

function CodigosEtapa({ codigos, onEntrar }: { codigos: string[]; onEntrar: () => Promise<void> }) {
  const [error, setError] = useState('')

  return (
    <div className="space-y-4">
      <CodigosRecuperacao
        codigos={codigos}
        rotuloContinuar="Entrar no sistema"
        onContinuar={() => {
          setError('')
          onEntrar().catch((err: unknown) =>
            setError(mensagemErroAuth(err, 'Não foi possível entrar. Tente novamente.'))
          )
        }}
      />
      {error && <MensagemErro texto={error} />}
    </div>
  )
}
