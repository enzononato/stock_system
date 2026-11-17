import { useEffect, useRef, useState } from 'react'
import { Link, useNavigate } from 'react-router-dom'
import { useAuth } from '@/contexts/AuthContext'
import {
  confirmarCadastroEmail,
  confirmarCadastroMfa,
  iniciarCadastroEmail,
  iniciarCadastroMfa,
  verificarMfa,
  type LoginResponse,
} from '@/api/auth'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { toast } from '@/components/ui/toast'
import AuthShell from '@/components/auth/AuthShell'
import { CadastroAutenticador } from '@/components/auth/CadastroAutenticador'
import { CodigosRecuperacao } from '@/components/auth/CodigosRecuperacao'
import { mensagemErroAuth, mensagemEtapaEncerrada } from '@/lib/auth-errors'
import { Eye, EyeOff, Lock, Mail, KeyRound } from 'lucide-react'

type Etapa =
  | { tipo: 'senha' }
  | { tipo: 'mfa'; mfaToken: string }
  | { tipo: 'cadastro'; mfaToken: string }
  | { tipo: 'email'; mfaToken: string; emailSugerido: string | null }
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
  email: {
    titulo: 'Confirme seu e-mail',
    descricao:
      'Agora o acesso é pelo e-mail corporativo. Informe o seu e-mail @revalle.com.br e digite o código que enviaremos para ele.',
  },
  codigos: { titulo: 'Códigos de recuperação' },
}

const SEGUNDOS_PARA_REENVIO = 60
const MSG_NAO_FOI_POSSIVEL_CONTINUAR = 'Não foi possível continuar. Entre novamente.'

/** Código de confirmação já enviado nesta tela (para não pedir outro ao voltar). */
interface EnvioDeCodigo {
  email: string
  mensagem: string
  enviadoEm: number
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
  // Aviso desta tela (ex.: a 2ª etapa foi encerrada por excesso de tentativas).
  const [avisoLocal, setAvisoLocal] = useState<string | null>(null)

  function voltarParaSenha(aviso: string | null = null) {
    setPassword('')
    setAvisoLocal(aviso)
    setEtapa({ tipo: 'senha' })
  }

  // Código de e-mail já enviado nesta tela, por conta (o que foi digitado no login): quem
  // volta e entra de novo dentro de 60 s digita o código que já recebeu.
  const codigosEnviados = useRef(new Map<string, EnvioDeCodigo>())
  const chaveDoLogin = username.trim().toLowerCase()

  /** Leva à etapa pedida pelo servidor (depois da senha ou da confirmação do e-mail).
   * Resposta sem o token da etapa: volta à senha com o aviso (nada fica girando). */
  function irParaEtapa(resposta: LoginResponse) {
    const mfaToken = resposta.mfa_token
    setPassword('')
    if (mfaToken && resposta.etapa === 'mfa') setEtapa({ tipo: 'mfa', mfaToken })
    else if (mfaToken && resposta.etapa === 'cadastro_2fa') setEtapa({ tipo: 'cadastro', mfaToken })
    else if (mfaToken && resposta.etapa === 'cadastro_email') {
      setEtapa({ tipo: 'email', mfaToken, emailSugerido: resposta.email_sugerido ?? null })
    } else {
      voltarParaSenha(MSG_NAO_FOI_POSSIVEL_CONTINUAR)
    }
  }

  /** Erro que encerra a etapa (token esgotado ou expirado): volta à senha com a explicação. */
  function encerrarSeEtapaAcabou(err: unknown) {
    const encerrada = mensagemEtapaEncerrada(err)
    if (!encerrada) throw err
    voltarParaSenha(encerrada)
  }

  const cabecalho = CABECALHOS[etapa.tipo]

  return (
    <AuthShell titulo={cabecalho.titulo} descricao={cabecalho.descricao}>
      {etapa.tipo === 'senha' && (
        <SenhaEtapa
          aviso={avisoLocal ?? avisoLogin}
          username={username}
          password={password}
          onUsername={setUsername}
          onPassword={setPassword}
          onEntrar={async () => {
            limparAviso()
            setAvisoLocal(null)
            const resposta = await login(username, password)
            if (resposta.etapa === 'ok') {
              navigate('/', { replace: true })
            } else {
              irParaEtapa(resposta)
            }
          }}
        />
      )}

      {etapa.tipo === 'email' && (
        <EmailEtapa
          key={etapa.mfaToken}
          emailSugerido={etapa.emailSugerido}
          envioAnterior={codigosEnviados.current.get(chaveDoLogin) ?? null}
          onVoltar={() => voltarParaSenha()}
          onEnviar={async (email) => {
            try {
              const { detail } = await iniciarCadastroEmail(etapa.mfaToken, email)
              codigosEnviados.current.set(chaveDoLogin, { email, mensagem: detail, enviadoEm: Date.now() })
              return detail
            } catch (err) {
              encerrarSeEtapaAcabou(err)
              return null
            }
          }}
          onConfirmar={async (codigo) => {
            let resposta
            try {
              resposta = await confirmarCadastroEmail(etapa.mfaToken, codigo)
            } catch (err) {
              encerrarSeEtapaAcabou(err)
              return
            }
            if (resposta.etapa === 'ok') {
              if (!resposta.access_token) throw new Error('Resposta de confirmação sem token.')
              await concluirLogin(resposta.access_token)
              navigate('/', { replace: true })
            } else {
              irParaEtapa(resposta)
            }
          }}
        />
      )}

      {etapa.tipo === 'mfa' && (
        <MfaEtapa
          onVoltar={() => voltarParaSenha()}
          onVerificar={async (codigo) => {
            let resposta
            try {
              resposta = await verificarMfa(etapa.mfaToken, codigo)
            } catch (err) {
              // Token esgotado (5 códigos errados) ou expirado: não adianta tentar de novo nesta
              // tela; volta à senha e explica.
              const encerrada = mensagemEtapaEncerrada(err)
              if (!encerrada) throw err
              voltarParaSenha(encerrada)
              return
            }
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
        <Label htmlFor="email" className="text-xs font-medium text-foreground">
          E-mail
        </Label>
        <div className="relative transition-all duration-200 focus-within:ring-2 focus-within:ring-[var(--login-wash)] focus-within:border-[var(--login-border-strong)] rounded-md">
          <Mail
            size={18}
            className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
            aria-hidden
          />
          {/* type="text": quem ainda não migrou entra uma última vez com o usuário antigo. */}
          <Input
            id="email"
            name="email"
            value={username}
            onChange={(e) => onUsername(e.target.value)}
            placeholder="nome@revalle.com.br"
            autoComplete="username"
            aria-describedby="email-dica"
            className="h-11 pl-10 bg-[var(--login-glass)] border-input text-sm"
            autoFocus
            required
          />
        </div>
        <p id="email-dica" className="text-[11px] text-muted-foreground leading-relaxed">
          Primeiro acesso depois da mudança? Entre uma vez com o seu usuário antigo.
        </p>
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

interface EmailEtapaProps {
  emailSugerido: string | null
  /** Código já enviado nesta tela para esta conta (a pessoa voltou e entrou de novo). */
  envioAnterior: EnvioDeCodigo | null
  /** Envia o código; devolve a mensagem do servidor, ou null quando a etapa foi encerrada. */
  onEnviar: (email: string) => Promise<string | null>
  onConfirmar: (codigo: string) => Promise<void>
  onVoltar: () => void
}

/** O código vale 15 minutos no servidor; depois disso não adianta oferecer o campo direto. */
const VALIDADE_DO_CODIGO_MS = 15 * 60 * 1000

function EmailEtapa({ emailSugerido, envioAnterior, onEnviar, onConfirmar, onVoltar }: EmailEtapaProps) {
  // Código enviado há pouco nesta tela: já mostra o campo do código, e o reenvio continua
  // preso à contagem dos 60 s contada desde aquele envio.
  const [anterior] = useState(() => {
    if (!envioAnterior) return null
    const decorrido = Date.now() - envioAnterior.enviadoEm
    return decorrido < VALIDADE_DO_CODIGO_MS ? { ...envioAnterior, decorrido } : null
  })
  const [email, setEmail] = useState(anterior?.email ?? emailSugerido ?? '')
  const [codigo, setCodigo] = useState('')
  const [enviado, setEnviado] = useState(anterior !== null)
  const [aviso, setAviso] = useState(anterior?.mensagem ?? '')
  const [error, setError] = useState('')
  const [enviando, setEnviando] = useState(false)
  const [confirmando, setConfirmando] = useState(false)
  const [espera, setEspera] = useState(() =>
    anterior ? Math.max(0, SEGUNDOS_PARA_REENVIO - Math.floor(anterior.decorrido / 1000)) : 0
  )

  // Contagem regressiva do "Reenviar código" (o servidor só aceita outro pedido depois de 60 s).
  useEffect(() => {
    if (espera <= 0) return
    const id = window.setTimeout(() => setEspera((s) => s - 1), 1000)
    return () => window.clearTimeout(id)
  }, [espera])

  async function enviar() {
    setError('')
    setEnviando(true)
    try {
      const mensagem = await onEnviar(email.trim())
      if (mensagem === null) return // etapa encerrada: a tela já voltou para a senha
      setAviso(mensagem)
      setEnviado(true)
      setCodigo('')
      setEspera(SEGUNDOS_PARA_REENVIO)
    } catch (err) {
      setError(mensagemErroAuth(err, 'Não foi possível enviar o código.'))
    } finally {
      setEnviando(false)
    }
  }

  async function confirmar(e: React.FormEvent) {
    e.preventDefault()
    setError('')
    setConfirmando(true)
    try {
      await onConfirmar(codigo.trim())
    } catch (err) {
      setError(mensagemErroAuth(err, 'Não foi possível confirmar o código.'))
      setConfirmando(false)
    }
  }

  return (
    <form
      onSubmit={(e) => {
        if (enviado) return void confirmar(e)
        e.preventDefault()
        void enviar()
      }}
      className="space-y-4"
      noValidate
    >
      <div className="space-y-1.5">
        <Label htmlFor="email-confirmacao" className="text-xs font-medium text-foreground">
          E-mail
        </Label>
        <div className="relative transition-all duration-200 focus-within:ring-2 focus-within:ring-[var(--login-wash)] focus-within:border-[var(--login-border-strong)] rounded-md">
          <Mail
            size={18}
            className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
            aria-hidden
          />
          <Input
            id="email-confirmacao"
            name="email"
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="nome@revalle.com.br"
            autoComplete="email"
            className="h-11 pl-10 bg-[var(--login-glass)] border-input text-sm"
            autoFocus={!enviado}
            required
          />
        </div>
      </div>

      {aviso && (
        <p
          role="status"
          className="rounded-lg border border-border bg-surface-alt px-3.5 py-2.5 text-xs font-medium text-foreground leading-relaxed"
        >
          {aviso}
        </p>
      )}

      {enviado && (
        <div className="space-y-1.5">
          <Label htmlFor="codigo-email" className="text-xs font-medium text-foreground">
            Código
          </Label>
          <div className="relative transition-all duration-200 focus-within:ring-2 focus-within:ring-[var(--login-wash)] focus-within:border-[var(--login-border-strong)] rounded-md">
            <KeyRound
              size={18}
              className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
              aria-hidden
            />
            <Input
              id="codigo-email"
              name="codigo"
              value={codigo}
              onChange={(e) => setCodigo(e.target.value.replace(/\D/g, '').slice(0, 6))}
              placeholder="000000"
              inputMode="numeric"
              autoComplete="one-time-code"
              maxLength={6}
              className="h-11 pl-10 bg-[var(--login-glass)] border-input text-sm tracking-[0.3em]"
              autoFocus
              required
            />
          </div>
        </div>
      )}

      {error && <MensagemErro texto={error} />}

      {enviado ? (
        <>
          <Button
            type="submit"
            size="lg"
            className="h-11 w-full font-semibold"
            loading={confirmando}
            disabled={codigo.length !== 6}
          >
            Confirmar
          </Button>
          <Button
            type="button"
            variant="outline"
            className="w-full"
            onClick={() => void enviar()}
            loading={enviando}
            disabled={espera > 0 || confirmando || !email.trim()}
          >
            {espera > 0 ? `Reenviar código (${espera} s)` : 'Reenviar código'}
          </Button>
        </>
      ) : (
        <Button
          type="submit"
          size="lg"
          className="h-11 w-full font-semibold"
          loading={enviando}
          disabled={!email.trim()}
        >
          Enviar código
        </Button>
      )}
      <Button
        type="button"
        variant="ghost"
        className="w-full"
        onClick={onVoltar}
        disabled={enviando || confirmando}
      >
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
