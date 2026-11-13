import { useState } from 'react'
import { Link, useSearchParams } from 'react-router-dom'
import { redefinirSenha } from '@/api/auth'
import AuthShell from '@/components/auth/AuthShell'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { mensagemErroAuth } from '@/lib/auth-errors'
import { Lock } from 'lucide-react'

const MSG_LINK_INVALIDO = 'Link inválido ou expirado. Peça um novo.'
const CLASSE_LINK = 'text-xs font-medium text-brand-text underline-offset-4 hover:underline'

function Alerta({ texto }: { texto: string }) {
  return (
    <div
      role="alert"
      className="rounded-lg border border-[var(--login-danger-border)] bg-[var(--login-danger-wash)] px-3.5 py-2.5 text-xs font-medium text-destructive text-center"
    >
      {texto}
    </div>
  )
}

export default function RedefinirSenhaPage() {
  const [params] = useSearchParams()
  const token = params.get('token') ?? ''
  const [novaSenha, setNovaSenha] = useState('')
  const [confirmacao, setConfirmacao] = useState('')
  const [enviando, setEnviando] = useState(false)
  const [erro, setErro] = useState('')
  const [linkRecusado, setLinkRecusado] = useState(false)
  const [concluido, setConcluido] = useState('')

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setErro('')
    if (novaSenha !== confirmacao) {
      setErro('As senhas não conferem.')
      return
    }
    setEnviando(true)
    try {
      const data = await redefinirSenha(token, novaSenha)
      setConcluido(data.detail)
    } catch (err) {
      setErro(mensagemErroAuth(err, 'Não foi possível redefinir a senha.'))
      // O servidor recusa o link (inválido, expirado ou já usado) com 400; senha em branco também é 400,
      // mas o formulário não deixa enviar vazio — então 400 aqui é o link.
      if ((err as { response?: { status?: number } }).response?.status === 400) setLinkRecusado(true)
    } finally {
      setEnviando(false)
    }
  }

  const semLink = !token || linkRecusado

  return (
    <AuthShell
      titulo="Redefinir senha"
      descricao={concluido || semLink ? undefined : 'Escolha uma nova senha para entrar no sistema.'}
    >
      {concluido ? (
        <div className="space-y-4">
          <p
            role="status"
            className="rounded-lg border border-border bg-surface-alt px-3.5 py-3 text-sm text-foreground leading-relaxed"
          >
            {concluido}
          </p>
          <div className="flex justify-center">
            <Link to="/login" className={CLASSE_LINK}>
              Ir para o login
            </Link>
          </div>
        </div>
      ) : semLink ? (
        <div className="space-y-4">
          <Alerta texto={erro || MSG_LINK_INVALIDO} />
          <div className="flex justify-center">
            <Link to="/esqueci-senha" className={CLASSE_LINK}>
              Pedir novo link
            </Link>
          </div>
        </div>
      ) : (
        <form onSubmit={handleSubmit} className="space-y-4" noValidate>
          <div className="space-y-1.5">
            <Label htmlFor="nova-senha" className="text-xs font-medium text-foreground">
              Nova senha
            </Label>
            <div className="relative rounded-md">
              <Lock
                size={18}
                className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
                aria-hidden
              />
              <Input
                id="nova-senha"
                type="password"
                value={novaSenha}
                onChange={(e) => setNovaSenha(e.target.value)}
                autoComplete="new-password"
                className="h-11 pl-10 bg-[var(--login-glass)] border-input text-sm"
                autoFocus
              />
            </div>
          </div>
          <div className="space-y-1.5">
            <Label htmlFor="confirmacao-senha" className="text-xs font-medium text-foreground">
              Confirme a nova senha
            </Label>
            <div className="relative rounded-md">
              <Lock
                size={18}
                className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
                aria-hidden
              />
              <Input
                id="confirmacao-senha"
                type="password"
                value={confirmacao}
                onChange={(e) => setConfirmacao(e.target.value)}
                autoComplete="new-password"
                className="h-11 pl-10 bg-[var(--login-glass)] border-input text-sm"
              />
            </div>
          </div>

          {erro && <Alerta texto={erro} />}

          <Button
            type="submit"
            size="lg"
            className="h-11 w-full font-semibold"
            loading={enviando}
            disabled={!novaSenha}
          >
            Salvar nova senha
          </Button>
        </form>
      )}
    </AuthShell>
  )
}
