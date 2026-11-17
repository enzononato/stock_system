import { useState } from 'react'
import { Link } from 'react-router-dom'
import { esqueciSenha } from '@/api/auth'
import AuthShell from '@/components/auth/AuthShell'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { mensagemErroAuth } from '@/lib/auth-errors'
import { Mail } from 'lucide-react'

const LINK_VOLTAR = 'text-xs font-medium text-brand-text underline-offset-4 hover:underline'

export default function EsqueciSenhaPage() {
  const [identificador, setIdentificador] = useState('')
  const [enviando, setEnviando] = useState(false)
  const [erro, setErro] = useState('')
  const [resposta, setResposta] = useState('')

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setErro('')
    setEnviando(true)
    try {
      const data = await esqueciSenha(identificador.trim())
      setResposta(data.detail)
    } catch (err) {
      setErro(mensagemErroAuth(err, 'Não foi possível enviar o pedido. Tente novamente.'))
    } finally {
      setEnviando(false)
    }
  }

  return (
    <AuthShell
      titulo="Esqueci minha senha"
      descricao="Informe o seu e-mail. Se ele estiver cadastrado e confirmado, enviaremos um link para criar uma nova senha."
    >
      {resposta ? (
        <p
          role="status"
          className="rounded-lg border border-border bg-surface-alt px-3.5 py-3 text-sm text-foreground leading-relaxed"
        >
          {resposta}
        </p>
      ) : (
        <form onSubmit={handleSubmit} className="space-y-4" noValidate>
          <div className="space-y-1.5">
            <Label htmlFor="identificador" className="text-xs font-medium text-foreground">
              E-mail
            </Label>
            <div className="relative rounded-md">
              <Mail
                size={18}
                className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
                aria-hidden
              />
              <Input
                id="identificador"
                name="identificador"
                value={identificador}
                onChange={(e) => setIdentificador(e.target.value)}
                autoComplete="username"
                placeholder="nome@revalle.com.br"
                className="h-11 pl-10 bg-[var(--login-glass)] border-input text-sm"
                autoFocus
                required
              />
            </div>
          </div>

          {erro && (
            <div
              role="alert"
              className="rounded-lg border border-[var(--login-danger-border)] bg-[var(--login-danger-wash)] px-3.5 py-2.5 text-xs font-medium text-destructive text-center"
            >
              {erro}
            </div>
          )}

          <Button
            type="submit"
            size="lg"
            className="h-11 w-full font-semibold"
            loading={enviando}
            disabled={!identificador.trim()}
          >
            Enviar link
          </Button>
        </form>
      )}

      <div className="mt-5 flex justify-center">
        <Link to="/login" className={LINK_VOLTAR}>
          Voltar para o login
        </Link>
      </div>
    </AuthShell>
  )
}
