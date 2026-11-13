import { useEffect, useRef, useState } from 'react'
import type { CadastroTotp } from '@/api/auth'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { QrCodeOtpauth } from '@/components/auth/QrCodeOtpauth'
import { mensagemErroAuth } from '@/lib/auth-errors'

interface CadastroAutenticadorProps {
  iniciar: () => Promise<CadastroTotp>
  confirmar: (codigo: string) => Promise<string[]>
  onConcluido: (codigos: string[]) => void
}

const MSG_ERRO = 'Não foi possível concluir o cadastro.'

/** Cadastro do aplicativo autenticador: QR code, chave em texto e confirmação com um código. */
export function CadastroAutenticador({ iniciar, confirmar, onConcluido }: CadastroAutenticadorProps) {
  const [dados, setDados] = useState<CadastroTotp | null>(null)
  const [codigo, setCodigo] = useState('')
  const [erro, setErro] = useState('')
  const [enviando, setEnviando] = useState(false)
  const iniciou = useRef(false)

  useEffect(() => {
    // `iniciar` gera um segredo novo no servidor: chamar uma única vez por montagem.
    if (iniciou.current) return
    iniciou.current = true
    iniciar()
      .then(setDados)
      .catch((err: unknown) => setErro(mensagemErroAuth(err, MSG_ERRO)))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setErro('')
    setEnviando(true)
    try {
      const codigos = await confirmar(codigo.replace(/\s/g, ''))
      onConcluido(codigos)
    } catch (err) {
      setErro(mensagemErroAuth(err, MSG_ERRO))
    } finally {
      setEnviando(false)
    }
  }

  return (
    <div className="space-y-4">
      {!dados && !erro && <p className="text-sm text-muted-foreground">Gerando QR code...</p>}

      {dados && (
        <form onSubmit={handleSubmit} className="space-y-4" noValidate>
          <div className="flex justify-center">
            <QrCodeOtpauth uri={dados.otpauth_uri} />
          </div>
          <div className="space-y-1 text-center">
            <p className="text-xs text-muted-foreground">
              Não consegue ler o QR code? Digite esta chave no aplicativo:
            </p>
            <p className="break-all font-mono text-sm font-semibold tracking-wide text-foreground">
              {dados.segredo.replace(/(.{4})/g, '$1 ').trim()}
            </p>
          </div>
          <div className="space-y-1.5">
            <Label htmlFor="codigo-cadastro" className="text-xs font-medium text-foreground">
              Código de 6 dígitos
            </Label>
            <Input
              id="codigo-cadastro"
              value={codigo}
              onChange={(e) => setCodigo(e.target.value)}
              inputMode="numeric"
              autoComplete="one-time-code"
              placeholder="000000"
              className="h-11 text-center font-mono text-lg tracking-widest"
              autoFocus
            />
          </div>
          {erro && <MensagemErro texto={erro} />}
          <Button type="submit" className="w-full" loading={enviando} disabled={!codigo.trim()}>
            Confirmar
          </Button>
        </form>
      )}

      {!dados && erro && <MensagemErro texto={erro} />}
    </div>
  )
}

function MensagemErro({ texto }: { texto: string }) {
  return (
    <div
      role="alert"
      className="rounded-lg border border-[var(--login-danger-border)] bg-[var(--login-danger-wash)] px-3.5 py-2.5 text-center text-xs font-medium text-destructive"
    >
      {texto}
    </div>
  )
}
