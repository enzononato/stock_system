import { useSearchParams } from 'react-router-dom'
import { CheckCircle2, TriangleAlert } from 'lucide-react'

/**
 * Tela final do coletor (rota pública): o cadastro foi salvo e, quando deu certo,
 * a conta do técnico já foi encerrada neste computador. Não depende de login.
 */
export default function ColetorConcluidoPage() {
  const [params] = useSearchParams()
  const idTexto = params.get('id') ?? ''
  const id = /^\d+$/.test(idTexto) ? idTexto : null
  // Sem `saiu` o logout ainda está em andamento.
  const saindo = !params.has('saiu')
  const saiu = params.get('saiu') === '1'

  return (
    <main className="flex min-h-screen items-center justify-center bg-background p-6 text-foreground">
      <div className="w-full max-w-md space-y-4 rounded border border-border bg-surface p-6">
        <div className="flex items-center gap-2">
          <CheckCircle2 className="size-6 text-brand-text" aria-hidden />
          <h1 className="text-heading-lg font-semibold tracking-tight">
            {id ? `Máquina cadastrada (#${id}).` : 'Máquina cadastrada.'}
          </h1>
        </div>

        {saindo ? (
          <p role="status" className="text-body-sm text-muted-foreground">
            Saindo da sua conta…
          </p>
        ) : saiu ? (
          <p role="status" className="text-body-sm text-foreground">
            Você saiu da sua conta neste computador.
          </p>
        ) : (
          <p
            role="alert"
            className="flex items-start gap-2 rounded border border-border bg-surface-alt px-3 py-2 text-body-sm text-foreground"
          >
            <TriangleAlert className="mt-0.5 size-4 shrink-0 text-destructive" aria-hidden />
            Não foi possível sair da conta automaticamente. Saia pelo menu antes de entregar a máquina.
          </p>
        )}

        <p className="text-body-sm text-muted-foreground">
          Entregue ao colaborador e imprima a etiqueta pelo seu computador (Empréstimo e Etiquetas).
        </p>
      </div>
    </main>
  )
}
