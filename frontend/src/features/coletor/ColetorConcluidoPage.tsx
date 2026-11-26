import { useEffect, useState } from 'react'
import { useSearchParams } from 'react-router-dom'
import { CheckCircle2, TriangleAlert } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { useAuth } from '@/contexts/AuthContext'

/** Sem `saiu` por mais que isso (ex.: página recarregada), o logout não vai mais terminar sozinho. */
const ESPERA_SAIDA_MS = 5000

/**
 * Tela final do coletor (rota pública): o cadastro foi salvo e, quando deu certo,
 * a conta do técnico já foi encerrada neste computador. Não depende de login.
 */
export default function ColetorConcluidoPage() {
  const [params, setParams] = useSearchParams()
  const { logout } = useAuth()
  const [tentando, setTentando] = useState(false)
  const [falhou, setFalhou] = useState(false)
  const [demorou, setDemorou] = useState(false)
  const idTexto = params.get('id') ?? ''
  const id = /^\d+$/.test(idTexto) ? idTexto : null
  // Sem `saiu` o logout ainda está em andamento.
  const saindo = !params.has('saiu')
  const saiu = params.get('saiu') === '1'

  useEffect(() => {
    if (!saindo) return
    const t = setTimeout(() => setDemorou(true), ESPERA_SAIDA_MS)
    return () => clearTimeout(t)
  }, [saindo])

  async function tentarSairDeNovo() {
    if (tentando) return
    setTentando(true)
    setFalhou(false)
    try {
      await logout()
      setParams(
        (atual) => {
          const novo = new URLSearchParams(atual)
          novo.set('saiu', '1')
          return novo
        },
        { replace: true }
      )
    } catch {
      setFalhou(true)
    } finally {
      setTentando(false)
    }
  }

  const botaoTentar = (
    <div className="space-y-2">
      <Button variant="outline" size="sm" onClick={() => void tentarSairDeNovo()} loading={tentando}>
        Tentar sair de novo
      </Button>
      {falhou && (
        <p role="alert" className="text-body-sm text-destructive">
          Ainda não foi possível sair.
        </p>
      )}
    </div>
  )

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
          <>
            <p role="status" className="text-body-sm text-muted-foreground">
              Saindo da sua conta…
            </p>
            {demorou && botaoTentar}
          </>
        ) : saiu ? (
          <p role="status" className="text-body-sm text-foreground">
            Você saiu da sua conta neste computador.
          </p>
        ) : (
          <>
            <p
              role="alert"
              className="flex items-start gap-2 rounded border border-border bg-surface-alt px-3 py-2 text-body-sm text-foreground"
            >
              <TriangleAlert className="mt-0.5 size-4 shrink-0 text-destructive" aria-hidden />
              Não foi possível sair da conta automaticamente. Saia pelo menu antes de entregar a máquina.
            </p>
            {botaoTentar}
          </>
        )}

        <p className="text-body-sm text-muted-foreground">
          Entregue ao colaborador e imprima a etiqueta pelo seu computador (Empréstimo e Etiquetas).
        </p>
      </div>
    </main>
  )
}
