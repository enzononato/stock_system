import { useEffect, useRef, useState } from 'react'
import { Link, useNavigate } from 'react-router-dom'
import { useQuery } from '@tanstack/react-query'
import { Copy, Download, LogOut } from 'lucide-react'
import { andamentoColeta, criarColeta, type AndamentoColeta, type Correspondencia } from '@/api/coletor'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { ErrorState } from '@/components/ui/StateBlocks'
import { useAuth } from '@/contexts/AuthContext'
import type { EstadoColetor } from './estadoColetor'

const INTERVALO_CONSULTA_MS = 2000

const TEXTO_MOTIVO: Record<string, string> = {
  serial: 'mesmo número de série',
  mac: 'mesmo MAC',
  anydesk: 'mesmo AnyDesk',
}

/** "#123 Notebook Dell Latitude 3420 · RV-JZ · mesmo MAC" */
function descreverCorrespondencia(c: Correspondencia): string {
  const equipamento = [`#${c.id}`, c.tipo, c.brand, c.model].filter(Boolean).join(' ')
  const motivos = c.motivos.map((m) => TEXTO_MOTIVO[m] ?? m).join(', ')
  return [equipamento, c.revenda, motivos].filter(Boolean).join(' · ')
}

function montarEstado(coletaId: number, andamento: AndamentoColeta): EstadoColetor {
  return {
    origem: 'coletor',
    coletaId,
    campos: andamento.campos ?? {},
    tipoSugerido: andamento.tipo_sugerido ?? null,
  }
}

export default function ColetorPage() {
  const navigate = useNavigate()
  const { logout } = useAuth()
  // Cada "tentativa" é uma coleta: "Gerar novo código" só incrementa o contador.
  const [tentativa, setTentativa] = useState(0)
  const [aviso, setAviso] = useState<string | null>(null)
  const [saindo, setSaindo] = useState(false)
  const [erroSaida, setErroSaida] = useState(false)
  const campoComando = useRef<HTMLTextAreaElement>(null)

  // Query (e não useEffect + mutation) para criar a coleta uma única vez por montagem, mesmo com o
  // StrictMode executando os efeitos duas vezes. Sem cache: voltar à página gera um código novo.
  const coleta = useQuery({
    queryKey: ['coletor', 'coleta', tentativa],
    queryFn: criarColeta,
    staleTime: Infinity,
    gcTime: 0,
    retry: false,
    refetchOnWindowFocus: false,
    refetchOnReconnect: false,
  })
  const coletaId = coleta.data?.id

  const andamento = useQuery({
    queryKey: ['coletor', 'andamento', coletaId],
    queryFn: () => andamentoColeta(coletaId as number),
    enabled: coletaId !== undefined,
    gcTime: 0,
    retry: false,
    refetchInterval: (query) => (query.state.data?.status === 'aguardando' ? INTERVALO_CONSULTA_MS : false),
  })

  const dados = andamento.data
  const recebido = dados?.status === 'recebido'
  const semCorrespondencias = recebido && (dados.correspondencias ?? []).length === 0

  useEffect(() => {
    if (coletaId !== undefined && dados && semCorrespondencias) {
      void navigate('/register', { state: montarEstado(coletaId, dados) })
    }
  }, [coletaId, dados, semCorrespondencias, navigate])

  async function copiar() {
    const comando = coleta.data?.comando
    if (!comando) return
    try {
      await navigator.clipboard.writeText(comando)
      setAviso('Comando copiado.')
    } catch {
      campoComando.current?.focus()
      campoComando.current?.select()
      setAviso('Selecione e copie com Ctrl+C')
    }
  }

  async function sair() {
    if (saindo) return
    setSaindo(true)
    setErroSaida(false)
    try {
      await logout()
      void navigate('/login')
    } catch {
      // A sessão continua aberta neste computador: não leva ao login como se tivesse saído.
      setErroSaida(true)
    } finally {
      setSaindo(false)
    }
  }

  function gerarNovo() {
    setAviso(null)
    setTentativa((t) => t + 1)
  }

  function irPara(destino: string) {
    if (coletaId !== undefined && dados) void navigate(destino, { state: montarEstado(coletaId, dados) })
  }

  let corpo
  if (coleta.isError) {
    corpo = (
      <ErrorState
        error={coleta.error}
        title="Não foi possível gerar o código"
        onRetry={() => void coleta.refetch()}
      />
    )
  } else if (!coleta.data) {
    corpo = <p className="text-sm text-muted-foreground">Gerando o código…</p>
  } else if (andamento.isError && !dados) {
    corpo = (
      <div className="space-y-4">
        <ErrorState
          title="Não foi possível acompanhar a coleta."
          message="Gere um novo código e tente de novo."
        />
        <Button onClick={gerarNovo}>Gerar novo código</Button>
      </div>
    )
  } else if (dados?.status === 'vencido') {
    corpo = (
      <div className="space-y-3">
        <p className="text-sm font-medium text-foreground">O código venceu.</p>
        <Button onClick={gerarNovo}>Gerar novo código</Button>
      </div>
    )
  } else if (recebido && !semCorrespondencias) {
    corpo = (
      <div className="space-y-4">
        <p className="text-sm text-foreground">
          Os dados chegaram. Esta máquina parece já estar cadastrada:
        </p>
        <ul className="space-y-3">
          {(dados.correspondencias ?? []).map((c) => (
            <li key={c.id} className="flex flex-col gap-2 rounded border border-border p-3 sm:flex-row sm:items-center sm:justify-between">
              <span className="text-sm text-foreground">{descreverCorrespondencia(c)}</span>
              {c.ativo ? (
                <Button size="sm" onClick={() => irPara(`/edit/${c.id}`)}>
                  Atualizar este cadastro
                </Button>
              ) : (
                <span className="flex items-center gap-3 text-sm text-muted-foreground">
                  Consta como baixada
                  <Link to={`/items/${c.id}`} className="underline underline-offset-4 text-foreground">
                    Ver registro
                  </Link>
                </span>
              )}
            </li>
          ))}
        </ul>
        <Button variant="outline" onClick={() => irPara('/register')}>
          Cadastrar como novo mesmo assim
        </Button>
      </div>
    )
  } else {
    corpo = (
      <div className="space-y-4">
        <p className="text-sm text-foreground">
          Aperte <strong>Windows + R</strong>, cole com <strong>Ctrl+V</strong> e dê <strong>Enter</strong>.
        </p>
        <textarea
          ref={campoComando}
          readOnly
          rows={4}
          aria-label="Comando para colar no Executar"
          value={coleta.data.comando}
          onFocus={(e) => e.currentTarget.select()}
          className="w-full resize-none rounded border border-input bg-surface-alt p-3 font-mono text-xs text-foreground focus-visible:ring-1 focus-visible:ring-ring"
        />
        <div className="flex flex-wrap items-center gap-3">
          <Button onClick={() => void copiar()}>
            <Copy /> Copiar comando
          </Button>
          {/* Região viva sempre montada: leitores de tela só anunciam texto que muda dentro dela. */}
          <span role="status" className="text-sm text-muted-foreground">
            {aviso}
          </span>
        </div>
        <p className="text-sm text-muted-foreground">⏳ Aguardando os dados da máquina…</p>
        <a
          href={coleta.data.url_arquivo}
          download
          className="inline-flex items-center gap-1.5 text-sm text-muted-foreground underline underline-offset-4 hover:text-foreground"
        >
          <Download className="size-3.5" /> Baixar o coletor
        </a>
      </div>
    )
  }

  return (
    <div className="mx-auto max-w-3xl space-y-6">
      <PageHeader
        eyebrow="Inventário"
        eyebrowDetail="Coletor"
        title="Cadastrar esta máquina"
        description="Leia automaticamente os dados deste computador e leve para o cadastro."
        actions={
          <Button variant="outline" size="sm" onClick={() => void sair()} loading={saindo}>
            {!saindo && <LogOut />} Sair deste computador
          </Button>
        }
      />
      {erroSaida && (
        <p role="alert" className="text-sm text-destructive">
          Não foi possível sair da conta. Tente de novo.
        </p>
      )}
      <section className="rounded border border-border bg-surface p-6">{corpo}</section>
    </div>
  )
}
