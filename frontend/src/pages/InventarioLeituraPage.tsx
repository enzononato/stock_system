import { useRef, useState } from 'react'
import { Link, useParams } from 'react-router-dom'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { isAxiosError } from 'axios'
import { Camera, CameraOff } from 'lucide-react'
import {
  atualizarLeitura,
  excluirLeitura,
  getInventario,
  listLeituras,
  registrarLeitura,
  type AtualizarLeituraInput,
  type Leitura,
} from '@/api/inventarios'
import { LeitorCamera } from '@/components/inventario/LeitorCamera'
import { LocalEncontradoSelect } from '@/components/inventario/LocalEncontradoSelect'
import { UltimaLeituraCard } from '@/components/inventario/UltimaLeituraCard'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { toast } from '@/components/ui/toast'
import { MENSAGEM_ERRO_CAMERA, type ErroCamera } from '@/features/inventario/scanner'
import { ROTULO_RESULTADO, TOM_RESULTADO } from '@/features/inventario/resultados'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
import { formatDateTime } from '@/lib/utils'

const linkClass =
  'rounded-sm text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2'

/**
 * Tela de leitura do inventário, pensada para o celular: câmera (QR da etiqueta),
 * digitação ou leitor USB (que digita o código e dá Enter), tudo pela mesma mutação.
 */
export default function InventarioLeituraPage() {
  const { id: idParam } = useParams()
  const id = Number(idParam)
  const queryClient = useQueryClient()
  const campoRef = useRef<HTMLInputElement>(null)
  const registrando = useRef(false)
  /** Códigos digitados/lidos pelo USB enquanto uma gravação está em andamento (FIFO). */
  const fila = useRef<string[]>([])

  const [localId, setLocalId] = useState<number | null>(null)
  const [codigo, setCodigo] = useState('')
  const [cameraAtiva, setCameraAtiva] = useState(false)
  const [avisoCamera, setAvisoCamera] = useState<{ texto: string; alerta: boolean } | null>(null)
  const [ultima, setUltima] = useState<Leitura | null>(null)

  const inventarioQuery = useQuery({
    queryKey: ['inventario', id],
    queryFn: () => getInventario(id),
    enabled: Number.isInteger(id),
  })
  const inventario = inventarioQuery.data
  const aberto = inventario?.status === 'aberto'

  const { data: recentes } = useQuery({
    queryKey: ['inventario-leituras', id],
    queryFn: () => listLeituras(id, { limit: 10 }),
    enabled: !!inventario && aberto,
  })

  function invalidar() {
    void queryClient.invalidateQueries({ queryKey: ['inventario', id] })
    void queryClient.invalidateQueries({ queryKey: ['inventario-leituras', id] })
    void queryClient.invalidateQueries({ queryKey: ['inventarios'] })
  }

  function focarCampo() {
    campoRef.current?.focus()
  }

  /** Erro de escrita: com 409 o inventário foi encerrado/cancelado por outra pessoa, e a tela precisa saber. */
  function tratarErro(err: unknown, fallback: string) {
    toast.error(getErrorMessage(err, fallback))
    if (isAxiosError(err) && err.response?.status === 409) {
      fila.current = []
      void queryClient.invalidateQueries({ queryKey: ['inventario', id] })
    }
  }

  const registrar = useMutation({
    mutationFn: ({ texto }: { texto: string; origem: 'formulario' | 'camera' }) =>
      registrarLeitura(id, { codigo: texto, local_id_encontrado: localId }),
    onSuccess: (leitura, { origem }) => {
      setUltima(leitura)
      invalidar()
      // Depois da câmera o foco não vai para o campo: abriria o teclado do celular por cima do vídeo.
      if (origem === 'formulario') focarCampo()
    },
    onError: (err: unknown, { origem }) => {
      tratarErro(err, 'Erro ao registrar a leitura.')
      if (origem === 'formulario') focarCampo()
    },
    onSettled: () => {
      registrando.current = false
      // O que chegou durante a gravação segue pelo mesmo caminho, um de cada vez.
      const proximo = fila.current.shift()
      if (proximo) enviarDoFormulario(proximo)
    },
  })

  function enviarDoFormulario(texto: string) {
    if (registrando.current) {
      // Duplicata exata do último código da fila (leitor que repete o envio) não entra de novo.
      if (fila.current[fila.current.length - 1] !== texto) fila.current.push(texto)
      return
    }
    registrando.current = true
    registrar.mutate({ texto, origem: 'formulario' })
  }

  const corrigir = useMutation({
    mutationFn: (dados: AtualizarLeituraInput) => atualizarLeitura(id, (ultima as Leitura).id, dados),
    onSuccess: (leitura) => {
      setUltima({ ...leitura, duplicada: false })
      invalidar()
    },
    onError: (err: unknown) => tratarErro(err, 'Erro ao corrigir a leitura.'),
  })

  const desfazer = useMutation({
    mutationFn: () => excluirLeitura(id, (ultima as Leitura).id),
    onSuccess: () => {
      setUltima(null)
      invalidar()
      toast.success('Leitura desfeita.')
      focarCampo()
    },
    onError: (err: unknown) => tratarErro(err, 'Erro ao desfazer a leitura.'),
  })

  /** `false` = ocupada: a câmera não conta esse código como entregue e o entrega de novo no próximo quadro. */
  function lerDaCamera(texto: string): boolean {
    if (registrando.current) return false
    registrando.current = true
    registrar.mutate({ texto, origem: 'camera' })
    return true
  }

  function aoErroDaCamera(erro: ErroCamera) {
    setCameraAtiva(false)
    setAvisoCamera({ texto: MENSAGEM_ERRO_CAMERA[erro.tipo], alerta: true })
    focarCampo()
  }

  function aoPausarCamera() {
    setCameraAtiva(false)
    setAvisoCamera({ texto: 'Câmera pausada. Toque em "Ler com a câmera" para continuar.', alerta: false })
  }

  if (inventarioQuery.isLoading) return <LoadingState label="Carregando inventário..." />
  if (inventarioQuery.isError || !inventario) {
    return <ErrorState error={inventarioQuery.error} onRetry={() => inventarioQuery.refetch()} />
  }

  if (!aberto) {
    return (
      <div className="mx-auto max-w-xl">
        <EmptyState
          title="Este inventário não está aberto."
          description="Só é possível ler equipamentos enquanto o inventário está aberto."
          action={
            <Link to={`/inventarios/${id}`} className={linkClass}>
              Ver relatório
            </Link>
          }
        />
      </div>
    )
  }

  const { resumo } = inventario
  const seguro = window.isSecureContext
  const outrasLeituras = (recentes?.items ?? []).filter((l) => l.id !== ultima?.id)

  return (
    <div className="mx-auto max-w-xl space-y-4">
      <header className="space-y-1">
        <h1 className="text-heading-sm font-semibold">{`Inventário #${inventario.id} · ${upper(inventario.unidade_nome)}`}</h1>
        <p className="text-body-sm text-muted-foreground">
          {`Encontrados ${resumo.encontrados} de ${resumo.esperados} · Faltam ${resumo.nao_encontrados}`}
        </p>
        <Link to={`/inventarios/${id}`} className={`${linkClass} text-body-sm`}>
          Relatório
        </Link>
      </header>

      <label className="flex flex-col gap-1 text-caption text-muted-foreground">
        Estou no local
        <LocalEncontradoSelect unidadeId={inventario.unidade_id} value={localId} onChange={setLocalId} />
      </label>

      <div className="space-y-3">
        {cameraAtiva ? (
          <>
            <LeitorCamera onCodigo={lerDaCamera} onErro={aoErroDaCamera} onPausar={aoPausarCamera} />
            <Button variant="outline" size="lg" onClick={() => setCameraAtiva(false)}>
              <CameraOff aria-hidden="true" />
              Parar câmera
            </Button>
          </>
        ) : (
          <Button
            size="lg"
            disabled={!seguro}
            onClick={() => {
              setAvisoCamera(null)
              setCameraAtiva(true)
            }}
          >
            <Camera aria-hidden="true" />
            Ler com a câmera
          </Button>
        )}
        {!seguro && <p className="text-body-sm text-muted-foreground">{MENSAGEM_ERRO_CAMERA.inseguro}</p>}
        {avisoCamera &&
          (avisoCamera.alerta ? (
            <p role="alert" className="text-body-sm text-destructive">
              {avisoCamera.texto}
            </p>
          ) : (
            <p className="text-body-sm text-muted-foreground">{avisoCamera.texto}</p>
          ))}
      </div>

      <form
        className="space-y-2"
        onSubmit={(e) => {
          e.preventDefault()
          const texto = codigo.trim()
          if (!texto) return
          // O campo esvazia sempre ao enviar: o próximo código do leitor USB nunca se soma a este.
          setCodigo('')
          enviarDoFormulario(texto)
        }}
      >
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Código ou número do patrimônio
          <Input
            ref={campoRef}
            autoFocus
            autoComplete="off"
            enterKeyHint="send"
            value={codigo}
            onChange={(e) => setCodigo(e.target.value)}
            className="h-11 text-base"
          />
        </label>
        {/* Sem `loading`: um botão desabilitado impede o Enter do leitor USB de enviar o formulário. */}
        <Button type="submit" size="lg" aria-busy={registrar.isPending || undefined}>
          Registrar
        </Button>
      </form>

      {ultima && (
        <UltimaLeituraCard
          leitura={ultima}
          ocupado={corrigir.isPending || desfazer.isPending}
          onInformarUsuario={(colaboradorId) => corrigir.mutate({ colaborador_id_encontrado: colaboradorId })}
          onSemUsuario={() => corrigir.mutate({ usuario_conferido: true, colaborador_id_encontrado: null })}
          onDesfazer={() => desfazer.mutate()}
        />
      )}

      {outrasLeituras.length > 0 && (
        <section aria-label="Últimas leituras" className="space-y-2">
          <h2 className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
            Últimas leituras
          </h2>
          <ul className="surface-panel divide-y divide-border">
            {outrasLeituras.map((l) => (
              <li key={l.id} className="flex flex-wrap items-center gap-2 px-3 py-2 text-body-sm">
                <StatusPill status={ROTULO_RESULTADO[l.resultado]} tone={TOM_RESULTADO[l.resultado]} />
                <span className="font-medium">
                  {l.item_id != null ? `#${l.item_id} ${l.item_tipo ?? ''}`.trim() : l.codigo_lido}
                </span>
                <span className="ml-auto text-muted-foreground">{formatDateTime(l.lido_em)}</span>
              </li>
            ))}
          </ul>
        </section>
      )}
    </div>
  )
}
