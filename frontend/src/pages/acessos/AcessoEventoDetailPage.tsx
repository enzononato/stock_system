import { useRef, useState } from 'react'
import { Link, useParams } from 'react-router-dom'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { isAxiosError } from 'axios'
import { ArrowLeft } from 'lucide-react'
import {
  cancelarEvento,
  concluirEvento,
  getEvento,
  type EventoDetalhe,
  type Tarefa,
} from '@/api/acessos'
import { ResolverTarefaDialog, type ModoResolucao } from '@/components/acessos/ResolverTarefaDialog'
import {
  ROTULO_ACAO_TAREFA,
  ROTULO_TIPO_EVENTO,
  STATUS_EVENTO,
  STATUS_TAREFA,
  descreverEquipamento,
} from '@/components/acessos/rotulos'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
import { formatDateTime } from '@/lib/utils'

const linkClass =
  'rounded-sm font-semibold text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2'

function cargoDoEvento(e: EventoDetalhe): string | null {
  if (e.tipo === 'mudanca_funcao') return `${e.cargo_anterior ?? '—'} → ${e.cargo_novo ?? '—'}`
  return e.cargo_novo ?? e.cargo_anterior
}

/** Detalhe do evento de acesso: checklist de tarefas, equipamentos (desligamento), conclusão e cancelamento. */
export default function AcessoEventoDetailPage() {
  const { id } = useParams()
  const eventoId = Number(id)
  const queryClient = useQueryClient()
  const chave = ['acessos', 'evento', eventoId] as const

  const [resolvendo, setResolvendo] = useState<{ tarefa: Tarefa; modo: ModoResolucao } | null>(null)
  const [cancelando, setCancelando] = useState(false)
  const concluindo = useRef(false)

  const { data: evento, isLoading, isError, error, refetch } = useQuery({
    queryKey: chave,
    queryFn: () => getEvento(eventoId),
    enabled: Number.isFinite(eventoId),
    retry: false,
  })

  function aplicar(novo: EventoDetalhe) {
    queryClient.setQueryData(chave, novo)
    queryClient.invalidateQueries({ queryKey: ['acessos'] })
  }

  const concluir = useMutation({
    mutationFn: () => concluirEvento(eventoId),
    onSuccess: (novo) => {
      aplicar(novo)
      toast('Evento concluído.')
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao concluir o evento.'), 'error'),
    onSettled: () => {
      concluindo.current = false
    },
  })

  if (isLoading) return <LoadingState label="Carregando evento..." />
  if (isError || !evento) {
    if (isAxiosError(error) && error.response?.status === 404) {
      return (
        <div className="space-y-4">
          <Link to="/acessos/eventos" className={`inline-flex items-center gap-1 text-sm ${linkClass}`}>
            <ArrowLeft size={14} aria-hidden="true" />
            Voltar aos eventos
          </Link>
          <EmptyState title="Evento não encontrado" description="O evento pode ter sido removido ou o endereço está incorreto." />
        </div>
      )
    }
    return <ErrorState error={error} onRetry={() => refetch()} />
  }

  const situacao = STATUS_EVENTO[evento.status]
  const aberto = evento.status === 'aberto'
  const cargo = cargoDoEvento(evento)

  return (
    <div className="space-y-6">
      <Link to="/acessos/eventos" className={`inline-flex items-center gap-1 text-sm ${linkClass}`}>
        <ArrowLeft size={14} aria-hidden="true" />
        Voltar aos eventos
      </Link>

      <header className="space-y-2 border-b border-border pb-4">
        <div className="flex flex-wrap items-center gap-3">
          <h1 className="text-2xl font-semibold tracking-tight">{ROTULO_TIPO_EVENTO[evento.tipo]}</h1>
          <StatusPill status={situacao.label} tone={situacao.tone} />
        </div>
        <p className="text-sm">
          <Link to={`/colaboradores/${evento.colaborador_id}`} className={linkClass}>
            {upper(evento.colaborador_nome)}
          </Link>
          {cargo && <span className="text-muted-foreground"> · </span>}
          {cargo && <span>{cargo}</span>}
        </p>
        <p className="text-sm text-muted-foreground">
          {`Aberto por ${evento.aberto_por} em ${formatDateTime(evento.aberto_em)}`}
          {evento.status === 'concluido' &&
            ` · Concluído por ${evento.concluido_por ?? '—'} em ${formatDateTime(evento.concluido_em)}`}
          {evento.status === 'cancelado' &&
            ` · Cancelado por ${evento.cancelado_por ?? '—'} em ${formatDateTime(evento.cancelado_em)} — Motivo: ${evento.motivo_cancelamento ?? '—'}`}
        </p>
        {evento.observacao && <p className="text-sm">{evento.observacao}</p>}
      </header>

      <section aria-labelledby="titulo-tarefas" className="space-y-3">
        <h2 id="titulo-tarefas" className="text-base font-semibold">
          Tarefas
        </h2>
        {evento.tarefas.length === 0 ? (
          <p className="text-sm text-muted-foreground">Este evento não gerou tarefas.</p>
        ) : (
          <ul aria-label="Tarefas" className="divide-y divide-border rounded-lg border border-border">
            {evento.tarefas.map((t) => {
              const st = STATUS_TAREFA[t.status]
              const pendente = aberto && t.status === 'pendente'
              return (
                <li key={t.id} className="flex flex-wrap items-start justify-between gap-3 px-4 py-3">
                  <div className="min-w-0 space-y-1">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="text-sm font-semibold">{t.sistema_nome}</span>
                      <span className="text-xs text-muted-foreground">{ROTULO_ACAO_TAREFA[t.acao]}</span>
                      <StatusPill status={st.label} tone={st.tone} />
                    </div>
                    {t.login && <p className="text-sm">{t.login}</p>}
                    {t.feito_por && (
                      <p className="text-xs text-muted-foreground">
                        {`Feito por ${t.feito_por} em ${formatDateTime(t.feito_em)}`}
                      </p>
                    )}
                    {t.observacao && <p className="text-xs text-muted-foreground">{t.observacao}</p>}
                  </div>
                  {pendente && (
                    <div className="flex flex-wrap gap-2">
                      <Button size="sm" onClick={() => setResolvendo({ tarefa: t, modo: 'feito' })}>
                        {t.acao === 'criar' ? 'Marcar como feito' : 'Marcar como desativado'}
                      </Button>
                      {t.acao === 'desativar' && evento.tipo === 'mudanca_funcao' ? (
                        <Button size="sm" variant="outline" onClick={() => setResolvendo({ tarefa: t, modo: 'manter' })}>
                          Manter acesso
                        </Button>
                      ) : (
                        <Button size="sm" variant="outline" onClick={() => setResolvendo({ tarefa: t, modo: 'nao_se_aplica' })}>
                          Não se aplica
                        </Button>
                      )}
                    </div>
                  )}
                </li>
              )
            })}
          </ul>
        )}
      </section>

      {evento.tipo === 'desligamento' && (
        <section aria-labelledby="titulo-equipamentos" className="space-y-3">
          <h2 id="titulo-equipamentos" className="text-base font-semibold">
            Equipamentos com o colaborador
          </h2>
          {evento.equipamentos.length === 0 ? (
            <p className="text-sm text-muted-foreground">Nenhum equipamento com o colaborador.</p>
          ) : (
            <>
              <ul className="divide-y divide-border rounded-lg border border-border">
                {evento.equipamentos.map((e) => (
                  <li key={e.id} className="flex flex-wrap items-center justify-between gap-2 px-4 py-3">
                    <div className="min-w-0">
                      <Link to={`/items/${e.id}`} className={`text-sm ${linkClass}`}>
                        {descreverEquipamento(e)}
                      </Link>
                      {e.revenda && <p className="text-xs text-muted-foreground">{upper(e.revenda)}</p>}
                    </div>
                    <StatusPill status={e.status} />
                  </li>
                ))}
              </ul>
              <Link to="/return" className={`inline-block text-sm ${linkClass}`}>
                Ir para Devolução
              </Link>
            </>
          )}
        </section>
      )}

      {aberto && (
        <section aria-label="Conclusão do evento" className="space-y-3">
          {evento.pendencias.length > 0 && (
            <div role="status" className="rounded-md bg-surface-alt p-3 text-sm">
              <p className="font-medium">Para concluir, falta:</p>
              <ul className="list-disc pl-5">
                {evento.pendencias.map((p) => (
                  <li key={p}>{p}</li>
                ))}
              </ul>
            </div>
          )}
          <div className="flex flex-wrap gap-2">
            <Button
              onClick={() => {
                // Trava síncrona além do `loading`: dois cliques seguidos antes do repintar.
                if (concluindo.current) return
                concluindo.current = true
                concluir.mutate()
              }}
              loading={concluir.isPending}
              disabled={!evento.pode_concluir}
            >
              Concluir evento
            </Button>
            <Button variant="outline" onClick={() => setCancelando(true)}>
              Cancelar evento
            </Button>
          </div>
        </section>
      )}

      <ResolverTarefaDialog
        tarefa={resolvendo?.tarefa ?? null}
        modo={resolvendo?.modo ?? null}
        colaboradorId={evento.colaborador_id}
        onClose={() => setResolvendo(null)}
        onResolved={aplicar}
      />
      <CancelarEventoDialog
        open={cancelando}
        eventoId={eventoId}
        onClose={() => setCancelando(false)}
        onCancelled={aplicar}
      />
    </div>
  )
}

function CancelarEventoDialog({
  open,
  eventoId,
  onClose,
  onCancelled,
}: {
  open: boolean
  eventoId: number
  onClose: () => void
  onCancelled: (evento: EventoDetalhe) => void
}) {
  return (
    <Dialog open={open} onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-md">
        <CancelarForm eventoId={eventoId} onClose={onClose} onCancelled={onCancelled} />
      </DialogContent>
    </Dialog>
  )
}

function CancelarForm({
  eventoId,
  onClose,
  onCancelled,
}: {
  eventoId: number
  onClose: () => void
  onCancelled: (evento: EventoDetalhe) => void
}) {
  const [motivo, setMotivo] = useState('')
  const [erro, setErro] = useState<string | null>(null)
  const mutation = useMutation({
    mutationFn: () => cancelarEvento(eventoId, { motivo_cancelamento: motivo.trim() }),
    onSuccess: (novo) => {
      toast('Evento cancelado.')
      onCancelled(novo)
      onClose()
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao cancelar o evento.'), 'error'),
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (mutation.isPending) return
    if (!motivo.trim()) {
      setErro('Informe o motivo do cancelamento.')
      return
    }
    setErro(null)
    mutation.mutate()
  }

  return (
    <form onSubmit={handleSubmit} noValidate className="space-y-4">
      <DialogHeader>
        <DialogTitle>Cancelar evento</DialogTitle>
        <DialogDescription>
          O evento deixa de aceitar tarefas. Os acessos já registrados não são desfeitos.
        </DialogDescription>
      </DialogHeader>
      <div className="flex flex-col gap-1.5">
        <Label htmlFor="cancelar-motivo">Motivo do cancelamento</Label>
        <Input
          id="cancelar-motivo"
          value={motivo}
          onChange={(e) => setMotivo(e.target.value)}
          maxLength={500}
        />
      </div>
      {erro && (
        <p role="alert" className="text-sm text-destructive">
          {erro}
        </p>
      )}
      <DialogFooter>
        <Button type="button" variant="outline" onClick={onClose}>
          Voltar
        </Button>
        <Button type="submit" variant="destructive" loading={mutation.isPending}>
          Confirmar cancelamento
        </Button>
      </DialogFooter>
    </form>
  )
}
