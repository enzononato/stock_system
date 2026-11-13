import { useEffect, useState } from 'react'
import { Link, useNavigate } from 'react-router-dom'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import { Plus } from 'lucide-react'
import { listEventos, type EventoResumo, type StatusEvento, type TipoEvento } from '@/api/acessos'
import { EventoFormDialog } from '@/components/acessos/EventoFormDialog'
import { ROTULO_TIPO_EVENTO, STATUS_EVENTO } from '@/components/acessos/rotulos'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Pagination } from '@/components/ui/Pagination'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { usePageSize } from '@/hooks/usePageSize'
import { upper } from '@/lib/format'
import { formatDateTime } from '@/lib/utils'

type FiltroSituacao = StatusEvento | 'todos'

const selectClass =
  'h-9 min-w-[160px] rounded border border-input bg-surface-alt px-3 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring'

function cargoDoEvento(e: EventoResumo): string {
  if (e.tipo === 'mudanca_funcao') return `${e.cargo_anterior ?? '—'} → ${e.cargo_novo ?? '—'}`
  return e.cargo_novo ?? e.cargo_anterior ?? '—'
}

/** Eventos de acesso (admissão, desligamento, mudança de função) — Gestor e Técnico. */
export default function AcessoEventosPage() {
  const navigate = useNavigate()
  const [pageSize, setStoredPageSize] = usePageSize('acessos-eventos')
  const [page, setPage] = useState(1)
  const [situacao, setSituacao] = useState<FiltroSituacao>('aberto')
  const [tipo, setTipo] = useState<TipoEvento | ''>('')
  const [busca, setBusca] = useState('')
  const [formAberto, setFormAberto] = useState(false)
  const buscaAplicada = useDebouncedValue(busca.trim(), 300)

  // Qualquer mudança de filtro ou de tamanho volta à página 1 no mesmo evento,
  // para nunca pedir a página antiga com o filtro novo.
  const mudarSituacao = (v: FiltroSituacao) => {
    setSituacao(v)
    setPage(1)
  }
  const mudarTipo = (v: TipoEvento | '') => {
    setTipo(v)
    setPage(1)
  }
  const mudarBusca = (v: string) => {
    setBusca(v)
    setPage(1)
  }
  const mudarTamanho = (n: number) => {
    setStoredPageSize(n)
    setPage(1)
  }

  const params = {
    ...(situacao !== 'todos' ? { status: situacao } : {}),
    ...(tipo ? { tipo } : {}),
    ...(buscaAplicada ? { busca: buscaAplicada } : {}),
    limit: pageSize,
    offset: (page - 1) * pageSize,
  }

  const { data, isLoading, isError, error, refetch, isFetching } = useQuery({
    queryKey: ['acessos', 'eventos', params],
    queryFn: () => listEventos(params),
    placeholderData: keepPreviousData,
  })

  const eventos = data?.items ?? []
  const total = data?.total ?? 0
  const lastPage = Math.max(1, Math.ceil(total / pageSize))
  useEffect(() => {
    if (data && page > lastPage) setPage(lastPage)
  }, [data, page, lastPage])

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Acessos"
        title="Eventos de acesso"
        description="Admissões, desligamentos e mudanças de função com as tarefas de criar e desativar acessos."
        actions={
          <Button onClick={() => setFormAberto(true)}>
            <Plus aria-hidden="true" />
            Registrar evento
          </Button>
        }
      />

      <div className="surface-panel flex flex-wrap items-end gap-3 p-3">
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Situação
          <select
            aria-label="Situação"
            className={selectClass}
            value={situacao}
            onChange={(e) => mudarSituacao(e.target.value as FiltroSituacao)}
          >
            <option value="aberto">Abertos</option>
            <option value="concluido">Concluídos</option>
            <option value="cancelado">Cancelados</option>
            <option value="todos">Todos</option>
          </select>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Tipo
          <select
            aria-label="Tipo"
            className={selectClass}
            value={tipo}
            onChange={(e) => mudarTipo(e.target.value as TipoEvento | '')}
          >
            <option value="">Todos</option>
            {(Object.keys(ROTULO_TIPO_EVENTO) as TipoEvento[]).map((t) => (
              <option key={t} value={t}>
                {ROTULO_TIPO_EVENTO[t]}
              </option>
            ))}
          </select>
        </label>
        <label className="flex min-w-[220px] flex-1 flex-col gap-1 text-caption text-muted-foreground">
          Buscar por nome ou CPF
          <Input
            aria-label="Buscar por nome ou CPF"
            value={busca}
            onChange={(e) => mudarBusca(e.target.value)}
            placeholder="Nome ou CPF do colaborador"
          />
        </label>
      </div>

      {isLoading ? (
        <LoadingState label="Carregando eventos..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : total === 0 ? (
        <EmptyState title="Nenhum evento encontrado." description="Ajuste os filtros ou registre um evento." />
      ) : (
        <div className="surface-panel overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="border-b border-border bg-surface-alt">
                <tr>
                  {['Colaborador', 'Tipo', 'Cargo', 'Situação', 'Tarefas', 'Aberto em', ''].map((h, i) => (
                    <th key={h || i} className="whitespace-nowrap px-4 py-2.5 text-left text-caption text-muted-foreground">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {eventos.map((e) => {
                  const situacaoEvento = STATUS_EVENTO[e.status]
                  return (
                    <tr key={e.id} className="border-b border-border last:border-0 hover:bg-surface-alt">
                      <td className="px-4 py-2.5 text-body-sm font-medium">{upper(e.colaborador_nome)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{ROTULO_TIPO_EVENTO[e.tipo]}</td>
                      <td className="px-4 py-2.5 text-body-sm">{cargoDoEvento(e)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5">
                        <StatusPill status={situacaoEvento.label} tone={situacaoEvento.tone} />
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm tabular-nums">
                        {`${e.total_tarefas - e.tarefas_pendentes}/${e.total_tarefas} resolvidas`}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{formatDateTime(e.aberto_em)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5">
                        <Link
                          to={`/acessos/eventos/${e.id}`}
                          aria-label={`Abrir evento #${e.id}`}
                          className="rounded-sm text-sm font-semibold text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring"
                        >
                          Abrir
                        </Link>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
          <div className="border-t border-border px-4 py-3">
            <Pagination
              page={Math.min(page, lastPage)}
              pageSize={pageSize}
              total={total}
              onPageChange={setPage}
              onPageSizeChange={mudarTamanho}
              isFetching={isFetching}
            />
          </div>
        </div>
      )}

      <EventoFormDialog
        open={formAberto}
        onOpenChange={setFormAberto}
        onCreated={(evento) => navigate(`/acessos/eventos/${evento.id}`)}
      />
    </div>
  )
}
