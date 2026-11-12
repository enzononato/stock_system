import { useEffect, useState } from 'react'
import { keepPreviousData, useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { Pencil, Plus, RefreshCw, Search } from 'lucide-react'
import {
  getSyncStatus,
  listColaboradores,
  sincronizarColaboradores,
  type Colaborador,
  type SyncStatus,
} from '@/api/colaboradores'
import { listUnidades } from '@/api/unidades'
import { ColaboradorFormDialog } from '@/components/colaboradores/ColaboradorFormDialog'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Pagination } from '@/components/ui/Pagination'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { toast } from '@/components/ui/toast'
import { useAuth } from '@/contexts/AuthContext'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { usePageSize } from '@/hooks/usePageSize'
import { getErrorMessage } from '@/lib/api-error'
import { formatTelefone, upper } from '@/lib/format'
import type { StatusTone } from '@/lib/status'
import { formatCpf, formatDateTime } from '@/lib/utils'

const SITUACOES: { value: Colaborador['situacao']; label: string; tone: StatusTone }[] = [
  { value: 'ativo', label: 'Ativo', tone: 'available' },
  { value: 'afastado', label: 'Afastado', tone: 'pending' },
  { value: 'desligado', label: 'Desligado', tone: 'retired' },
]

const filtroClass =
  'h-9 rounded border border-input bg-surface-alt px-3 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring'

function plural(n: number, singular: string, pluralForm: string) {
  return `${n} ${n === 1 ? singular : pluralForm}`
}

function resumo(r: { criados: number | null; atualizados: number | null; desligados: number | null }) {
  const criados = r.criados ?? 0
  const atualizados = r.atualizados ?? 0
  const desligados = r.desligados ?? 0
  return `${plural(criados, 'criado', 'criados')}, ${plural(atualizados, 'atualizado', 'atualizados')}, ${plural(desligados, 'desligado', 'desligados')}`
}

function textoSincronizacao(status: SyncStatus | undefined): string | null {
  if (!status) return null
  if (!status.ultima) {
    return status.configurado
      ? 'Nenhuma sincronização realizada ainda.'
      : 'Sincronização com a base de RH não configurada.'
  }
  const u = status.ultima
  const quando = formatDateTime(u.concluido_em ?? u.iniciado_em)
  if (u.status === 'erro') return `Última sincronização em ${quando} falhou: ${u.erro ?? 'erro desconhecido'}`
  if (u.status === 'em_andamento') return `Última sincronização iniciada em ${quando}: em andamento.`
  return `Última sincronização em ${quando}: concluída (${resumo(u)}).`
}

export default function ColaboradoresPage() {
  const { user } = useAuth()
  const isGestor = user?.role === 'Gestor'
  const queryClient = useQueryClient()

  const [pageSize, setPageSize] = usePageSize('colaboradores')
  const [page, setPage] = useState(1)
  const [search, setSearch] = useState('')
  const [situacao, setSituacao] = useState('')
  const [unidadeId, setUnidadeId] = useState('')
  const busca = useDebouncedValue(search.trim(), 300)

  const [dialog, setDialog] = useState<{ open: boolean; colaborador: Colaborador | null }>({
    open: false,
    colaborador: null,
  })

  const { data, isLoading, isError, error, isFetching, refetch } = useQuery({
    queryKey: ['colaboradores', 'lista', { busca, situacao, unidadeId, page, pageSize }],
    queryFn: () =>
      listColaboradores({
        busca: busca || undefined,
        situacao: situacao || undefined,
        unidade_id: unidadeId ? Number(unidadeId) : undefined,
        limit: pageSize,
        offset: (page - 1) * pageSize,
      }),
    placeholderData: keepPreviousData,
  })

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades', { includeInactive: false }],
    queryFn: () => listUnidades(),
  })

  const { data: syncStatus } = useQuery({
    queryKey: ['colaboradores', 'sincronizacao'],
    queryFn: getSyncStatus,
    enabled: isGestor,
  })

  const sincronizar = useMutation({
    mutationFn: sincronizarColaboradores,
    onSuccess: (r) => {
      queryClient.invalidateQueries({ queryKey: ['colaboradores'] })
      toast(`Sincronização concluída: ${r.total_base} na base, ${resumo(r)}.`)
    },
    onError: (err: unknown) => {
      queryClient.invalidateQueries({ queryKey: ['colaboradores', 'sincronizacao'] })
      toast(getErrorMessage(err, 'Erro ao sincronizar colaboradores.'), 'error')
    },
  })

  const total = data?.total ?? 0
  const items = data?.items ?? []

  // Se a lista encolheu (filtro, sincronização) e a página atual deixou de existir, volta para a última.
  useEffect(() => {
    if (!data) return
    const lastPage = Math.max(1, Math.ceil(data.total / pageSize))
    if (page > lastPage) setPage(lastPage)
  }, [data, page, pageSize])

  const textoSync = isGestor ? textoSincronizacao(syncStatus) : null

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Pessoas"
        eyebrowDetail="Base de colaboradores"
        title="Colaboradores"
        description="Colaboradores sincronizados da base de RH e cadastros manuais, usados nos empréstimos e responsabilidades."
        actions={
          <>
            {isGestor && (
              <Button variant="outline" loading={sincronizar.isPending} onClick={() => sincronizar.mutate()}>
                <RefreshCw aria-hidden="true" />
                Sincronizar agora
              </Button>
            )}
            <Button onClick={() => setDialog({ open: true, colaborador: null })}>
              <Plus aria-hidden="true" />
              Novo colaborador
            </Button>
          </>
        }
      />

      {textoSync && (
        <p
          className={
            syncStatus?.ultima?.status === 'erro'
              ? 'text-body-sm text-destructive'
              : 'text-body-sm text-muted-foreground'
          }
        >
          {textoSync}
        </p>
      )}

      <div className="surface-panel flex flex-wrap items-end gap-3 p-3">
        <div className="relative min-w-[240px] flex-1">
          <Search
            size={16}
            aria-hidden="true"
            className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
          />
          <Input
            aria-label="Buscar colaborador"
            placeholder="Buscar colaborador por nome, CPF ou matrícula..."
            value={search}
            onChange={(e) => {
              setSearch(e.target.value)
              setPage(1)
            }}
            className="pl-10"
          />
        </div>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Situação
          <select
            aria-label="Situação"
            className={filtroClass}
            value={situacao}
            onChange={(e) => {
              setSituacao(e.target.value)
              setPage(1)
            }}
          >
            <option value="">Todas</option>
            {SITUACOES.map((s) => (
              <option key={s.value} value={s.value}>
                {s.label}
              </option>
            ))}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Unidade
          <select
            aria-label="Unidade"
            className={filtroClass}
            value={unidadeId}
            onChange={(e) => {
              setUnidadeId(e.target.value)
              setPage(1)
            }}
          >
            <option value="">Todas</option>
            {unidades.map((u) => (
              <option key={u.id} value={u.id}>
                {upper(u.nome)}
              </option>
            ))}
          </select>
        </label>
      </div>

      {isLoading ? (
        <LoadingState label="Carregando colaboradores..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : items.length === 0 ? (
        <EmptyState
          title="Nenhum colaborador encontrado"
          description="Ajuste a busca ou os filtros, ou cadastre um colaborador manualmente."
        />
      ) : (
        <div className="surface-panel overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="border-b border-border bg-surface-alt">
                <tr>
                  {['Nome', 'CPF', 'Cargo', 'Setor', 'Unidade', 'Situação', 'E-mail', 'Telefone', 'Superior', ''].map(
                    (h, i) => (
                      <th
                        key={h || i}
                        className="whitespace-nowrap px-4 py-2.5 text-left text-caption text-muted-foreground"
                      >
                        {h}
                      </th>
                    )
                  )}
                </tr>
              </thead>
              <tbody>
                {items.map((c) => {
                  const sit = SITUACOES.find((s) => s.value === c.situacao)
                  return (
                    <tr
                      key={c.id}
                      className="border-b border-border transition-colors duration-micro last:border-0 hover:bg-surface-alt"
                    >
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm font-medium">{upper(c.nome)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm tabular-nums">{formatCpf(c.cpf)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{c.cargo || '—'}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{c.setor || '—'}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">
                        {upper(c.unidade_nome ?? c.unidade_rh) || '—'}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5">
                        <StatusPill status={sit?.label ?? c.situacao} tone={sit?.tone ?? 'neutral'} />
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{c.email || '—'}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{formatTelefone(c.telefone) || '—'}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{upper(c.superior_nome) || '—'}</td>
                      <td className="whitespace-nowrap px-4 py-2.5">
                        <Button
                          size="sm"
                          variant="ghost"
                          aria-label={`Editar ${c.nome}`}
                          onClick={() => setDialog({ open: true, colaborador: c })}
                        >
                          <Pencil size={14} />
                        </Button>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
          <div className="border-t border-border px-4 py-3">
            <Pagination
              page={page}
              pageSize={pageSize}
              total={total}
              isFetching={isFetching}
              onPageChange={setPage}
              onPageSizeChange={(n) => {
                setPageSize(n)
                setPage(1)
              }}
            />
          </div>
        </div>
      )}

      <ColaboradorFormDialog
        open={dialog.open}
        onOpenChange={(open) => setDialog((d) => ({ ...d, open }))}
        colaborador={dialog.colaborador}
      />
    </div>
  )
}
