import { useEffect, useState } from 'react'
import { Link, useSearchParams } from 'react-router-dom'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import { BarChart3, Plus, Search } from 'lucide-react'
import { listLinhas, OPERADORAS, type Linha, type StatusLinha } from '@/api/linhas'
import { listUnidades } from '@/api/unidades'
import { LinhaFormDialog } from '@/components/linhas/LinhaFormDialog'
import { LinhaHistoricoDialog } from '@/components/linhas/LinhaHistoricoDialog'
import { LinhaStatusDialog } from '@/components/linhas/LinhaStatusDialog'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Pagination } from '@/components/ui/Pagination'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { useAuth } from '@/contexts/AuthContext'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { usePageSize } from '@/hooks/usePageSize'
import { descreverAparelho } from '@/features/linhas/format'
import {
  acoesDisponiveis,
  STATUS_LINHA_OPCOES,
  statusLinhaInfo,
  type AcaoStatusLinha,
} from '@/features/linhas/status'
import { formatMoeda, formatTelefone, upper } from '@/lib/format'

const filtroClass =
  'h-9 rounded border border-input bg-surface-alt px-3 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring'

const ROTULO_ACAO: Record<AcaoStatusLinha, string> = {
  suspender: 'Suspender',
  reativar: 'Reativar',
  cancelar: 'Cancelar',
}

const linkClass =
  'rounded-sm text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2'

export default function LinhasPage() {
  const { user } = useAuth()
  const podeEscrever = user?.role === 'Gestor' || user?.role === 'Técnico'
  const [searchParams] = useSearchParams()

  const [pageSize, setStoredPageSize] = usePageSize('linhas')
  const [search, setSearch] = useState(() => searchParams.get('busca') ?? '')
  const [status, setStatus] = useState('')
  const [unidadeId, setUnidadeId] = useState('')
  const [operadora, setOperadora] = useState('')
  const busca = useDebouncedValue(search.trim(), 300)
  // A página vale só para o termo (já com debounce) em que foi escolhida: quando a busca
  // aplicada muda, volta para a 1 na mesma renderização, sem requisição com a página antiga.
  const [paginaDaBusca, setPaginaDaBusca] = useState({ page: 1, busca })
  if (paginaDaBusca.busca !== busca) setPaginaDaBusca({ page: 1, busca })
  const page = paginaDaBusca.busca === busca ? paginaDaBusca.page : 1
  const setPage = (p: number) => setPaginaDaBusca({ page: p, busca })

  const [formDialog, setFormDialog] = useState<{ open: boolean; linha: Linha | null }>({ open: false, linha: null })
  const [historico, setHistorico] = useState<{ open: boolean; linha: Linha | null }>({ open: false, linha: null })
  const [statusDialog, setStatusDialog] = useState<{ open: boolean; linha: Linha | null; acao: AcaoStatusLinha }>({
    open: false,
    linha: null,
    acao: 'suspender',
  })

  const { data, isLoading, isError, error, isFetching, refetch } = useQuery({
    queryKey: ['linhas', 'lista', { busca, status, unidadeId, operadora, page, pageSize }],
    queryFn: () =>
      listLinhas({
        busca: busca || undefined,
        status: (status || undefined) as StatusLinha | undefined,
        unidade_id: unidadeId ? Number(unidadeId) : undefined,
        operadora: operadora || undefined,
        limit: pageSize,
        offset: (page - 1) * pageSize,
      }),
    placeholderData: keepPreviousData,
  })

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades', { includeInactive: true }],
    queryFn: () => listUnidades({ include_inactive: true }),
  })

  const total = data?.total ?? 0
  const items = data?.items ?? []

  // Se o total encolheu e a página atual deixou de existir, volta para a última.
  useEffect(() => {
    if (!data) return
    const lastPage = Math.max(1, Math.ceil(data.total / pageSize))
    if (page > lastPage) setPage(lastPage)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, page, pageSize])

  // Trocar filtro ou tamanho da página volta para a primeira no mesmo evento.
  const aoMudarFiltro = (aplicar: () => void) => {
    aplicar()
    setPage(1)
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Telefonia"
        title="Linhas telefônicas"
        description="Linhas com operadora, plano, custo mensal e vínculo a aparelho e colaborador. Trocas e mudanças de status ficam no histórico."
        actions={
          podeEscrever && (
            <>
              <Button asChild variant="outline">
                <Link to="/linhas/relatorio">
                  <BarChart3 aria-hidden="true" />
                  Relatório de custo
                </Link>
              </Button>
              <Button onClick={() => setFormDialog({ open: true, linha: null })}>
                <Plus aria-hidden="true" />
                Nova linha
              </Button>
            </>
          )
        }
      />

      <div className="surface-panel flex flex-wrap items-end gap-3 p-3">
        <div className="relative min-w-[240px] flex-1">
          <Search
            size={16}
            aria-hidden="true"
            className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
          />
          <Input
            aria-label="Buscar linha"
            placeholder="Buscar por número, plano, ICCID ou colaborador..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="pl-10"
          />
        </div>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Status
          <select
            aria-label="Status"
            className={filtroClass}
            value={status}
            onChange={(e) => aoMudarFiltro(() => setStatus(e.target.value))}
          >
            <option value="">Todos</option>
            {STATUS_LINHA_OPCOES.map((s) => (
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
            onChange={(e) => aoMudarFiltro(() => setUnidadeId(e.target.value))}
          >
            <option value="">Todas</option>
            {unidades.map((u) => (
              <option key={u.id} value={u.id}>
                {upper(u.nome)}
              </option>
            ))}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Operadora
          <select
            aria-label="Operadora"
            className={filtroClass}
            value={operadora}
            onChange={(e) => aoMudarFiltro(() => setOperadora(e.target.value))}
          >
            <option value="">Todas</option>
            {OPERADORAS.map((o) => (
              <option key={o} value={o}>
                {o}
              </option>
            ))}
          </select>
        </label>
      </div>

      {isLoading ? (
        <LoadingState label="Carregando linhas..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : items.length === 0 ? (
        <EmptyState
          title="Nenhuma linha encontrada"
          description="Ajuste a busca ou os filtros, ou cadastre uma nova linha."
        />
      ) : (
        <div className="surface-panel overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="border-b border-border bg-surface-alt">
                <tr>
                  {['Número', 'Operadora', 'Plano', 'Custo mensal', 'Status', 'Unidade', 'Aparelho', 'Colaborador', 'Ações'].map(
                    (h) => (
                      <th key={h} className="whitespace-nowrap px-4 py-2.5 text-left text-caption text-muted-foreground">
                        {h}
                      </th>
                    )
                  )}
                </tr>
              </thead>
              <tbody>
                {items.map((l) => {
                  const info = statusLinhaInfo(l.status)
                  const numero = formatTelefone(l.numero)
                  const aparelho = descreverAparelho({
                    tipo: l.item_tipo,
                    brand: l.item_brand,
                    model: l.item_model,
                    identificador: l.item_identificador,
                  })
                  return (
                    <tr
                      key={l.id}
                      className="border-b border-border transition-colors duration-micro last:border-0 hover:bg-surface-alt"
                    >
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm font-medium tabular-nums">{numero}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{l.operadora}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{l.plano || '—'}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm tabular-nums">
                        {formatMoeda(l.custo_mensal) || '—'}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5">
                        <StatusPill status={info.label} tone={info.tone} />
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{upper(l.unidade_nome) || '—'}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">
                        {l.item_id !== null ? (
                          <>
                            <Link to={`/items/${l.item_id}`} className={linkClass}>
                              {aparelho || `Equipamento ${l.item_id}`}
                            </Link>
                            {l.item_ativo === false && <span className="text-muted-foreground"> (baixado)</span>}
                          </>
                        ) : (
                          '—'
                        )}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">
                        {l.colaborador_id !== null && l.colaborador_nome ? (
                          <>
                            {podeEscrever ? (
                              <Link to={`/colaboradores/${l.colaborador_id}`} className={linkClass}>
                                {upper(l.colaborador_nome)}
                              </Link>
                            ) : (
                              <span>{upper(l.colaborador_nome)}</span>
                            )}
                            {l.colaborador_situacao === 'desligado' && (
                              <span className="text-muted-foreground"> (desligado)</span>
                            )}
                          </>
                        ) : (
                          '—'
                        )}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5">
                        <div className="flex items-center gap-1">
                          <Button
                            size="sm"
                            variant="ghost"
                            aria-label={`Histórico ${numero}`}
                            onClick={() => setHistorico({ open: true, linha: l })}
                          >
                            Histórico
                          </Button>
                          {podeEscrever && l.status !== 'cancelada' && (
                            <Button
                              size="sm"
                              variant="ghost"
                              aria-label={`Editar ${numero}`}
                              onClick={() => setFormDialog({ open: true, linha: l })}
                            >
                              Editar
                            </Button>
                          )}
                          {podeEscrever &&
                            acoesDisponiveis(l.status).map((acao) => (
                              <Button
                                key={acao}
                                size="sm"
                                variant="ghost"
                                aria-label={`${ROTULO_ACAO[acao]} ${numero}`}
                                onClick={() => setStatusDialog({ open: true, linha: l, acao })}
                              >
                                {ROTULO_ACAO[acao]}
                              </Button>
                            ))}
                        </div>
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
                setStoredPageSize(n)
                setPage(1)
              }}
            />
          </div>
        </div>
      )}

      <LinhaFormDialog
        open={formDialog.open}
        onOpenChange={(open) => setFormDialog((d) => ({ ...d, open }))}
        linha={formDialog.linha}
      />
      <LinhaHistoricoDialog
        open={historico.open}
        onOpenChange={(open) => setHistorico((d) => ({ ...d, open }))}
        linha={historico.linha}
      />
      {statusDialog.linha && (
        <LinhaStatusDialog
          open={statusDialog.open}
          onOpenChange={(open) => setStatusDialog((d) => ({ ...d, open }))}
          linha={statusDialog.linha}
          acao={statusDialog.acao}
        />
      )}
    </div>
  )
}
