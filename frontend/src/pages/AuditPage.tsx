import { useEffect, useState } from 'react'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import { listAudit, type AuditEntry } from '@/api/audit'
import { PageHeader } from '@/components/layout/PageHeader'
import { Input } from '@/components/ui/input'
import { Pagination } from '@/components/ui/Pagination'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { usePageSize } from '@/hooks/usePageSize'
import { formatDateTime } from '@/lib/utils'

/** Entidades auditadas mais comuns, só como sugestão: o campo aceita qualquer valor. */
const ENTIDADES_SUGERIDAS = ['item', 'unidade', 'local', 'colaborador', 'usuario', 'termo']

/**
 * O backend manda `ocorrido_em` em UTC, sem offset. Sem o `Z`, o navegador
 * leria como horário local e a hora exibida sairia deslocada.
 */
function emHorarioLocal(utc: string): string {
  const temOffset = /(Z|[+-]\d{2}:?\d{2})$/i.test(utc)
  return formatDateTime(temOffset ? utc : `${utc}Z`)
}

function formatarValor(v: unknown): string {
  if (v === null || v === undefined) return '—'
  if (typeof v === 'boolean') return v ? 'sim' : 'não'
  if (typeof v === 'object') return JSON.stringify(v)
  return String(v)
}

/** "campo: antes → depois" para cada chave presente em `antes` ou `depois` (o backend só envia as alteradas). */
function resumo(e: AuditEntry): string[] {
  const antes = e.antes ?? {}
  const depois = e.depois ?? {}
  const campos = [...new Set([...Object.keys(antes), ...Object.keys(depois)])]
  return campos.map((c) => `${c}: ${formatarValor(antes[c])} → ${formatarValor(depois[c])}`)
}

const filtroClass =
  'h-9 rounded border border-input bg-surface-alt px-3 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring'

export default function AuditPage() {
  const [pageSize, setPageSize] = usePageSize('auditoria')
  const [entidade, setEntidade] = useState('')
  const [usuario, setUsuario] = useState('')
  const [de, setDe] = useState('')
  const [ate, setAte] = useState('')
  const entidadeAplicada = useDebouncedValue(entidade.trim(), 300)
  const usernameAplicado = useDebouncedValue(usuario.trim(), 300)

  // A página vale só para o conjunto de filtros em que foi escolhida: quando
  // qualquer filtro aplicado (ou o tamanho da página) muda, volta para a 1 na
  // mesma renderização, sem disparar uma requisição com filtro novo e página antiga.
  const filtroKey = JSON.stringify([entidadeAplicada, usernameAplicado, de, ate, pageSize])
  const [paginaDoFiltro, setPaginaDoFiltro] = useState({ page: 1, key: filtroKey })
  if (paginaDoFiltro.key !== filtroKey) setPaginaDoFiltro({ page: 1, key: filtroKey })
  const page = paginaDoFiltro.key === filtroKey ? paginaDoFiltro.page : 1
  const setPage = (p: number) => setPaginaDoFiltro({ page: p, key: filtroKey })

  const { data, isLoading, isError, error, isFetching, refetch } = useQuery({
    queryKey: ['audit', { entidade: entidadeAplicada, username: usernameAplicado, de, ate, page, pageSize }],
    queryFn: () =>
      listAudit({
        entidade: entidadeAplicada || undefined,
        username: usernameAplicado || undefined,
        de: de || undefined,
        ate: ate || undefined,
        limit: pageSize,
        offset: (page - 1) * pageSize,
      }),
    placeholderData: keepPreviousData,
  })

  const total = data?.total ?? 0
  const items = data?.items ?? []

  // Se a lista encolheu e a página atual deixou de existir, volta para a última.
  useEffect(() => {
    if (!data) return
    const lastPage = Math.max(1, Math.ceil(data.total / pageSize))
    if (page > lastPage) setPage(lastPage)
  }, [data, page, pageSize])

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Gestão"
        eyebrowDetail="Trilha de auditoria"
        title="Auditoria"
        description="Quem alterou o quê e quando. Os horários são exibidos no fuso do seu navegador."
      />

      <div className="surface-panel flex flex-wrap items-end gap-3 p-3">
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Entidade
          <Input
            aria-label="Entidade"
            list="auditoria-entidades"
            placeholder="Ex.: item"
            value={entidade}
            onChange={(e) => setEntidade(e.target.value)}
            className="w-44"
          />
          <datalist id="auditoria-entidades">
            {ENTIDADES_SUGERIDAS.map((e) => (
              <option key={e} value={e} />
            ))}
          </datalist>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Usuário
          <Input
            aria-label="Usuário"
            placeholder="Login do usuário"
            value={usuario}
            onChange={(e) => setUsuario(e.target.value)}
            className="w-48"
          />
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          De
          <input
            type="date"
            aria-label="De"
            className={filtroClass}
            value={de}
            max={ate || undefined}
            onChange={(e) => setDe(e.target.value)}
          />
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Até
          <input
            type="date"
            aria-label="Até"
            className={filtroClass}
            value={ate}
            min={de || undefined}
            onChange={(e) => setAte(e.target.value)}
          />
        </label>
      </div>

      {isLoading ? (
        <LoadingState label="Carregando auditoria..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : items.length === 0 ? (
        <EmptyState
          title="Nenhum registro encontrado"
          description="Ajuste os filtros ou o período para ver outros registros."
        />
      ) : (
        <div className="surface-panel overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="border-b border-border bg-surface-alt">
                <tr>
                  {['Data/hora', 'Usuário', 'Entidade', 'Ação', 'Alterações'].map((h) => (
                    <th key={h} className="whitespace-nowrap px-4 py-2.5 text-left text-caption text-muted-foreground">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {items.map((e) => {
                  const linhas = resumo(e)
                  return (
                    <tr
                      key={e.id}
                      className="border-b border-border align-top transition-colors duration-micro last:border-0 hover:bg-surface-alt"
                    >
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm tabular-nums">
                        {emHorarioLocal(e.ocorrido_em)}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">
                        {e.username || '—'}
                        {e.papel && <span className="ml-1 text-caption text-muted-foreground">({e.papel})</span>}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">
                        {e.entidade_id ? `${e.entidade} #${e.entidade_id}` : e.entidade}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{e.acao}</td>
                      <td className="px-4 py-2.5 text-body-sm">
                        {linhas.length === 0 ? (
                          <span className="text-muted-foreground">—</span>
                        ) : (
                          <ul className="space-y-0.5">
                            {linhas.map((l) => (
                              <li key={l} className="break-words">
                                {l}
                              </li>
                            ))}
                          </ul>
                        )}
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
    </div>
  )
}
