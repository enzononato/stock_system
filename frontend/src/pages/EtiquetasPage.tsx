import { useEffect, useMemo, useState } from 'react'
import { keepPreviousData, useMutation, useQuery } from '@tanstack/react-query'
import { Printer, Search } from 'lucide-react'
import { listItemsPaginated } from '@/api/items'
import { gerarEtiquetas, listIdsParaEtiquetas, type Etiqueta } from '@/api/etiquetas'
import { listUnidades } from '@/api/unidades'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { Checkbox } from '@/components/ui/checkbox'
import { Input } from '@/components/ui/input'
import { Pagination } from '@/components/ui/Pagination'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { toast } from '@/components/ui/toast'
import { FolhasEtiquetas } from '@/features/etiquetas/FolhasEtiquetas'
import { ImpressaoEtiquetas } from '@/features/etiquetas/ImpressaoEtiquetas'
import { LayoutForm } from '@/features/etiquetas/LayoutForm'
import {
  LAYOUT_PADRAO,
  carregarLayout,
  salvarLayout,
  totalFolhas,
  validarLayout,
  validarPosicaoInicial,
  type LayoutEtiquetas,
} from '@/features/etiquetas/layout'
import { gerarQrs } from '@/features/etiquetas/qr'
import { useConstants } from '@/hooks/useConstants'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { usePageSize } from '@/hooks/usePageSize'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'

const STATUS_OPCOES = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução', 'Em Manutenção']
const LIMITE_ETIQUETAS = 2000
const MSG_LIMITE = `Selecione no máximo ${LIMITE_ETIQUETAS} equipamentos por vez.`

const filtroClass =
  'h-9 rounded border border-input bg-surface-alt px-3 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring'

interface Gerado {
  baseUrl: string
  etiquetas: Etiqueta[]
  qrs: Record<number, string>
}

function origemDiferente(baseUrl: string): boolean {
  try {
    return new URL(baseUrl).origin !== window.location.origin
  } catch {
    return true
  }
}

/** Seleção de equipamentos e impressão de etiquetas com QR code em folha A4. */
export default function EtiquetasPage() {
  const { equipmentTypes } = useConstants()
  const [pageSize, setPageSize] = usePageSize('etiquetas')
  const [revenda, setRevenda] = useState('')
  const [tipo, setTipo] = useState('')
  const [status, setStatus] = useState('')
  const [search, setSearch] = useState('')
  const busca = useDebouncedValue(search.trim(), 300)

  // A página vale só para o conjunto de filtros em que foi escolhida; mudar qualquer filtro volta à 1.
  const chaveFiltros = JSON.stringify([revenda, tipo, status, busca])
  const [paginaDosFiltros, setPaginaDosFiltros] = useState({ page: 1, chave: chaveFiltros })
  if (paginaDosFiltros.chave !== chaveFiltros) setPaginaDosFiltros({ page: 1, chave: chaveFiltros })
  const page = paginaDosFiltros.chave === chaveFiltros ? paginaDosFiltros.page : 1
  const setPage = (p: number) => setPaginaDosFiltros({ page: p, chave: chaveFiltros })

  const [selecionados, setSelecionados] = useState<Set<number>>(new Set())
  const [gerado, setGerado] = useState<Gerado | null>(null)
  const [layout, setLayout] = useState<LayoutEtiquetas>(() => carregarLayout())
  const [posicaoInicial, setPosicaoInicial] = useState(1)

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades', { includeInactive: false }],
    queryFn: () => listUnidades(),
  })

  const { data, isLoading, isError, error, isFetching, refetch } = useQuery({
    queryKey: ['etiquetas', 'itens', { revenda, tipo, status, busca, page, pageSize }],
    queryFn: () =>
      listItemsPaginated({
        revenda: revenda || undefined,
        tipo: tipo || undefined,
        status: status || undefined,
        search: busca || undefined,
        limit: pageSize,
        offset: (page - 1) * pageSize,
      }),
    placeholderData: keepPreviousData,
  })
  const itens = data?.items ?? []
  const total = data?.total ?? 0

  // Se a lista encolheu e a página atual deixou de existir, volta para a última.
  useEffect(() => {
    if (!data) return
    const ultima = Math.max(1, Math.ceil(data.total / pageSize))
    if (page > ultima) setPage(ultima)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, page, pageSize])

  const paginaSelecionada = itens.length > 0 && itens.every((i) => selecionados.has(i.id))

  function alternarPagina(marcar: boolean) {
    setSelecionados((atual) => {
      const novo = new Set(atual)
      itens.forEach((i) => (marcar ? novo.add(i.id) : novo.delete(i.id)))
      return novo
    })
  }

  function alternarItem(id: number, marcar: boolean) {
    setSelecionados((atual) => {
      const novo = new Set(atual)
      if (marcar) novo.add(id)
      else novo.delete(id)
      return novo
    })
  }

  const selecionarTodos = useMutation({
    mutationFn: () =>
      listIdsParaEtiquetas({
        revenda,
        tipo: tipo || undefined,
        status: status || undefined,
        search: busca || undefined,
      }),
    onSuccess: ({ ids, total: n }) => {
      setSelecionados((atual) => new Set([...atual, ...ids]))
      toast.success(`${n} equipamento(s) selecionado(s).`)
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, 'Erro ao selecionar os equipamentos da unidade.')),
  })

  const gerar = useMutation({
    mutationFn: async (ids: number[]) => {
      const resposta = await gerarEtiquetas(ids)
      const qrs = await gerarQrs(resposta.etiquetas)
      return { resposta, qrs }
    },
    onSuccess: ({ resposta, qrs }) => {
      setGerado({ baseUrl: resposta.base_url, etiquetas: resposta.etiquetas, qrs })
      setPosicaoInicial(1)
      if (resposta.ignorados.length > 0) {
        toast.error(`${resposta.ignorados.length} equipamento(s) ignorado(s): baixados ou inexistentes.`)
      }
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, 'Erro ao gerar as etiquetas.')),
  })

  function aoGerar() {
    if (selecionados.size > LIMITE_ETIQUETAS) {
      toast.error(MSG_LIMITE)
      return
    }
    gerar.mutate([...selecionados].sort((a, b) => a - b))
  }

  function alterarLayout(novo: LayoutEtiquetas) {
    setLayout(novo)
    if (validarLayout(novo) === null) salvarLayout(novo)
  }

  function restaurarPadrao() {
    alterarLayout(LAYOUT_PADRAO)
    setPosicaoInicial(1)
  }

  const erroLayout = validarLayout(layout)
  const erroPosicao = erroLayout ? null : validarPosicaoInicial(layout, posicaoInicial)
  const erro = erroLayout ?? erroPosicao
  const folhas = useMemo(
    () => (gerado && !erro ? totalFolhas(layout, gerado.etiquetas.length, posicaoInicial) : 0),
    [gerado, erro, layout, posicaoInicial]
  )

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Inventário"
        eyebrowDetail="Etiquetas com QR code"
        title="Etiquetas"
        description="Selecione os equipamentos e imprima etiquetas com QR code em folha A4. O QR abre o equipamento no sistema."
      />

      <div className="surface-panel flex flex-wrap items-end gap-3 p-3">
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Unidade
          <select
            aria-label="Unidade"
            className={`${filtroClass} min-w-[200px]`}
            value={revenda}
            onChange={(e) => setRevenda(e.target.value)}
          >
            <option value="">Todas</option>
            {unidades.map((u) => (
              <option key={u.id} value={u.nome}>
                {upper(u.nome)}
              </option>
            ))}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Tipo
          <select aria-label="Tipo" className={filtroClass} value={tipo} onChange={(e) => setTipo(e.target.value)}>
            <option value="">Todos</option>
            {equipmentTypes.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Status
          <select aria-label="Status" className={filtroClass} value={status} onChange={(e) => setStatus(e.target.value)}>
            <option value="">Todos</option>
            {STATUS_OPCOES.map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
        </label>
        <div className="relative min-w-[220px] flex-1">
          <Search
            size={16}
            aria-hidden="true"
            className="pointer-events-none absolute left-3.5 top-1/2 -translate-y-1/2 text-muted-foreground"
          />
          <Input
            aria-label="Buscar"
            placeholder="Buscar por marca, modelo, série ou usuário..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="pl-10"
          />
        </div>
      </div>

      <div className="surface-panel flex flex-wrap items-center gap-3 p-3">
        <span className="text-body-sm font-medium" aria-live="polite">{`${selecionados.size} selecionado(s)`}</span>
        <Button
          variant="outline"
          size="sm"
          disabled={selecionados.size === 0}
          onClick={() => setSelecionados(new Set())}
        >
          Limpar seleção
        </Button>
        <Button
          variant="outline"
          size="sm"
          disabled={!revenda}
          loading={selecionarTodos.isPending}
          onClick={() => selecionarTodos.mutate()}
        >
          Selecionar todos da unidade
        </Button>
        <div className="ml-auto">
          <Button loading={gerar.isPending} disabled={selecionados.size === 0} onClick={aoGerar}>
            Gerar etiquetas
          </Button>
        </div>
      </div>

      {isLoading ? (
        <LoadingState label="Carregando equipamentos..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : itens.length === 0 ? (
        <EmptyState title="Nenhum equipamento encontrado" description="Ajuste os filtros para listar equipamentos." />
      ) : (
        <div className="surface-panel overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="border-b border-border bg-surface-alt">
                <tr>
                  <th className="w-10 px-4 py-2.5 text-left">
                    <Checkbox
                      aria-label="Selecionar a página"
                      checked={paginaSelecionada}
                      onCheckedChange={alternarPagina}
                    />
                  </th>
                  {['Patrimônio', 'Tipo', 'Equipamento', 'Nº de série', 'Unidade', 'Status'].map((h) => (
                    <th key={h} className="whitespace-nowrap px-4 py-2.5 text-left text-caption text-muted-foreground">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {itens.map((item) => (
                  <tr
                    key={item.id}
                    className="border-b border-border transition-colors duration-micro last:border-0 hover:bg-surface-alt"
                  >
                    <td className="px-4 py-2.5">
                      <Checkbox
                        aria-label={`Selecionar #${item.id}`}
                        checked={selecionados.has(item.id)}
                        onCheckedChange={(marcar) => alternarItem(item.id, marcar)}
                      />
                    </td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm font-mono tabular-nums">#{item.id}</td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{item.tipo || '—'}</td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">
                      {[item.brand, item.model].filter(Boolean).join(' ') || '—'}
                    </td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{item.identificador || '—'}</td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{upper(item.revenda) || '—'}</td>
                    <td className="whitespace-nowrap px-4 py-2.5">
                      <StatusPill status={item.status} />
                    </td>
                  </tr>
                ))}
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

      {gerado && (
        <section aria-labelledby="titulo-folha" className="surface-panel space-y-4 p-4">
          <h2 id="titulo-folha" className="text-heading-sm font-semibold">
            Folha de etiquetas
          </h2>

          <LayoutForm
            layout={layout}
            posicaoInicial={posicaoInicial}
            onLayoutChange={alterarLayout}
            onPosicaoChange={setPosicaoInicial}
            onRestaurar={restaurarPadrao}
          />

          {erro && (
            <p role="alert" className="text-body-sm text-destructive">
              {erro}
            </p>
          )}

          {origemDiferente(gerado.baseUrl) && (
            <p className="text-body-sm text-muted-foreground">
              {`Atenção: os QR codes vão apontar para ${gerado.baseUrl}, diferente do endereço desta página (${window.location.origin}). Confira APP_PUBLIC_URL no servidor antes de imprimir.`}
            </p>
          )}

          <div className="flex flex-wrap items-center gap-3">
            <span className="text-body-sm font-medium">
              {erro ? '' : `${gerado.etiquetas.length} etiqueta(s) em ${folhas} folha(s)`}
            </span>
            <Button disabled={!!erro} onClick={() => window.print()}>
              <Printer aria-hidden="true" />
              Imprimir
            </Button>
          </div>
          <p className="text-body-sm text-muted-foreground">
            Na janela de impressão, escolha tamanho real (100%) e margens &quot;Nenhuma&quot;.
          </p>

          {!erro && (
            <>
              <div
                className="max-h-[70vh] overflow-auto rounded border border-border bg-surface-alt p-2"
                role="group"
                aria-label="Prévia da folha"
              >
                <div
                  style={{
                    width: '105mm',
                    height: `calc(${folhas * 297 * 0.5}mm + ${(folhas - 1) * 8}px)`,
                  }}
                >
                  <div
                    className="flex origin-top-left scale-50 flex-col gap-4 [&_.folha]:border [&_.folha]:border-border"
                    style={{ width: '210mm' }}
                  >
                    <FolhasEtiquetas
                      etiquetas={gerado.etiquetas}
                      qrs={gerado.qrs}
                      layout={layout}
                      posicaoInicial={posicaoInicial}
                    />
                  </div>
                </div>
              </div>
              <ImpressaoEtiquetas>
                <FolhasEtiquetas
                  etiquetas={gerado.etiquetas}
                  qrs={gerado.qrs}
                  layout={layout}
                  posicaoInicial={posicaoInicial}
                />
              </ImpressaoEtiquetas>
            </>
          )}
        </section>
      )}
    </div>
  )
}
