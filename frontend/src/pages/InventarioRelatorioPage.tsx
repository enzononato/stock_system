import { useMemo, useState } from 'react'
import { Link, useParams, useSearchParams } from 'react-router-dom'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { Download } from 'lucide-react'
import {
  cancelarInventario,
  encerrarInventario,
  exportarRelatorio,
  getRelatorio,
  type CategoriaRelatorio,
  type RelatorioLinha,
} from '@/api/inventarios'
import { PageHeader } from '@/components/layout/PageHeader'
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from '@/components/ui/alert-dialog'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { DetailTabs } from '@/components/ui/DetailTabs'
import { Pagination } from '@/components/ui/Pagination'
import { StatBlock } from '@/components/ui/StatBlock'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { toast } from '@/components/ui/toast'
import { useAuth } from '@/contexts/AuthContext'
import {
  ORDEM_CATEGORIAS,
  ROTULO_STATUS_INVENTARIO,
  rotuloCategoria,
} from '@/features/inventario/resultados'
import { usePageSize } from '@/hooks/usePageSize'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
import { formatDateTime } from '@/lib/utils'

const linkClass =
  'rounded-sm text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2'

const COLUNAS = [
  'Patrimônio',
  'Tipo',
  'Equipamento',
  'Nº de série',
  'Unidade do equipamento',
  'Local cadastrado',
  'Local encontrado',
  'Usuário cadastrado',
  'Usuário encontrado',
  'Lido por / em',
  'Detalhe / observação',
]

function celula(valor: string | null | undefined): string {
  return valor || '—'
}

/** Relatório de um inventário: indicadores, abas por resultado, encerramento, cancelamento e exportação. */
export default function InventarioRelatorioPage() {
  const { id: idParam } = useParams()
  const id = Number(idParam)
  const { user } = useAuth()
  const isGestor = user?.role === 'Gestor'
  const queryClient = useQueryClient()
  const [searchParams, setSearchParams] = useSearchParams()
  const [pageSize, setPageSize] = usePageSize('inventario-relatorio')

  const [confirmandoEncerrar, setConfirmandoEncerrar] = useState(false)
  const [cancelando, setCancelando] = useState(false)
  const [motivo, setMotivo] = useState('')
  const [exportando, setExportando] = useState<'xlsx' | 'pdf' | null>(null)

  const { data, isLoading, isError, error, refetch } = useQuery({
    queryKey: ['inventario-relatorio', id],
    queryFn: () => getRelatorio(id),
    enabled: Number.isInteger(id),
  })

  const linhas = data?.linhas
  const contagem = useMemo(() => {
    const c = Object.fromEntries(ORDEM_CATEGORIAS.map((cat) => [cat, 0])) as Record<CategoriaRelatorio, number>
    linhas?.forEach((l) => {
      c[l.categoria] += 1
    })
    return c
  }, [linhas])

  const abaParam = searchParams.get('aba')
  const abaPadrao: CategoriaRelatorio = ORDEM_CATEGORIAS.find((c) => contagem[c] > 0) ?? 'nao_encontrado'
  const aba: CategoriaRelatorio = ORDEM_CATEGORIAS.includes(abaParam as CategoriaRelatorio)
    ? (abaParam as CategoriaRelatorio)
    : abaPadrao

  // A página vale só para a aba em que foi escolhida: trocar de aba volta à 1.
  const [paginaDaAba, setPaginaDaAba] = useState({ page: 1, aba })
  if (paginaDaAba.aba !== aba) setPaginaDaAba({ page: 1, aba })
  const page = paginaDaAba.aba === aba ? paginaDaAba.page : 1
  const setPage = (p: number) => setPaginaDaAba({ page: p, aba })

  function recarregar() {
    void queryClient.invalidateQueries({ queryKey: ['inventario-relatorio', id] })
    void queryClient.invalidateQueries({ queryKey: ['inventario', id] })
    void queryClient.invalidateQueries({ queryKey: ['inventarios'] })
  }

  const encerrar = useMutation({
    mutationFn: () => encerrarInventario(id),
    onSuccess: () => {
      toast.success('Inventário encerrado.')
      recarregar()
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, 'Erro ao encerrar o inventário.')),
  })

  const cancelar = useMutation({
    mutationFn: () => cancelarInventario(id, motivo.trim()),
    onSuccess: () => {
      toast.success('Inventário cancelado.')
      setCancelando(false)
      setMotivo('')
      recarregar()
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, 'Erro ao cancelar o inventário.')),
  })

  async function exportar(formato: 'xlsx' | 'pdf') {
    setExportando(formato)
    try {
      await exportarRelatorio(id, formato)
    } catch (err) {
      toast.error(getErrorMessage(err, 'Erro ao exportar o relatório.'))
    } finally {
      setExportando(null)
    }
  }

  if (isLoading) return <LoadingState label="Carregando relatório..." />
  if (isError || !data) return <ErrorState error={error} onRetry={() => refetch()} />

  const { inventario, linhas: todas } = data
  const { resumo } = inventario
  const aberto = inventario.status === 'aberto'
  const daAba: RelatorioLinha[] = todas.filter((l) => l.categoria === aba)
  const fatia = daAba.slice((page - 1) * pageSize, page * pageSize)
  const abas = ORDEM_CATEGORIAS.map((c) => ({
    id: c,
    label: `${rotuloCategoria(c, inventario.status)} (${contagem[c]})`,
  }))

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Inventário"
        eyebrowDetail={upper(inventario.unidade_nome)}
        title={`Inventário #${inventario.id}`}
        description="Resultado da contagem física da unidade."
        actions={
          <>
            <Button variant="outline" loading={exportando === 'xlsx'} onClick={() => exportar('xlsx')}>
              <Download aria-hidden="true" />
              Exportar Excel
            </Button>
            <Button variant="outline" loading={exportando === 'pdf'} onClick={() => exportar('pdf')}>
              <Download aria-hidden="true" />
              Exportar PDF
            </Button>
          </>
        }
      />

      <section aria-label="Dados do inventário" className="surface-panel space-y-2 p-4 text-body-sm">
        <div className="flex flex-wrap items-center gap-3">
          <span className="font-medium">{upper(inventario.unidade_nome)}</span>
          <StatusPill status={ROTULO_STATUS_INVENTARIO[inventario.status]} />
        </div>
        <p className="text-muted-foreground">
          {`Aberto por ${inventario.iniciado_por} em ${formatDateTime(inventario.iniciado_em)}`}
        </p>
        {inventario.encerrado_em && (
          <p className="text-muted-foreground">
            {`${inventario.status === 'cancelado' ? 'Cancelado' : 'Encerrado'} por ${inventario.encerrado_por ?? '—'} em ${formatDateTime(inventario.encerrado_em)}`}
          </p>
        )}
        {inventario.observacao && <p>{`Observação: ${inventario.observacao}`}</p>}
        {inventario.motivo_cancelamento && <p>{`Motivo do cancelamento: ${inventario.motivo_cancelamento}`}</p>}
      </section>

      <section aria-label="Indicadores" className="grid grid-cols-2 gap-4 sm:grid-cols-4 lg:grid-cols-7">
        <StatBlock label="Esperados" value={resumo.esperados} />
        {ORDEM_CATEGORIAS.map((c) => (
          <StatBlock key={c} label={rotuloCategoria(c, inventario.status)} value={contagem[c]} />
        ))}
      </section>

      {aberto && (
        <div className="flex flex-wrap gap-3">
          <Button asChild>
            <Link to={`/inventarios/${id}/leitura`}>Continuar leitura</Link>
          </Button>
          <Button variant="outline" onClick={() => setConfirmandoEncerrar(true)}>
            Encerrar inventário
          </Button>
          {isGestor && (
            <Button variant="outline" onClick={() => setCancelando(true)}>
              Cancelar inventário
            </Button>
          )}
        </div>
      )}

      <DetailTabs
        tabs={abas}
        activeId={aba}
        label="Resultados do inventário"
        onSelect={(c) => setSearchParams({ aba: c }, { replace: true })}
      >
        {daAba.length === 0 ? (
          <EmptyState title="Nenhum equipamento nesta categoria" />
        ) : (
          <div className="surface-panel overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="border-b border-border bg-surface-alt">
                  <tr>
                    {COLUNAS.map((h) => (
                      <th key={h} className="whitespace-nowrap px-4 py-2.5 text-left text-caption text-muted-foreground">
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {fatia.map((l, i) => (
                    <tr
                      key={l.leitura_id ?? `n-${l.item_id}-${i}`}
                      className="border-b border-border transition-colors duration-micro last:border-0 hover:bg-surface-alt"
                    >
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm font-mono tabular-nums">
                        {l.item_id != null ? (
                          <Link to={`/items/${l.item_id}`} className={linkClass}>{`#${l.item_id}`}</Link>
                        ) : (
                          '—'
                        )}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{celula(l.tipo)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">
                        {[l.brand, l.model].filter(Boolean).join(' ') || '—'}
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{celula(l.identificador)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{upper(l.unidade_item) || '—'}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{celula(l.local_esperado)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{celula(l.local_encontrado)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{celula(l.usuario_esperado)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">{celula(l.usuario_encontrado)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm">
                        {l.lido_por ? `${l.lido_por} · ${formatDateTime(l.lido_em)}` : '—'}
                      </td>
                      <td className="min-w-[240px] px-4 py-2.5 text-body-sm">
                        {l.detalhe || l.observacao ? (
                          <div className="space-y-1">
                            {l.detalhe && <p>{l.detalhe}</p>}
                            {l.observacao && <p className="text-muted-foreground">{l.observacao}</p>}
                          </div>
                        ) : (
                          '—'
                        )}
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
                total={daAba.length}
                onPageChange={setPage}
                onPageSizeChange={(n) => {
                  setPageSize(n)
                  setPage(1)
                }}
              />
            </div>
          </div>
        )}
      </DetailTabs>

      <AlertDialog open={confirmandoEncerrar} onOpenChange={setConfirmandoEncerrar}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Encerrar inventário?</AlertDialogTitle>
            <AlertDialogDescription>
              {`${resumo.nao_encontrados} equipamento(s) ainda não lido(s) serão marcados como não encontrados. Depois de encerrado, o inventário não aceita novas leituras.`}
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Voltar</AlertDialogCancel>
            <AlertDialogAction
              onClick={() => {
                setConfirmandoEncerrar(false)
                encerrar.mutate()
              }}
            >
              Encerrar
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      <Dialog open={cancelando} onOpenChange={setCancelando}>
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle>Cancelar inventário</DialogTitle>
            <DialogDescription>
              O inventário deixa de valer e a unidade pode abrir outro. Informe o motivo.
            </DialogDescription>
          </DialogHeader>
          <label className="flex flex-col gap-1 text-caption text-muted-foreground">
            Motivo do cancelamento
            <textarea
              rows={3}
              maxLength={255}
              className="flex w-full rounded border border-input bg-surface-alt px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring"
              value={motivo}
              onChange={(e) => setMotivo(e.target.value)}
            />
          </label>
          <DialogFooter>
            <Button variant="outline" onClick={() => setCancelando(false)}>
              Voltar
            </Button>
            <Button
              variant="destructive"
              loading={cancelar.isPending}
              disabled={motivo.trim().length < 3}
              onClick={() => cancelar.mutate()}
            >
              Cancelar inventário
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  )
}

