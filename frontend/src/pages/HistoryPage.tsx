import { useCallback, useEffect, useState } from 'react'
import { useMutation, useQuery, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import {
  Download,
  Paperclip,
  RotateCcw,
  Search,
  AlertTriangle,
  ChevronDown,
  ChevronsUpDown,
} from 'lucide-react'

import { listHistoryPaginated, reverseEntryWithPassword, type HistoryEntry } from '@/api/history'
import { downloadAuthenticated } from '@/api/client'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { Pagination } from '@/components/ui/Pagination'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { useAuth } from '@/contexts/AuthContext'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { usePageSize } from '@/hooks/usePageSize'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
import { cn, formatDateTime, exportToCsv } from '@/lib/utils'
import { toast } from '@/components/ui/toast'

const REVERSIBLE_OPS = [
  'Cadastro',
  'Empréstimo',
  'Confirmação Empréstimo',
  'Devolução',
  'Confirmação Devolução',
]

function operationSymbol(op?: string) {
  if (!op) return '•'
  if (op.includes('Empréstimo')) return '→'
  if (op.includes('Devolução')) return '←'
  if (op === 'Cadastro') return '+'
  if (op === 'Exclusão') return '×'
  if (op === 'Estorno') return '↺'
  return '•'
}

function operationColor(op?: string) {
  if (!op) return 'text-muted-foreground'
  if (op.includes('Empréstimo')) return 'text-foreground'
  if (op.includes('Devolução')) return 'text-foreground'
  if (op === 'Exclusão') return 'text-destructive'
  if (op === 'Estorno') return 'text-muted-foreground'
  return 'text-muted-foreground'
}

/**
 * Deriva um nome de arquivo legível para download a partir da chave de storage
 * ("categoria/arquivo.ext" — ver `app/core/storage.py` no backend). O trecho
 * após a primeira barra já é o nome original enviado pelo usuário, então não
 * há por que desmontar o padrão — só extrair esse trecho.
 */
function attachmentFilename(key: string): string {
  const idx = key.indexOf('/')
  return idx >= 0 ? key.slice(idx + 1) : key
}

export default function HistoryPage() {
  const { hasRole } = useAuth()
  const queryClient = useQueryClient()

  // A página vale só para o termo de busca em que foi escolhida: ao mudar o termo (já com atraso) a
  // página efetiva volta a 1 no mesmo render, sem pedido com (termo novo, página antiga).
  const [paginaEscolhida, setPaginaEscolhida] = useState({ n: 1, busca: '' })
  const [pageSize, setPageSize] = usePageSize('historico')
  // `search` é o que o usuário digita; `busca` é o que vai para a query, com
  // atraso, para não disparar uma requisição por tecla.
  const [search, setSearch] = useState('')
  const busca = useDebouncedValue(search.trim(), 400)
  const page = paginaEscolhida.busca === busca ? paginaEscolhida.n : 1
  const setPage = useCallback((n: number) => setPaginaEscolhida({ n, busca }), [busca])
  const [downloadingKey, setDownloadingKey] = useState<string | null>(null)

  // Confirmação de estorno: entrada sendo estornada (null = diálogo fechado),
  // senha digitada e mensagem de erro da última tentativa (ex.: 403).
  const [reversingEntry, setReversingEntry] = useState<HistoryEntry | null>(null)
  const [password, setPassword] = useState('')
  const [reverseError, setReverseError] = useState<string | null>(null)

  // Quais cartões da linha do tempo estão com os detalhes expandidos.
  const [expandedIds, setExpandedIds] = useState<Record<number, boolean>>({})

  const { data, isLoading, isFetching } = useQuery({
    queryKey: ['history', page, pageSize, busca],
    queryFn: () =>
      listHistoryPaginated({
        search: busca || undefined,
        limit: pageSize,
        offset: (page - 1) * pageSize,
      }),
    placeholderData: keepPreviousData,
  })

  const history = data?.items ?? []
  const total = data?.total ?? 0
  const lastPage = Math.max(1, Math.ceil(total / pageSize))

  const allExpanded = history.length > 0 && history.every((e) => expandedIds[e.id])

  function toggleExpand(id: number) {
    setExpandedIds((prev) => ({ ...prev, [id]: !prev[id] }))
  }

  function toggleAll() {
    if (allExpanded) {
      setExpandedIds({})
    } else {
      const next: Record<number, boolean> = {}
      history.forEach((e) => {
        next[e.id] = true
      })
      setExpandedIds(next)
    }
  }

  // Se um estorno (ou qualquer outra invalidação) reduzir o total de
  // registros a ponto da página atual deixar de existir (ex.: estornar o
  // único item da última página), volta para a última página válida em vez
  // de deixar a lista "presa" numa página vazia.
  useEffect(() => {
    if (total > 0 && page > lastPage) setPage(lastPage)
  }, [total, page, lastPage])

  function openReverseConfirm(entry: HistoryEntry) {
    setReversingEntry(entry)
    setPassword('')
    setReverseError(null)
  }

  function closeReverseConfirm() {
    setReversingEntry(null)
    setPassword('')
    setReverseError(null)
  }

  const reverseMutation = useMutation({
    mutationFn: ({ id, password }: { id: number; password: string }) =>
      reverseEntryWithPassword(id, password),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['history'] })
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      toast.success('Operação estornada com sucesso!')
      closeReverseConfirm()
    },
    onError: (err: unknown) => {
      const status = (err as { response?: { status?: number } })?.response?.status
      // 403 é a resposta específica de senha incorreta: mensagem fixa e
      // clara, sem fechar o painel, para o usuário poder tentar de novo.
      // Qualquer outro erro usa a mensagem traduzida genérica.
      const msg =
        status === 403
          ? 'Senha incorreta. Ação não autorizada.'
          : getErrorMessage(err, 'Erro ao estornar.')
      setReverseError(msg)
      toast.error(msg)
    },
  })

  function handleConfirmReverse() {
    if (!reversingEntry || !password) return
    reverseMutation.mutate({ id: reversingEntry.id, password })
  }

  async function handleDownloadAttachment(key: string) {
    setDownloadingKey(key)
    try {
      await downloadAuthenticated(`/api/documents/files/${key}`, attachmentFilename(key))
    } catch (err) {
      toast.error(getErrorMessage(err, 'Erro ao baixar anexo.'))
    } finally {
      setDownloadingKey(null)
    }
  }

  return (
    <div className="space-y-6 max-w-7xl mx-auto">
      {/* Cabeçalho */}
      <PageHeader
        eyebrow="Auditoria e Rastreabilidade"
        eyebrowDetail="Linha Temporal"
        title="Histórico de Operações"
        description="Registro cronológico e imutável de todas as ações realizadas no sistema patrimonial."
        actions={
        <Button
          variant="outline"
          size="sm"
          disabled={history.length === 0}
          onClick={() =>
            exportToCsv('historico', history.map((h) => ({ ...h, revenda: upper(h.revenda) })) as unknown as Record<string, unknown>[], [
              { key: 'id', label: 'ID' },
              { key: 'operador', label: 'Operador' },
              { key: 'operation', label: 'Operação' },
              { key: 'tipo', label: 'Tipo' },
              { key: 'marca', label: 'Marca' },
              { key: 'modelo', label: 'Modelo' },
              { key: 'identificador', label: 'Identificador' },
              { key: 'usuario', label: 'Usuário' },
              { key: 'setor', label: 'Setor' },
              { key: 'revenda', label: 'Revenda' },
              { key: 'data_operacao', label: 'Data' },
            ])
          }
        >
          <Download />
          Exportar CSV
        </Button>
        }
      />

      {/* Grid: 280px filtros / 1fr timeline */}
      <div className="grid grid-cols-1 lg:grid-cols-[280px_1fr] gap-6 items-start">
        {/* Painel de Filtros */}
        <div className="rounded-md border border-border bg-surface p-5 space-y-5">
          <div className="border-b border-border pb-2.5">
            <span className="text-caption font-semibold uppercase tracking-wider text-brand-text">
              Filtros de Consulta
            </span>
          </div>

          <div className="space-y-2">
            <Label className="text-caption font-medium">Busca Geral</Label>
            <div className="relative">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
              <Input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                aria-label="Busca geral"
                placeholder="Operador, usuário, operação…"
                className="pl-8 h-8 text-body-sm"
              />
            </div>
          </div>

          {/* Sumário de Página */}
          <div className="space-y-2 pt-2 border-t border-border">
            <span className="text-caption font-semibold uppercase tracking-wider text-brand-text block">
              Sumário — Página Atual
            </span>
            <div className="grid grid-cols-2 gap-2">
              {[
                { label: 'Total de Registros', value: total, symbol: '#' },
                {
                  label: 'Empréstimos',
                  value: history.filter((h) => h.operation?.includes('Empréstimo')).length,
                  symbol: '→',
                },
                {
                  label: 'Devoluções',
                  value: history.filter((h) => h.operation?.includes('Devolução')).length,
                  symbol: '←',
                },
                {
                  label: 'Estornos',
                  value: history.filter((h) => h.operation === 'Estorno').length,
                  symbol: '↺',
                },
              ].map((stat) => (
                <div key={stat.label} className="rounded-md bg-brand-soft p-3">
                  <span className="block text-caption text-brand-soft-foreground">{stat.label}</span>
                  <span className="block font-mono text-heading font-bold text-brand-soft-foreground">
                    {stat.symbol} {stat.value}
                  </span>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Timeline Editorial */}
        <div className="min-w-0 rounded-md border border-border bg-surface p-5 space-y-1">
          <div className="border-b border-border pb-3 mb-4 flex flex-wrap items-center justify-between gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-brand-text">
              Linha Temporal — {total} evento{total !== 1 ? 's' : ''} auditados
            </span>
            {history.length > 0 && (
              <Button
                type="button"
                size="sm"
                variant="ghost"
                onClick={toggleAll}
                className="h-6 px-2 text-[11px] text-muted-foreground hover:text-foreground"
              >
                <ChevronsUpDown />
                {allExpanded ? 'Recolher todos' : 'Expandir todos'}
              </Button>
            )}
          </div>

          {isLoading ? (
            <div className="py-12 text-center text-caption text-muted-foreground">
              Carregando registros…
            </div>
          ) : history.length === 0 ? (
            <div className="py-12 text-center text-caption text-muted-foreground">
              Nenhum registro encontrado para os filtros aplicados.
            </div>
          ) : (
            <div className="relative">
              {/* Linha vertical da timeline */}
              <div className="absolute left-[13.5px] top-2 bottom-2 w-px bg-border" />

              <div className="space-y-0">
                {history.map((entry) => {
                  const symbol = operationSymbol(entry.operation)
                  const colorClass = operationColor(entry.operation)
                  const isReversible =
                    REVERSIBLE_OPS.includes(entry.operation ?? '') && hasRole('Gestor')
                  const hasAttachments =
                    Boolean(entry.operacao_anexo) || Boolean(entry.termo_assinado_anexo)
                  const isExpanded = Boolean(expandedIds[entry.id])

                  return (
                    <div key={entry.id} className="relative flex gap-3 pb-2.5 last:pb-0">
                      {/* Nó da Timeline — `shadow-sm` no lugar do `shadow-xs` da FT_STC,
                          que não existe nesta configuração do Tailwind. */}
                      <div className="relative z-10 flex-shrink-0 mt-0.5">
                        <div
                          className={cn(
                            'size-7 rounded border border-border bg-surface flex items-center justify-center font-mono text-xs font-bold shadow-sm',
                            colorClass
                          )}
                        >
                          {symbol}
                        </div>
                      </div>

                      {/* Card Expansível da Entrada */}
                      <div className="flex-1 rounded-md border border-border bg-surface transition-colors overflow-hidden min-w-0">
                        {/* Linha Resumo */}
                        <div className="w-full px-3 py-2 flex flex-wrap items-center justify-between gap-2.5">
                          <button
                            type="button"
                            onClick={() => toggleExpand(entry.id)}
                            className="flex-1 flex items-center gap-2 flex-wrap min-w-0 text-left cursor-pointer hover:opacity-80 transition-opacity"
                            aria-expanded={isExpanded}
                          >
                            <span className="font-semibold text-body-sm text-foreground">
                              {entry.operation ?? '—'}
                            </span>
                            {(entry.item_id || entry.peripheral_id) && (
                              <span className="font-mono text-caption text-muted-foreground">
                                {entry.item_id ? `Item #${entry.item_id}` : ''}
                                {entry.peripheral_id ? ` Periférico #${entry.peripheral_id}` : ''}
                              </span>
                            )}
                            {!isExpanded && (entry.usuario || entry.tipo || entry.operador) && (
                              <span className="text-caption text-muted-foreground truncate hidden sm:inline">
                                •{' '}
                                {[entry.usuario, entry.tipo, entry.operador]
                                  .filter(Boolean)
                                  .slice(0, 2)
                                  .join(' • ')}
                              </span>
                            )}
                          </button>

                          <div className="flex items-center gap-2 shrink-0">
                            <span className="font-mono text-caption text-muted-foreground whitespace-nowrap">
                              {formatDateTime(entry.data_operacao)}
                            </span>

                            {isReversible && (
                              <Button
                                type="button"
                                size="sm"
                                variant="outline"
                                className="h-6 px-2 text-[11px] text-muted-foreground hover:text-destructive"
                                onClick={() => openReverseConfirm(entry)}
                              >
                                <RotateCcw />
                                Estornar
                              </Button>
                            )}

                            <button
                              type="button"
                              onClick={() => toggleExpand(entry.id)}
                              className="p-0.5 text-muted-foreground hover:text-foreground cursor-pointer"
                              aria-label={isExpanded ? 'Recolher detalhes' : 'Expandir detalhes'}
                            >
                              <ChevronDown
                                className={cn(
                                  'size-3.5 transition-transform duration-200',
                                  isExpanded && 'rotate-180 text-foreground'
                                )}
                              />
                            </button>
                          </div>
                        </div>

                        {/* Detalhes Recolhíveis */}
                        {isExpanded && (
                          <div className="px-3 pb-3 pt-1 border-t border-border space-y-2 bg-surface">
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-caption pt-1">
                              <div className="p-2 rounded-md bg-brand-soft text-brand-soft-foreground">
                                <span className="block text-[11px] uppercase tracking-wider">
                                  Operador
                                </span>
                                <span className="font-medium">
                                  {entry.operador || '—'}
                                </span>
                              </div>
                              {entry.usuario && (
                                <div className="p-2 rounded-md bg-brand-soft text-brand-soft-foreground">
                                  <span className="block text-[11px] uppercase tracking-wider">
                                    Colaborador
                                  </span>
                                  <span className="font-medium">
                                    {entry.usuario}
                                  </span>
                                </div>
                              )}
                              {(entry.tipo || entry.marca) && (
                                <div className="p-2 rounded-md bg-brand-soft text-brand-soft-foreground">
                                  <span className="block text-[11px] uppercase tracking-wider">
                                    Equipamento
                                  </span>
                                  <span className="font-medium">
                                    {entry.tipo} {entry.marca} {entry.modelo}
                                  </span>
                                </div>
                              )}
                              {entry.revenda && (
                                <div className="p-2 rounded-md bg-brand-soft text-brand-soft-foreground">
                                  <span className="block text-[11px] uppercase tracking-wider">
                                    Revenda
                                  </span>
                                  <span>{upper(entry.revenda)}</span>
                                </div>
                              )}
                            </div>

                            {entry.details && (
                              <p className="text-caption text-muted-foreground border-t border-border pt-1.5">
                                {entry.details}
                              </p>
                            )}

                            {/* Ações da Entrada */}
                            {(hasAttachments || isReversible) && (
                              <div className="flex flex-wrap items-center gap-2 pt-1 border-t border-border">
                                {entry.operacao_anexo && (
                                  <Button
                                    type="button"
                                    size="sm"
                                    variant="outline"
                                    className="h-6 px-2 text-[11px] text-muted-foreground hover:text-foreground"
                                    loading={downloadingKey === entry.operacao_anexo}
                                    onClick={(e) => {
                                      e.stopPropagation()
                                      void handleDownloadAttachment(entry.operacao_anexo!)
                                    }}
                                  >
                                    {downloadingKey !== entry.operacao_anexo && <Paperclip />}
                                    Comprovante
                                  </Button>
                                )}
                                {entry.termo_assinado_anexo && (
                                  <Button
                                    type="button"
                                    size="sm"
                                    variant="outline"
                                    className="h-6 px-2 text-[11px] text-muted-foreground hover:text-foreground"
                                    loading={downloadingKey === entry.termo_assinado_anexo}
                                    onClick={(e) => {
                                      e.stopPropagation()
                                      void handleDownloadAttachment(entry.termo_assinado_anexo!)
                                    }}
                                  >
                                    {downloadingKey !== entry.termo_assinado_anexo && <Paperclip />}
                                    Termo
                                  </Button>
                                )}
                                {isReversible && (
                                  <Button
                                    type="button"
                                    size="sm"
                                    variant="outline"
                                    className="h-6 px-2 text-[11px] text-muted-foreground hover:text-foreground ml-auto"
                                    onClick={(e) => {
                                      e.stopPropagation()
                                      openReverseConfirm(entry)
                                    }}
                                  >
                                    <RotateCcw />
                                    Estornar
                                  </Button>
                                )}
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  )
                })}
              </div>
            </div>
          )}

          <div className="border-t border-border pt-4 mt-4">
            <Pagination
              page={page}
              pageSize={pageSize}
              total={total}
              onPageChange={setPage}
              onPageSizeChange={(n) => {
                setPageSize(n)
                setPage(1)
              }}
              isFetching={isFetching}
            />
          </div>
        </div>
      </div>

      {/* Dialog de Estorno — exige a senha do operador logado. Alto contraste
          proposital: essa ação reescreve o histórico registrado. */}
      <Dialog
        open={reversingEntry !== null}
        onOpenChange={(open) => !open && closeReverseConfirm()}
      >
        <DialogContent className="max-w-md border border-border bg-surface p-6">
          <DialogHeader>
            <div className="flex items-start gap-3 mb-2">
              <div className="p-2 rounded bg-foreground text-background shrink-0">
                <AlertTriangle className="size-4" />
              </div>
              <div>
                <DialogTitle asChild>
                  <h3 className="text-body-lg font-semibold text-foreground">
                    Confirmar Estorno — Operação #<span className="num">{reversingEntry?.id}</span>
                  </h3>
                </DialogTitle>
                <DialogDescription className="text-caption text-muted-foreground mt-0.5">
                  Esta ação reverterá os efeitos de "{reversingEntry?.operation}" no estoque de forma imediata.
                </DialogDescription>
              </div>
            </div>
          </DialogHeader>

          <div className="space-y-4 py-2">
            <div className="rounded-md bg-brand-soft p-3 text-caption text-brand-soft-foreground space-y-1">
              <div className="flex justify-between">
                <span>Operação:</span>
                <span className="font-semibold">{reversingEntry?.operation}</span>
              </div>
              <div className="flex justify-between">
                <span>Item envolvido:</span>
                <span className="font-mono">#{reversingEntry?.item_id || '—'}</span>
              </div>
              <div className="flex justify-between">
                <span>Operador original:</span>
                <span>{reversingEntry?.operador || '—'}</span>
              </div>
            </div>

            <div className="space-y-1.5">
              <Label htmlFor="senha-estorno" className="text-caption font-medium text-foreground">
                Confirme sua senha
              </Label>
              <Input
                id="senha-estorno"
                type="password"
                value={password}
                onChange={(e) => {
                  setPassword(e.target.value)
                  setReverseError(null)
                }}
                placeholder="Sua senha de acesso"
                className="h-9"
                autoFocus
              />
              {reverseError && (
                <p className="text-caption text-destructive">{reverseError}</p>
              )}
            </div>
          </div>

          <DialogFooter className="pt-3 border-t border-border flex justify-end gap-2">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={closeReverseConfirm}
            >
              Cancelar
            </Button>
            <Button
              type="button"
              size="sm"
              disabled={!password}
              loading={reverseMutation.isPending}
              onClick={handleConfirmReverse}
            >
              {!reverseMutation.isPending && <RotateCcw />}
              Confirmar
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  )
}
