import { useEffect, useState } from 'react'
import { useMutation, useQuery, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import {
  CheckCircle2,
  FileDown,
  Search,
  ClipboardCheck,
  PackageCheck,
} from 'lucide-react'

import { getItem, listItemsPaginated } from '@/api/items'
import { confirmReturn, downloadReturnTerm } from '@/api/loans'
import { getErrorMessage } from '@/lib/api-error'
import { formatDate } from '@/lib/utils'
import { FileUpload } from '@/components/ui/FileUpload'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Badge } from '@/components/ui/badge'
import { Checkbox } from '@/components/ui/checkbox'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { toast } from '@/components/ui/toast'

// Layout visual idêntico ao da FT_STC; a lógica de dados é a da main: as
// duas tabelas são de NAVEGAÇÃO, com paginação e busca reais no servidor
// (limit/offset/search), em vez de buscar um lote fixo de itens e filtrar no
// cliente — que ficava incorreto em silêncio quando o total passava do teto
// de página do backend.
const PAGE_SIZE = 7

export default function ReturnPage() {
  const queryClient = useQueryClient()
  const [pendingReturnId, setPendingReturnId] = useState<number | null>(null)
  const [signedPdf, setSignedPdf] = useState<File | null>(null)

  // Checklist de inspeção física: o operador confirma o estado do
  // equipamento antes de reintegrá-lo ao estoque disponível — sem os três
  // itens marcados, o botão de finalizar continua desabilitado.
  const [checkChassi, setCheckChassi] = useState(false)
  const [checkTelaTecladoFonte, setCheckTelaTecladoFonte] = useState(false)
  const [checkPerifericos, setCheckPerifericos] = useState(false)
  const checklistCompleto = checkChassi && checkTelaTecladoFonte && checkPerifericos

  function resetChecklist() {
    setCheckChassi(false)
    setCheckTelaTecladoFonte(false)
    setCheckPerifericos(false)
  }

  function fecharConfirmacao() {
    setPendingReturnId(null)
    setSignedPdf(null)
    resetChecklist()
  }

  // Equipamentos em campo (aguardando geração do termo de devolução): busca
  // com debounce enviada ao servidor.
  const [indisponivelPage, setIndisponivelPage] = useState(0)
  const [searchActive, setSearchActive] = useState('')
  const [searchActiveAplicado, setSearchActiveAplicado] = useState('')
  useEffect(() => {
    const timer = setTimeout(() => {
      setSearchActiveAplicado(searchActive.trim())
      setIndisponivelPage(0)
    }, 400)
    return () => clearTimeout(timer)
  }, [searchActive])

  const { data: activeData } = useQuery({
    queryKey: ['items', 'devolucao-ativos', indisponivelPage, searchActiveAplicado],
    queryFn: () =>
      listItemsPaginated({
        status: 'Indisponível',
        search: searchActiveAplicado || undefined,
        limit: PAGE_SIZE,
        offset: indisponivelPage * PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })
  // Sem o `Boolean(assigned_to)`, itens "Indisponível" órfãos (sem
  // colaborador vinculado) apareciam com "Colaborador" vazio e o botão de
  // iniciar a devolução respondia 400. O backend não tem filtro por
  // "assigned_to preenchido", daí o filtro ficar no cliente — mas sobre a
  // página pequena que vem do servidor, não sobre um lote grande.
  const indisponivel = (activeData?.items ?? []).filter((i) => Boolean(i.assigned_to))
  const indisponivelTotal = activeData?.total ?? 0
  const indisponivelTotalPages = Math.ceil(indisponivelTotal / PAGE_SIZE)

  // Contador "Em Campo" do cabeçalho: total real no servidor, sem a busca
  // (que alteraria o número enquanto o operador digita).
  const { data: emCampoCountData } = useQuery({
    queryKey: ['items', 'contagem', 'Indisponível'],
    queryFn: () => listItemsPaginated({ status: 'Indisponível', limit: 1 }),
  })
  const emCampoTotal = emCampoCountData?.total ?? 0

  // Devoluções pendentes de confirmação (termo já gerado, aguardando upload)
  const [pendentesPage, setPendentesPage] = useState(0)
  const { data: pendentesData } = useQuery({
    queryKey: ['items', 'devolucao-pendentes', pendentesPage],
    queryFn: () =>
      listItemsPaginated({
        status: 'Pendente Devolução',
        limit: PAGE_SIZE,
        offset: pendentesPage * PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })
  const pendenteDevolucao = pendentesData?.items ?? []
  const pendenteDevolucaoTotal = pendentesData?.total ?? 0
  const pendentesTotalPages = Math.ceil(pendenteDevolucaoTotal / PAGE_SIZE)

  // Se a última página esvaziar (item mudou de status), volta para a última
  // página que ainda existe em vez de mostrar uma tabela vazia.
  useEffect(() => {
    if (pendentesTotalPages > 0 && pendentesPage > pendentesTotalPages - 1) {
      setPendentesPage(pendentesTotalPages - 1)
    }
  }, [pendentesPage, pendentesTotalPages])
  useEffect(() => {
    if (indisponivelTotalPages > 0 && indisponivelPage > indisponivelTotalPages - 1) {
      setIndisponivelPage(indisponivelTotalPages - 1)
    }
  }, [indisponivelPage, indisponivelTotalPages])

  const initiateMutation = useMutation({
    // downloadReturnTerm encapsula initiate + download autenticado: abrir a
    // URL em outra aba nunca envia o header Authorization (sempre 401).
    mutationFn: (itemId: number) => downloadReturnTerm(itemId),
    onSuccess: (_, itemId) => {
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      setPendingReturnId(itemId)
      setSignedPdf(null)
      resetChecklist()
      toast('Termo de devolução gerado! Realize a inspeção física e anexe o documento.')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao gerar termo de devolução.'), 'error')
    },
  })

  const confirmMutation = useMutation({
    mutationFn: ({ itemId, pdf }: { itemId: number; pdf: File }) => confirmReturn(itemId, pdf),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      fecharConfirmacao()
      toast('Devolução concluída com sucesso! Item retornado ao estoque disponível.')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao confirmar devolução.'), 'error')
    },
  })

  // Busca dedicada do item em confirmação: as tabelas guardam só a página
  // atual, e o item que disparou a confirmação pode já ter "mudado de
  // tabela" após a invalidação — buscar por id direto no servidor evita
  // depender de ele estar na página carregada.
  const { data: activeSelectedItem } = useQuery({
    queryKey: ['items', 'detail', pendingReturnId],
    queryFn: () => getItem(pendingReturnId as number),
    enabled: pendingReturnId !== null,
  })

  return (
    <div className="space-y-6 max-w-6xl mx-auto">
      {/* Cabeçalho */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Operações de Campo
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Conferência Física</span>
          </div>
          <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
            Devolução de Equipamentos
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Recepção patrimonial com checklist de integridade, laudo de avarias e arquivamento de termo.
          </p>
        </div>

        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded border border-border bg-surface text-caption">
            <PackageCheck className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Em Campo:</span>
            <span className="font-mono font-semibold text-foreground">{emCampoTotal}</span>
          </div>
          <div className="flex items-center gap-2 px-3 py-1.5 rounded border border-border bg-surface text-caption">
            <ClipboardCheck className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Aguardando Confirmação:</span>
            <span className="font-mono font-semibold text-foreground">
              {pendenteDevolucaoTotal}
            </span>
          </div>
        </div>
      </div>

      {/* Modal de conferência física e termo de devolução */}
      <Dialog
        open={pendingReturnId !== null}
        onOpenChange={(open) => {
          if (!open) fecharConfirmacao()
        }}
      >
        <DialogContent className="max-w-lg border border-border bg-surface p-6">
          <DialogHeader>
            <DialogTitle className="text-body-lg font-semibold text-foreground">
              Conferência Física & Termo de Devolução
            </DialogTitle>
            <DialogDescription className="text-caption text-muted-foreground">
              Verifique o estado físico do equipamento #{pendingReturnId} antes de reintegrar ao estoque disponível.
            </DialogDescription>
          </DialogHeader>

          {/* Mostra o equipamento e o colaborador antes de confirmar — sem
              isso o operador confirmava a devolução às cegas. */}
          {activeSelectedItem && (
            <div className="rounded border border-border bg-surface-alt p-3 text-caption space-y-1">
              <div className="flex justify-between">
                <span className="text-muted-foreground">Equipamento:</span>
                <span className="font-medium text-foreground">
                  {activeSelectedItem.tipo} {activeSelectedItem.brand} {activeSelectedItem.model}
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Responsável atual:</span>
                <span className="font-medium text-foreground">
                  {activeSelectedItem.assigned_to}
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Revenda:</span>
                <span className="text-foreground">{activeSelectedItem.revenda || '—'}</span>
              </div>
            </div>
          )}

          {/* Checklist físico */}
          <div className="space-y-3 pt-2 border-t border-border">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
              Checklist de Inspeção Física
            </span>

            <label className="flex items-center gap-2.5 cursor-pointer text-caption text-foreground select-none">
              <Checkbox checked={checkChassi} onCheckedChange={setCheckChassi} />
              <span>
                Integridade do chassi / gabinete (sem trincas, amassados ou danos estruturais)
              </span>
            </label>

            <label className="flex items-center gap-2.5 cursor-pointer text-caption text-foreground select-none">
              <Checkbox checked={checkTelaTecladoFonte} onCheckedChange={setCheckTelaTecladoFonte} />
              <span>Tela, teclado e fonte de alimentação testados e em pleno funcionamento</span>
            </label>

            <label className="flex items-center gap-2.5 cursor-pointer text-caption text-foreground select-none">
              <Checkbox checked={checkPerifericos} onCheckedChange={setCheckPerifericos} />
              <span>Periféricos e cabos complementares entregues e conferidos</span>
            </label>
          </div>

          {/* Upload do termo assinado */}
          <div className="space-y-2 pt-3 border-t border-border">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
              Termo de Devolução Assinado (PDF) *
            </span>
            <FileUpload
              onFile={setSignedPdf}
              label="Arraste ou clique para anexar o termo assinado (PDF)"
            />
          </div>

          <DialogFooter className="pt-3 border-t border-border flex justify-end gap-2">
            <Button variant="outline" size="sm" onClick={fecharConfirmacao}>
              Cancelar
            </Button>
            <Button
              size="sm"
              disabled={confirmMutation.isPending || !signedPdf || !checklistCompleto}
              onClick={() => {
                if (pendingReturnId && signedPdf) {
                  confirmMutation.mutate({ itemId: pendingReturnId, pdf: signedPdf })
                }
              }}
            >
              <CheckCircle2 className="mr-1.5 size-3.5" />
              {confirmMutation.isPending ? 'Concluindo…' : 'Finalizar Devolução'}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Seção 1: pendentes de confirmação de devolução */}
      {pendenteDevolucaoTotal > 0 && (
        <div className="rounded border border-border bg-surface p-5 space-y-3">
          <div className="flex items-center justify-between border-b border-border pb-2.5">
            <div>
              <h2 className="text-body font-semibold text-foreground">
                Aguardando Confirmação e Termo ({pendenteDevolucaoTotal})
              </h2>
              <p className="text-caption text-muted-foreground">
                Equipamentos cujo termo de devolução já foi gerado e aguardam upload da via assinada.
              </p>
            </div>
            <Badge variant="warning" showDot className="font-mono text-[11px] border-border">
              RECOLHIMENTO
            </Badge>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-body-sm">
              <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground">
                <tr>
                  <th className="py-2 px-3 font-semibold">ID</th>
                  <th className="py-2 px-3 font-semibold">Equipamento</th>
                  <th className="py-2 px-3 font-semibold">Colaborador</th>
                  <th className="py-2 px-3 font-semibold">CPF</th>
                  <th className="py-2 px-3 font-semibold">Revenda</th>
                  <th className="py-2 px-3 text-right font-semibold">Ação</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {pendenteDevolucao.map((item) => (
                  <tr key={item.id} className="hover:bg-surface-alt transition-colors">
                    <td className="py-2 px-3 font-mono font-medium">#{item.id}</td>
                    <td className="py-2 px-3">
                      {item.tipo} {item.brand} {item.model}
                    </td>
                    <td className="py-2 px-3 font-medium">{item.assigned_to}</td>
                    <td className="py-2 px-3 font-mono text-muted-foreground">
                      {item.cpf || '—'}
                    </td>
                    <td className="py-2 px-3 text-muted-foreground">{item.revenda || '—'}</td>
                    <td className="py-2 px-3 text-right">
                      <Button
                        size="sm"
                        onClick={() => {
                          setSignedPdf(null)
                          resetChecklist()
                          setPendingReturnId(item.id)
                        }}
                        className="text-xs h-7 px-2.5"
                      >
                        Confirmar Devolução
                      </Button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {pendentesTotalPages > 1 && (
            <div className="flex items-center justify-between border-t border-border pt-3">
              <span className="text-caption text-muted-foreground">
                Página {pendentesPage + 1} de {pendentesTotalPages} (
                {pendenteDevolucaoTotal} pendentes)
              </span>
              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pendentesPage === 0}
                  onClick={() => setPendentesPage((p) => Math.max(0, p - 1))}
                  className="text-xs h-7 px-2.5"
                >
                  ← Anterior
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pendentesPage >= pendentesTotalPages - 1}
                  onClick={() => setPendentesPage((p) => p + 1)}
                  className="text-xs h-7 px-2.5"
                >
                  Próxima →
                </Button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Seção 2: lista operacional de equipamentos em campo */}
      <div className="rounded border border-border bg-surface p-5 space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-border pb-3">
          <div>
            <h2 className="text-body font-semibold text-foreground">
              Equipamentos Alocados em Campo ({indisponivelTotal})
            </h2>
            <p className="text-caption text-muted-foreground">
              Selecione um ativo para iniciar o procedimento de devolução e emitir o documento formal.
            </p>
          </div>

          <div className="relative w-full sm:w-72">
            <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
            <Input
              value={searchActive}
              onChange={(e) => setSearchActive(e.target.value)}
              placeholder="Buscar por ID, modelo, CPF, colaborador…"
              className="pl-8 h-8 text-body-sm"
            />
          </div>
        </div>

        {indisponivel.length === 0 ? (
          <div className="py-8 text-center text-caption text-muted-foreground">
            Nenhum equipamento em campo encontrado para os critérios de busca.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-body-sm">
              <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground">
                <tr>
                  <th className="py-2 px-3 font-semibold">ID</th>
                  <th className="py-2 px-3 font-semibold">Tipo / Modelo</th>
                  <th className="py-2 px-3 font-semibold">Colaborador</th>
                  <th className="py-2 px-3 font-semibold">CPF</th>
                  <th className="py-2 px-3 font-semibold">Revenda</th>
                  <th className="py-2 px-3 font-semibold">Data Empréstimo</th>
                  <th className="py-2 px-3 text-right font-semibold">Ação</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {indisponivel.map((item) => (
                  <tr key={item.id} className="hover:bg-surface-alt transition-colors">
                    <td className="py-2 px-3 font-mono font-medium">#{item.id}</td>
                    <td className="py-2 px-3">
                      <span className="font-medium text-foreground">{item.tipo}</span>{' '}
                      <span className="text-muted-foreground">
                        {item.brand} {item.model}
                      </span>
                    </td>
                    <td className="py-2 px-3 font-medium">{item.assigned_to}</td>
                    <td className="py-2 px-3 font-mono text-muted-foreground">
                      {item.cpf || '—'}
                    </td>
                    <td className="py-2 px-3 text-muted-foreground">{item.revenda || '—'}</td>
                    <td className="py-2 px-3 font-mono text-muted-foreground">
                      {formatDate(item.date_issued)}
                    </td>
                    <td className="py-2 px-3 text-right">
                      <Button
                        size="sm"
                        variant="outline"
                        disabled={initiateMutation.isPending}
                        onClick={() => initiateMutation.mutate(item.id)}
                        className="text-xs h-7 px-2.5"
                      >
                        <FileDown className="mr-1 size-3" />
                        {initiateMutation.isPending ? 'Gerando…' : 'Iniciar Devolução'}
                      </Button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {indisponivelTotalPages > 1 && (
          <div className="flex items-center justify-between border-t border-border pt-3">
            <span className="text-caption text-muted-foreground">
              Página {indisponivelPage + 1} de {indisponivelTotalPages} (
              {indisponivelTotal} registros)
            </span>
            <div className="flex items-center gap-2">
              <Button
                size="sm"
                variant="outline"
                disabled={indisponivelPage === 0}
                onClick={() => setIndisponivelPage((p) => Math.max(0, p - 1))}
                className="text-xs h-7 px-2.5"
              >
                ← Anterior
              </Button>
              <Button
                size="sm"
                variant="outline"
                disabled={indisponivelPage >= indisponivelTotalPages - 1}
                onClick={() => setIndisponivelPage((p) => p + 1)}
                className="text-xs h-7 px-2.5"
              >
                Próxima →
              </Button>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}
