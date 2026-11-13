import { useEffect, useState } from 'react'
import { useMutation, useQuery, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import {
  CheckCircle2,
  FileDown,
  Search,
  ClipboardCheck,
  PackageCheck,
} from 'lucide-react'

import { downloadAuthenticated } from '@/api/client'
import { getColaborador, getColaboradorPorCpf, type Colaborador } from '@/api/colaboradores'
import { getItem, listItemsPaginated, type Item } from '@/api/items'
import { confirmReturn, initiateReturn } from '@/api/loans'
import { usePageSize } from '@/hooks/usePageSize'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
import { formatDate } from '@/lib/utils'
import { ColaboradorSelect } from '@/components/colaboradores/ColaboradorSelect'
import { ColaboradorEmailField } from '@/components/termos/ColaboradorEmailField'
import { garantirEmailDoColaborador } from '@/components/termos/colaboradorEmail'
import { FileUpload } from '@/components/ui/FileUpload'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Checkbox } from '@/components/ui/checkbox'
import { Pagination } from '@/components/ui/Pagination'
import { StatusPill } from '@/components/ui/StatusPill'
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

const MSG_CPF_NAO_CONFERE = 'O CPF do colaborador escolhido não confere com o do responsável pelo equipamento.'

type ResultadoDevolucao = Awaited<ReturnType<typeof initiateReturn>>

interface IniciarDevolucaoDialogProps {
  item: Item
  onClose: () => void
  onConcluido: (resultado: ResultadoDevolucao, colaborador: Colaborador) => void
}

/**
 * Escolha de quem devolve antes de gerar o termo: o colaborador do cadastro vinculado ao
 * equipamento (fixo) ou, em equipamento antigo sem vínculo, qualquer colaborador — inclusive
 * desligado, que também devolve. O e-mail é obrigatório porque o RevSign envia o termo para ele.
 */
function IniciarDevolucaoDialog({ item, onClose, onConcluido }: IniciarDevolucaoDialogProps) {
  const vinculadoId = item.colaborador_id ?? null
  const cpfDoItem = (item.cpf ?? '').replace(/\D/g, '')
  const [escolhido, setEscolhido] = useState<Colaborador | null>(null)
  const [email, setEmail] = useState('')

  const { data: vinculado } = useQuery({
    queryKey: ['colaboradores', 'detalhe', vinculadoId],
    queryFn: () => getColaborador(vinculadoId as number),
    enabled: vinculadoId !== null,
  })
  const colaborador = vinculadoId !== null ? (vinculado ?? null) : escolhido

  useEffect(() => {
    if (vinculado) setEmail(vinculado.email ?? '')
  }, [vinculado])

  // Equipamento antigo (sem vínculo de cadastro): o termo sai com o nome e o CPF registrados no
  // equipamento, então já deixa escolhido o colaborador desse CPF (se houver no cadastro).
  const { data: doCpf } = useQuery({
    queryKey: ['colaboradores', 'por-cpf', cpfDoItem],
    queryFn: () => getColaboradorPorCpf(cpfDoItem),
    enabled: vinculadoId === null && cpfDoItem !== '',
    retry: false,
  })
  useEffect(() => {
    if (doCpf) {
      setEscolhido((atual) => atual ?? doCpf)
      setEmail((atual) => atual || (doCpf.email ?? ''))
    }
  }, [doCpf])

  function escolher(c: Colaborador | null) {
    setEscolhido(c)
    setEmail(c?.email ?? '')
  }

  const cpfNaoConfere =
    vinculadoId === null &&
    cpfDoItem !== '' &&
    colaborador !== null &&
    colaborador.cpf.replace(/\D/g, '') !== cpfDoItem

  const mutation = useMutation({
    mutationFn: async () => {
      if (!colaborador) throw new Error('Selecione o colaborador.')
      if (cpfNaoConfere) throw new Error(MSG_CPF_NAO_CONFERE)
      const atualizado = await garantirEmailDoColaborador(colaborador, email)
      const resultado = await initiateReturn(item.id, { colaborador_id: atualizado.id })
      return { resultado, atualizado }
    },
    onSuccess: ({ resultado, atualizado }) => onConcluido(resultado, atualizado),
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao gerar termo de devolução.'), 'error'),
  })

  return (
    <Dialog open onOpenChange={(open) => !open && onClose()}>
      <DialogContent className="max-w-lg border border-border bg-surface p-6">
        <DialogHeader>
          <DialogTitle className="text-body-lg font-semibold text-foreground">Iniciar devolução</DialogTitle>
          <DialogDescription className="text-caption text-muted-foreground">
            {item.tipo} {item.brand} {item.model} — responsável registrado: {item.assigned_to}
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-4">
          {vinculadoId === null && (
            <p className="text-caption text-muted-foreground">
              Equipamento sem colaborador do cadastro: escolha quem está devolvendo (responsável registrado:{' '}
              {item.assigned_to}).
            </p>
          )}
          <div className="space-y-1.5">
            <span id="devolucao-colaborador-rotulo" className="text-caption font-medium text-foreground">
              Colaborador *
            </span>
            <div role="group" aria-labelledby="devolucao-colaborador-rotulo">
              <ColaboradorSelect
                value={colaborador?.id ?? vinculadoId}
                onChange={escolher}
                somenteAtivos={false}
                disabled={vinculadoId !== null || mutation.isPending}
              />
            </div>
            {cpfNaoConfere && (
              <p role="alert" className="text-xs text-destructive">
                {MSG_CPF_NAO_CONFERE}
              </p>
            )}
          </div>
          <ColaboradorEmailField
            colaborador={colaborador}
            value={email}
            onChange={setEmail}
            disabled={mutation.isPending}
          />
        </div>

        <DialogFooter className="gap-2">
          <Button variant="outline" size="sm" onClick={onClose}>
            Cancelar
          </Button>
          <Button
            size="sm"
            onClick={() => mutation.mutate()}
            loading={mutation.isPending}
            disabled={cpfNaoConfere}
          >
            Gerar termo e iniciar
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  )
}

export default function ReturnPage() {
  const queryClient = useQueryClient()
  const [pendingReturnId, setPendingReturnId] = useState<number | null>(null)
  const [signedPdf, setSignedPdf] = useState<File | null>(null)
  // Equipamento cuja devolução está sendo iniciada (diálogo "Iniciar devolução").
  const [iniciando, setIniciando] = useState<Item | null>(null)

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
  const [emCampoPageSize, setStoredEmCampoPageSize] = usePageSize('devolucao-em-campo')
  const [indisponivelPage, setIndisponivelPage] = useState(1)
  const setEmCampoPageSize = (n: number) => {
    setStoredEmCampoPageSize(n)
    setIndisponivelPage(1)
  }
  const [searchActive, setSearchActive] = useState('')
  const [searchActiveAplicado, setSearchActiveAplicado] = useState('')
  useEffect(() => {
    const timer = setTimeout(() => {
      setSearchActiveAplicado(searchActive.trim())
      setIndisponivelPage(1)
    }, 400)
    return () => clearTimeout(timer)
  }, [searchActive])

  const { data: activeData, isFetching: activeFetching } = useQuery({
    queryKey: ['items', 'devolucao-ativos', indisponivelPage, emCampoPageSize, searchActiveAplicado],
    queryFn: () =>
      listItemsPaginated({
        status: 'Indisponível',
        search: searchActiveAplicado || undefined,
        limit: emCampoPageSize,
        offset: (indisponivelPage - 1) * emCampoPageSize,
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
  const indisponivelLastPage = Math.max(1, Math.ceil(indisponivelTotal / emCampoPageSize))

  // Contador "Em Campo" do cabeçalho: total real no servidor, sem a busca
  // (que alteraria o número enquanto o operador digita).
  const { data: emCampoCountData } = useQuery({
    queryKey: ['items', 'contagem', 'Indisponível'],
    queryFn: () => listItemsPaginated({ status: 'Indisponível', limit: 1 }),
  })
  const emCampoTotal = emCampoCountData?.total ?? 0

  // Devoluções pendentes de confirmação (termo já gerado, aguardando upload)
  const [pendentesPageSize, setStoredPendentesPageSize] = usePageSize('devolucao-pendentes')
  const [pendentesPage, setPendentesPage] = useState(1)
  const setPendentesPageSize = (n: number) => {
    setStoredPendentesPageSize(n)
    setPendentesPage(1)
  }
  const { data: pendentesData, isFetching: pendentesFetching } = useQuery({
    queryKey: ['items', 'devolucao-pendentes', pendentesPage, pendentesPageSize],
    queryFn: () =>
      listItemsPaginated({
        status: 'Pendente Devolução',
        limit: pendentesPageSize,
        offset: (pendentesPage - 1) * pendentesPageSize,
      }),
    placeholderData: keepPreviousData,
  })
  const pendenteDevolucao = pendentesData?.items ?? []
  const pendenteDevolucaoTotal = pendentesData?.total ?? 0
  const pendentesLastPage = Math.max(1, Math.ceil(pendenteDevolucaoTotal / pendentesPageSize))

  // Se a última página esvaziar (item mudou de status), volta para a última
  // página que ainda existe em vez de mostrar uma tabela vazia.
  useEffect(() => {
    if (pendentesData && pendentesPage > pendentesLastPage) setPendentesPage(pendentesLastPage)
  }, [pendentesData, pendentesPage, pendentesLastPage])
  useEffect(() => {
    if (activeData && indisponivelPage > indisponivelLastPage) setIndisponivelPage(indisponivelLastPage)
  }, [activeData, indisponivelPage, indisponivelLastPage])

  function devolucaoIniciada(resultado: ResultadoDevolucao, colaborador: Colaborador) {
    const item = iniciando
    void queryClient.invalidateQueries({ queryKey: ['items'] })
    setIniciando(null)
    if (resultado.termo?.status === 'enviado') {
      // O RevSign já mandou o termo: não há documento para baixar nem conferência a abrir agora.
      toast(
        `Termo de devolução enviado para assinatura em ${resultado.termo.destinatario_email ?? colaborador.email}.`
      )
      return
    }
    // Assinatura manual: baixa o termo (download autenticado, nunca window.open) e abre a conferência.
    downloadAuthenticated(resultado.download_url, resultado.filename).catch((err: unknown) => {
      toast(getErrorMessage(err, 'Não foi possível baixar o termo de devolução.'), 'error')
    })
    if (item) {
      setPendingReturnId(item.id)
      setSignedPdf(null)
      resetChecklist()
    }
    toast('Termo de devolução gerado! Realize a inspeção física e anexe o documento.')
  }

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

      {/* Escolha do colaborador que devolve e geração do termo */}
      {iniciando && (
        <IniciarDevolucaoDialog
          item={iniciando}
          onClose={() => setIniciando(null)}
          onConcluido={devolucaoIniciada}
        />
      )}

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
                <span className="text-foreground">{upper(activeSelectedItem.revenda) || '—'}</span>
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
              disabled={!signedPdf || !checklistCompleto}
              loading={confirmMutation.isPending}
              onClick={() => {
                if (pendingReturnId && signedPdf) {
                  confirmMutation.mutate({ itemId: pendingReturnId, pdf: signedPdf })
                }
              }}
            >
              <CheckCircle2 className="mr-1.5 size-3.5" />
              Finalizar Devolução
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
            <StatusPill status="Pendente Devolução" />
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
                    <td className="py-2 px-3 text-muted-foreground">{upper(item.revenda) || '—'}</td>
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

          <Pagination
            page={Math.min(pendentesPage, pendentesLastPage)}
            pageSize={pendentesPageSize}
            total={pendenteDevolucaoTotal}
            onPageChange={setPendentesPage}
            onPageSizeChange={setPendentesPageSize}
            isFetching={pendentesFetching}
          />
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
                    <td className="py-2 px-3 text-muted-foreground">{upper(item.revenda) || '—'}</td>
                    <td className="py-2 px-3 font-mono text-muted-foreground">
                      {formatDate(item.date_issued)}
                    </td>
                    <td className="py-2 px-3 text-right">
                      <Button
                        size="sm"
                        variant="outline"
                        loading={iniciando?.id === item.id}
                        onClick={() => setIniciando(item)}
                        className="text-xs h-7 px-2.5"
                      >
                        <FileDown className="mr-1 size-3" />
                        Iniciar Devolução
                      </Button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {indisponivelTotal > 0 && (
          <Pagination
            page={Math.min(indisponivelPage, indisponivelLastPage)}
            pageSize={emCampoPageSize}
            total={indisponivelTotal}
            onPageChange={setIndisponivelPage}
            onPageSizeChange={setEmCampoPageSize}
            isFetching={activeFetching}
          />
        )}
      </div>
    </div>
  )
}
