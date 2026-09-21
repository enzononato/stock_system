import { useEffect, useState } from 'react'
import { useQuery, useMutation, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { getItem, listItemsPaginated } from '@/api/items'
import { downloadReturnTerm, confirmReturn } from '@/api/loans'
import { Button } from '@/components/ui/button'
import { FileUpload } from '@/components/ui/FileUpload'
import { DataTable } from '@/components/ui/DataTable'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { formatDate } from '@/lib/utils'
import { FileDown, CheckCircle } from 'lucide-react'

// Os dois blocos desta tela são tabelas de NAVEGAÇÃO (browsing/procura, não
// escolha via SearchableSelect) — paginação real no servidor, mesmo padrão do
// DataTable já usado em HistoryPage/StockPage (T3).
const PAGE_SIZE = 10

export default function ReturnPage() {
  const queryClient = useQueryClient()
  const [pendingReturnId, setPendingReturnId] = useState<number | null>(null)
  const [signedPdf, setSignedPdf] = useState<File | null>(null)

  // Empréstimos ativos (aguardando geração do termo de devolução)
  const [activePageIndex, setActivePageIndex] = useState(0)
  const [activeSearch, setActiveSearch] = useState('')
  const [activeSearchAplicado, setActiveSearchAplicado] = useState('')
  useEffect(() => {
    const timer = setTimeout(() => {
      setActiveSearchAplicado(activeSearch.trim())
      setActivePageIndex(0)
    }, 400)
    return () => clearTimeout(timer)
  }, [activeSearch])

  const { data: activeData } = useQuery({
    queryKey: ['items', 'devolucao-ativos', activePageIndex, activeSearchAplicado],
    queryFn: () =>
      listItemsPaginated({
        status: 'Indisponível',
        search: activeSearchAplicado || undefined,
        limit: PAGE_SIZE,
        offset: activePageIndex * PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })
  // T3 (defeito real): sem o `Boolean(assigned_to)`, itens marcados
  // "Indisponível" mas órfãos (sem colaborador vinculado) apareciam aqui com
  // "Usuário" vazio, e o botão "Gerar Termo" respondia 400 ao tentar iniciar
  // a devolução para um item sem responsável. O backend não filtra por
  // "assigned_to preenchido" (não há esse parâmetro em `GET /api/items`), daí
  // o filtro ficar no cliente, sobre a página já pequena (10 itens) que vem
  // do servidor — não é o mesmo padrão do bug antigo (que filtrava 500 linhas
  // inteiras no cliente).
  const indisponivel = (activeData?.items ?? []).filter((i) => Boolean(i.assigned_to))
  const indisponivelTotal = activeData?.total ?? 0

  // Devoluções pendentes de confirmação (termo já gerado, aguardando upload)
  const [pendentesPageIndex, setPendentesPageIndex] = useState(0)
  const [pendentesSearch, setPendentesSearch] = useState('')
  const [pendentesSearchAplicado, setPendentesSearchAplicado] = useState('')
  useEffect(() => {
    const timer = setTimeout(() => {
      setPendentesSearchAplicado(pendentesSearch.trim())
      setPendentesPageIndex(0)
    }, 400)
    return () => clearTimeout(timer)
  }, [pendentesSearch])

  const { data: pendentesData } = useQuery({
    queryKey: ['items', 'devolucao-pendentes', pendentesPageIndex, pendentesSearchAplicado],
    queryFn: () =>
      listItemsPaginated({
        status: 'Pendente Devolução',
        search: pendentesSearchAplicado || undefined,
        limit: PAGE_SIZE,
        offset: pendentesPageIndex * PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })
  const pendenteDevolucao = pendentesData?.items ?? []
  const pendenteDevolucaoTotal = pendentesData?.total ?? 0

  const initiateMutation = useMutation({
    // downloadReturnTerm encapsula initiate + download autenticado (T1): o
    // fluxo antigo fazia window.open(data.download_url, '_blank'), que nunca
    // envia o header Authorization — o termo de devolução sempre voltava 401
    // e nunca abria de fato.
    mutationFn: (itemId: number) => downloadReturnTerm(itemId),
    onSuccess: (_, itemId) => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setPendingReturnId(itemId)
      toast('Termo de devolução gerado! Faça a assinatura e confirme.')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao gerar termo.'), 'error')
    },
  })

  const confirmMutation = useMutation({
    mutationFn: ({ itemId, pdf }: { itemId: number; pdf: File }) => confirmReturn(itemId, pdf),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setPendingReturnId(null)
      setSignedPdf(null)
      toast('Devolução confirmada com sucesso!')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao confirmar devolução.'), 'error')
    },
  })

  // Busca dedicada do item em confirmação: `indisponivel`/`pendenteDevolucao`
  // agora são só a página atual (10 itens), então o item que disparou a
  // confirmação pode não estar mais nela (a invalidação de ['items'] muda o
  // status dele e ele "muda de tabela"). Antes, com os 500 itens inteiros em
  // memória, `items.find(...)` sempre achava — buscar por id direto no
  // servidor reproduz o mesmo resultado sem depender do array completo.
  const { data: pendingReturnItem } = useQuery({
    queryKey: ['items', 'detail', pendingReturnId],
    queryFn: () => getItem(pendingReturnId as number),
    enabled: pendingReturnId !== null,
  })

  const activeColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'cpf', header: 'CPF', cell: ({ getValue }) => <span className="num">{getValue() as string}</span> },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Empréstimo', cell: ({ getValue }) => <span className="num">{formatDate(getValue() as string)}</span> },
    {
      id: 'actions',
      header: '',
      cell: ({ row }) => (
        <Button size="sm" variant="outline" onClick={() => initiateMutation.mutate(row.original.id)}>
          <FileDown size={14} />Gerar Termo
        </Button>
      ),
    },
  ]

  const pendingColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'revenda', header: 'Revenda' },
    {
      id: 'actions',
      header: '',
      cell: ({ row }) => (
        <Button size="sm" variant="ghost" onClick={() => setPendingReturnId(row.original.id)}>
          Confirmar
        </Button>
      ),
    },
  ]

  return (
    <div className="space-y-8">
      <PageHeader
        eyebrow="Movimentação de Ativos"
        eyebrowDetail="Encerramento de Empréstimo"
        title="Devolver Equipamento"
        description="Gerencie as devoluções de equipamentos emprestados."
      />

      {/* Empréstimos ativos */}
      <div className="space-y-3">
        <PanelHeader
          title={<>Empréstimos Ativos (<span className="num">{indisponivelTotal}</span>)</>}
          description="Selecione um item para gerar o termo de devolução."
        />
        <DataTable
          data={indisponivel}
          columns={activeColumns}
          searchPlaceholder="Buscar por usuário, marca, modelo..."
          pagination={{
            total: indisponivelTotal,
            pageIndex: activePageIndex,
            pageSize: PAGE_SIZE,
            onPageChange: setActivePageIndex,
            search: activeSearch,
            onSearchChange: setActiveSearch,
          }}
        />
      </div>

      {/* Confirmação de devolução */}
      {pendingReturnId && (
        <div className="figure-ground-panel space-y-4">
          <PanelHeader title="Confirmação de Devolução" />
          <h3 className="text-heading-sm text-foreground">
            Confirmar Devolução — Item #<span className="num">{pendingReturnId}</span>
          </h3>
          {/* Mostra o equipamento e o colaborador antes de confirmar — sem isso o
              operador confirmava a devolução às cegas, vendo apenas o #id. */}
          {pendingReturnItem && (
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 border-t border-border pt-3">
              <div>
                <p className="text-caption text-muted-foreground">Tipo</p>
                <p className="text-body-sm text-foreground">{pendingReturnItem.tipo || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Marca</p>
                <p className="text-body-sm text-foreground">{pendingReturnItem.brand || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Modelo</p>
                <p className="text-body-sm text-foreground">{pendingReturnItem.model || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Colaborador</p>
                <p className="text-body-sm text-foreground">{pendingReturnItem.assigned_to || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Unidade</p>
                <p className="text-body-sm text-foreground">{pendingReturnItem.revenda || '-'}</p>
              </div>
            </div>
          )}
          <p className="text-body-sm text-muted-foreground">
            O termo de devolução foi gerado. Faça o upload do PDF assinado para confirmar.
          </p>
          <FileUpload onFile={setSignedPdf} label="Upload do Termo de Devolução Assinado (PDF)" />
          <div className="flex gap-3">
            <Button
              disabled={!signedPdf || confirmMutation.isPending}
              onClick={() => signedPdf && confirmMutation.mutate({ itemId: pendingReturnId, pdf: signedPdf })}
            >
              <CheckCircle size={14} />
              {confirmMutation.isPending ? 'Confirmando...' : 'Confirmar Devolução'}
            </Button>
            <Button variant="ghost" onClick={() => setPendingReturnId(null)}>Cancelar</Button>
          </div>
        </div>
      )}

      {/* Devoluções pendentes de confirmação */}
      {pendenteDevolucaoTotal > 0 && (
        <div className="space-y-3">
          <h3 className="text-body-lg font-semibold text-foreground">Pendente de Confirmação (<span className="num">{pendenteDevolucaoTotal}</span>)</h3>
          <DataTable
            data={pendenteDevolucao}
            columns={pendingColumns}
            searchPlaceholder="Buscar por usuário, marca, modelo..."
            pagination={{
              total: pendenteDevolucaoTotal,
              pageIndex: pendentesPageIndex,
              pageSize: PAGE_SIZE,
              onPageChange: setPendentesPageIndex,
              search: pendentesSearch,
              onSearchChange: setPendentesSearch,
            }}
          />
        </div>
      )}
    </div>
  )
}
