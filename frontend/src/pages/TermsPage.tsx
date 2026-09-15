import { useEffect, useState } from 'react'
import { useQuery, keepPreviousData } from '@tanstack/react-query'
import { listItemsPaginated } from '@/api/items'
import { downloadSignedTerm } from '@/api/loans'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { toast } from '@/components/ui/toast'
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { cn, formatDate } from '@/lib/utils'
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { FileDown, CheckCircle } from 'lucide-react'

// Os dois blocos desta tela ("Pendentes" e "Assinados") são tabelas de
// NAVEGAÇÃO — paginação real no servidor, mesmo padrão do DataTable já usado
// em HistoryPage/StockPage (T3). Diferente das outras quatro telas, aqui as
// duas listas convivem sob um único controle segmentado Pendentes/Assinados
// (T3, Step 6) — a única "aba" de verdade do sistema: as duas consultas ficam
// sempre ativas (não só a do lado visível), então alternar de aba não dispara
// requisição nova, só troca qual tabela já carregada aparece. A referência
// (origin/redesign-frontend) resolve isso com uma única busca de até 500
// itens re-fatiada no cliente — mas isso reintroduziria exatamente o truncamento
// silencioso que esta task inteira existe para eliminar (Step 4), então aqui
// cada aba tem sua própria paginação de servidor, e só a busca (texto digitado)
// é compartilhada entre as duas, no espírito do "re-fatia a mesma busca".
const PAGE_SIZE = 10

type FilterTab = 'pendentes' | 'assinados'

export default function TermsPage() {
  const [confirmingId, setConfirmingId] = useState<number | null>(null)
  const [filterTab, setFilterTab] = useState<FilterTab>('pendentes')

  // Busca compartilhada pelas duas abas (debounce de 400ms, mesmo padrão de
  // HistoryPage/StockPage) — trocar de aba não reseta o texto digitado.
  const [search, setSearch] = useState('')
  const [searchAplicado, setSearchAplicado] = useState('')
  const [pendentesPageIndex, setPendentesPageIndex] = useState(0)
  const [ativosPageIndex, setAtivosPageIndex] = useState(0)
  useEffect(() => {
    const timer = setTimeout(() => {
      setSearchAplicado(search.trim())
      setPendentesPageIndex(0)
      setAtivosPageIndex(0)
    }, 400)
    return () => clearTimeout(timer)
  }, [search])

  const { data: pendentesData } = useQuery({
    queryKey: ['items', 'termos-pendentes', pendentesPageIndex, searchAplicado],
    queryFn: () =>
      listItemsPaginated({
        status: 'Pendente',
        search: searchAplicado || undefined,
        limit: PAGE_SIZE,
        offset: pendentesPageIndex * PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })
  const pendentes = pendentesData?.items ?? []
  const pendentesTotal = pendentesData?.total ?? 0

  const { data: ativosData } = useQuery({
    queryKey: ['items', 'termos-ativos', ativosPageIndex, searchAplicado],
    queryFn: () =>
      listItemsPaginated({
        status: 'Indisponível',
        search: searchAplicado || undefined,
        limit: PAGE_SIZE,
        offset: ativosPageIndex * PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })
  // Mesmo ajuste do Step 5 (ReturnPage): itens "Indisponível" órfãos (sem
  // `assigned_to`) não são empréstimos ativos de verdade — não deveriam
  // aparecer aqui com "Usuário" vazio. A referência já filtra assim.
  const ativos = (ativosData?.items ?? []).filter((i) => Boolean(i.assigned_to))
  const ativosTotal = ativosData?.total ?? 0

  const pendingColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'model', header: 'Modelo' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'cpf', header: 'CPF', cell: ({ getValue }) => <span className="num">{getValue() as string}</span> },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Data', cell: ({ getValue }) => <span className="num">{formatDate(getValue() as string)}</span> },
    {
      id: 'actions',
      header: 'Ações',
      cell: ({ row }) => (
        <div className="flex gap-2">
          <Button size="sm" variant="outline" onClick={() => generateAndDownloadLoanTerm(row.original.id)}>
            <FileDown size={13} />Gerar Termo
          </Button>
          <Button size="sm" onClick={() => setConfirmingId(row.original.id)}>
            <CheckCircle size={13} />Confirmar
          </Button>
        </div>
      ),
    },
  ]

  const activeColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'cpf', header: 'CPF', cell: ({ getValue }) => <span className="num">{getValue() as string}</span> },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Data Empréstimo', cell: ({ getValue }) => <span className="num">{formatDate(getValue() as string)}</span> },
    {
      id: 'actions',
      header: 'Termo',
      cell: ({ row }) => (
        <Button size="sm" variant="outline" onClick={async () => {
          try {
            const blob = await downloadSignedTerm(row.original.id)
            const url = URL.createObjectURL(blob)
            window.open(url, '_blank')
            setTimeout(() => URL.revokeObjectURL(url), 10000)
          } catch {
            toast('Termo assinado não encontrado.', 'error')
          }
        }}>
          <FileDown size={13} />Ver Termo
        </Button>
      ),
    },
  ]

  return (
    <div className="space-y-8">
      <PageHeader
        eyebrow="Conformidade Documental"
        eyebrowDetail="Assinaturas Pendentes"
        title="Termos de Responsabilidade"
        description="Gerencie os termos de empréstimo pendentes e confirmados."
      />

      <div className="space-y-3">
        <PanelHeader
          title="Termos"
          description="Gere o termo, imprima, colete a assinatura e confirme o empréstimo com o PDF assinado."
        />

        {/* Controle segmentado Pendentes/Assinados (T3, Step 6) — única "aba"
            de verdade do sistema. Dois botões simples, sem biblioteca de tabs;
            as duas consultas acima já ficam ativas o tempo todo, então clicar
            aqui só troca qual tabela aparece, sem nova requisição. */}
        <div className="inline-flex rounded-md border border-border bg-surface-alt p-1 gap-1">
          <button
            type="button"
            onClick={() => setFilterTab('pendentes')}
            className={cn(
              'px-3 py-1.5 rounded-sm text-body-sm font-medium transition-colors duration-micro',
              filterTab === 'pendentes'
                ? 'bg-foreground text-background'
                : 'text-muted-foreground hover:text-foreground'
            )}
          >
            Pendentes (<span className="num">{pendentesTotal}</span>)
          </button>
          <button
            type="button"
            onClick={() => setFilterTab('assinados')}
            className={cn(
              'px-3 py-1.5 rounded-sm text-body-sm font-medium transition-colors duration-micro',
              filterTab === 'assinados'
                ? 'bg-foreground text-background'
                : 'text-muted-foreground hover:text-foreground'
            )}
          >
            Assinados (<span className="num">{ativosTotal}</span>)
          </button>
        </div>

        {filterTab === 'pendentes' ? (
          <DataTable
            data={pendentes}
            columns={pendingColumns}
            searchPlaceholder="Buscar por usuário, marca, modelo, CPF..."
            pagination={{
              total: pendentesTotal,
              pageIndex: pendentesPageIndex,
              pageSize: PAGE_SIZE,
              onPageChange: setPendentesPageIndex,
              search,
              onSearchChange: setSearch,
            }}
          />
        ) : (
          <DataTable
            data={ativos}
            columns={activeColumns}
            searchPlaceholder="Buscar por usuário, marca, modelo, CPF..."
            pagination={{
              total: ativosTotal,
              pageIndex: ativosPageIndex,
              pageSize: PAGE_SIZE,
              onPageChange: setAtivosPageIndex,
              search,
              onSearchChange: setSearch,
            }}
          />
        )}
      </div>

      {/* Painel de confirmação com upload (T5: painel compartilhado com LoanPage) */}
      {confirmingId && (
        <ConfirmacaoTermo
          itemId={confirmingId}
          description="Faça o upload do termo de responsabilidade assinado (PDF)."
          variant="blue"
          uploadLabel="Arraste ou clique para enviar o PDF assinado"
          errorMessage="Erro ao confirmar."
          onConfirmed={() => setConfirmingId(null)}
          onCancel={() => setConfirmingId(null)}
        />
      )}
    </div>
  )
}
