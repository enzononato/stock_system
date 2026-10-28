import { useEffect, useState } from 'react'
import { useQuery, keepPreviousData } from '@tanstack/react-query'
import { listItemsPaginated } from '@/api/items'
import { downloadSignedTerm } from '@/api/loans'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { toast } from '@/components/ui/toast'
import { Dialog, DialogContent, DialogDescription, DialogTitle } from '@/components/ui/dialog'
import { getErrorMessage } from '@/lib/api-error'
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { cn, formatDate } from '@/lib/utils'
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { FileDown, CheckCircle, Clock, ShieldCheck, Printer, Upload } from 'lucide-react'

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

// Abre o PDF assinado numa nova aba. Compartilhada pela coluna "Termo" da
// tabela de assinados e pelo botão "Baixar PDF Assinado" do visualizador.
async function viewSignedTerm(itemId: number): Promise<void> {
  try {
    const blob = await downloadSignedTerm(itemId)
    const url = URL.createObjectURL(blob)
    window.open(url, '_blank')
    setTimeout(() => URL.revokeObjectURL(url), 10000)
  } catch (err) {
    toast(getErrorMessage(err, 'Termo assinado não encontrado.'), 'error')
  }
}

type FilterTab = 'pendentes' | 'assinados'

export default function TermsPage() {
  const [confirmingId, setConfirmingId] = useState<number | null>(null)
  const [filterTab, setFilterTab] = useState<FilterTab>('pendentes')
  // Linha aberta no visualizador de documento (modal). Guarda o próprio item
  // clicado — nunca cai em "primeiro da lista" quando nada foi selecionado.
  const [viewingItem, setViewingItem] = useState<Item | null>(null)

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
        <Button size="sm" variant="outline" onClick={() => viewSignedTerm(row.original.id)}>
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
        actions={
          <>
            <div className="flex items-center gap-2 px-3 py-1.5 rounded border border-border bg-surface text-caption">
              <Clock className="size-3.5 text-muted-foreground" />
              <span className="text-muted-foreground">Pendentes de Assinatura:</span>
              <span className="num font-semibold text-foreground">{pendentesTotal}</span>
            </div>
            <div className="flex items-center gap-2 px-3 py-1.5 rounded border border-border bg-surface text-caption">
              <ShieldCheck className="size-3.5 text-muted-foreground" />
              <span className="text-muted-foreground">Vigentes / Assinados:</span>
              <span className="num font-semibold text-foreground">{ativosTotal}</span>
            </div>
          </>
        }
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
            onRowClick={setViewingItem}
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
            onRowClick={setViewingItem}
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

      {/* Visualizador de documento: termo formal do item clicado na tabela.
          Abre em modal, sem substituir o DataTable. */}
      <Dialog open={viewingItem !== null} onOpenChange={(open) => { if (!open) setViewingItem(null) }}>
        <DialogContent className="max-w-3xl max-h-[90vh] overflow-y-auto">
          {viewingItem && (
            <>
              <div className="border-b border-border pb-5 pr-6 space-y-3">
                <div>
                  <span className="text-caption font-semibold uppercase tracking-widest text-muted-foreground">
                    Documento Oficial de Custódia de Ativo
                  </span>
                  <DialogTitle className="text-heading-sm font-bold mt-1 leading-snug">
                    TERMO DE RESPONSABILIDADE E GUARDA DE EQUIPAMENTO
                  </DialogTitle>
                  <DialogDescription className="num text-caption text-muted-foreground mt-0.5">
                    REF: TR-{String(viewingItem.id).padStart(6, '0')} • EMISSÃO: {formatDate(viewingItem.date_issued)}
                  </DialogDescription>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <Button size="sm" variant="outline" onClick={() => generateAndDownloadLoanTerm(viewingItem.id)}>
                    <Printer size={13} />Imprimir Minuta
                  </Button>
                  {viewingItem.status === 'Pendente' ? (
                    <Button
                      size="sm"
                      onClick={() => {
                        setConfirmingId(viewingItem.id)
                        setViewingItem(null)
                      }}
                    >
                      <Upload size={13} />Anexar Assinado
                    </Button>
                  ) : (
                    <Button size="sm" variant="outline" onClick={() => viewSignedTerm(viewingItem.id)}>
                      <FileDown size={13} />Baixar PDF Assinado
                    </Button>
                  )}
                </div>
              </div>

              {/* Cláusula 1: partes */}
              <div className="space-y-3">
                <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                  1. Identificação das Partes
                </span>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 rounded border border-border bg-surface-alt p-4 text-body-sm">
                  <div>
                    <span className="text-caption text-muted-foreground block">Custodiante / Colaborador</span>
                    <span className="font-semibold text-foreground">{viewingItem.assigned_to || '—'}</span>
                    <span className="num text-caption text-muted-foreground block mt-0.5">
                      CPF: {viewingItem.cpf || '—'}
                    </span>
                  </div>
                  <div>
                    <span className="text-caption text-muted-foreground block">Lotação e Unidade</span>
                    <span className="font-medium text-foreground">{viewingItem.revenda || '—'}</span>
                    <span className="text-caption text-muted-foreground block mt-0.5">
                      Setor: {viewingItem.setor || '—'}
                    </span>
                  </div>
                </div>
              </div>

              {/* Cláusula 2: especificação do equipamento */}
              <div className="space-y-3">
                <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                  2. Especificação do Equipamento Entregue
                </span>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 rounded border border-border bg-surface-alt p-4 text-caption">
                  <div>
                    <span className="text-muted-foreground block">ID Patrimonial</span>
                    <span className="num font-semibold text-foreground">#{viewingItem.id}</span>
                  </div>
                  <div>
                    <span className="text-muted-foreground block">Tipo / Categoria</span>
                    <span className="font-medium text-foreground">{viewingItem.tipo || '—'}</span>
                  </div>
                  <div>
                    <span className="text-muted-foreground block">Marca e Modelo</span>
                    <span className="font-medium text-foreground">
                      {[viewingItem.brand, viewingItem.model].filter(Boolean).join(' ') || '—'}
                    </span>
                  </div>
                  <div>
                    <span className="text-muted-foreground block">Serial / Identificador</span>
                    <span className="num text-foreground">{viewingItem.identificador || '—'}</span>
                  </div>
                </div>
              </div>

              {/* Cláusula 3: obrigações de guarda */}
              <div className="space-y-3">
                <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                  3. Condições e Obrigações de Guarda
                </span>
                <div className="space-y-2 text-caption text-muted-foreground leading-relaxed border-l-2 border-border pl-4">
                  <p>
                    <strong>3.1. Finalidade:</strong> O equipamento acima discriminado é de propriedade exclusiva
                    da Revalle e é cedido em comodato exclusivamente para a execução das atividades funcionais.
                  </p>
                  <p>
                    <strong>3.2. Conservação:</strong> O colaborador compromete-se a zelar pelo perfeito estado de
                    funcionamento e conservação do bem, comunicando imediatamente ao departamento de TI qualquer
                    falha, dano ou extravio.
                  </p>
                  <p>
                    <strong>3.3. Restituição:</strong> O colaborador compromete-se a restituir o equipamento nas
                    mesmas condições em que o recebeu em caso de rescisão contratual, transferência ou solicitação
                    expressa da empresa.
                  </p>
                </div>
              </div>

              {/* Linhas de assinatura */}
              <div className="pt-6 border-t border-border">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-8 pt-4">
                  <div className="border-t border-border-strong pt-2 text-center text-caption">
                    <span className="font-semibold text-foreground block">
                      {viewingItem.assigned_to || 'Colaborador'}
                    </span>
                    <span className="text-muted-foreground">Custodiante Responsável</span>
                  </div>
                  <div className="border-t border-border-strong pt-2 text-center text-caption">
                    <span className="font-semibold text-foreground block">Departamento de TI / Patrimônio</span>
                    <span className="text-muted-foreground">Revalle</span>
                  </div>
                </div>
              </div>
            </>
          )}
        </DialogContent>
      </Dialog>

      {/* Painel de confirmação com upload (T5: painel compartilhado com LoanPage) */}
      {confirmingId && (
        <ConfirmacaoTermo
          itemId={confirmingId}
          description="Faça o upload do termo de responsabilidade assinado (PDF)."
          variant="blue"
          uploadLabel="Arraste ou clique para enviar o PDF assinado"
          errorMessage="Erro ao confirmar."
          showGenerateButton
          onConfirmed={() => setConfirmingId(null)}
          onCancel={() => setConfirmingId(null)}
        />
      )}
    </div>
  )
}
