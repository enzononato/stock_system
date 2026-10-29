import { useEffect, useState } from 'react'
import { useQuery, keepPreviousData } from '@tanstack/react-query'
import { FileDown, Search, Upload, Clock, Printer, ShieldCheck } from 'lucide-react'

import { listItemsPaginated } from '@/api/items'
import type { Item } from '@/api/items'
import { downloadSignedTerm } from '@/api/loans'
import { getErrorMessage } from '@/lib/api-error'
import { formatDate } from '@/lib/utils'
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { toast } from '@/components/ui/toast'

// Layout mestre/detalhe da FT_STC (lista lateral de cartões + visualizador do
// termo formal), mas com a camada de dados da main: cada aba tem sua própria
// consulta PAGINADA NO SERVIDOR, com chave de cache específica. A FT_STC
// buscava até 500 itens numa chave genérica ['items'] e filtrava/re-fatiava
// tudo no cliente — o que trunca em silêncio quando o estoque passa disso.
// As duas consultas ficam sempre ativas porque os contadores do cabeçalho e
// das abas precisam do total de ambas; alternar de aba não dispara requisição.
const PAGE_SIZE = 10

type FilterTab = 'pendentes' | 'assinados'

// Abre o PDF assinado numa nova aba (botão "Baixar PDF Assinado").
async function viewSignedTerm(itemId: number): Promise<void> {
  try {
    const blob = await downloadSignedTerm(itemId)
    const url = URL.createObjectURL(blob)
    window.open(url, '_blank')
    window.setTimeout(() => URL.revokeObjectURL(url), 10000)
  } catch (err) {
    toast(getErrorMessage(err, 'Termo assinado não encontrado no repositório.'), 'error')
  }
}

export default function TermsPage() {
  const [filterTab, setFilterTab] = useState<FilterTab>('pendentes')
  const [confirmingId, setConfirmingId] = useState<number | null>(null)
  // Termo aberto no visualizador. Guarda o próprio item clicado — nunca cai
  // sozinho no "primeiro da lista": sem clique, o visualizador pede seleção.
  const [selectedItem, setSelectedItem] = useState<Item | null>(null)

  // Busca no servidor, compartilhada pelas duas abas (debounce de 400ms,
  // mesmo padrão de HistoryPage/StockPage). Trocar de aba não apaga o texto.
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

  const pendentesQuery = useQuery({
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
  const pendentes = pendentesQuery.data?.items ?? []
  const pendentesTotal = pendentesQuery.data?.total ?? 0

  const ativosQuery = useQuery({
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
  // Itens "Indisponível" órfãos (sem `assigned_to`) não são empréstimos ativos
  // de verdade — não aparecem na lista de assinados.
  const ativos = (ativosQuery.data?.items ?? []).filter((i) => Boolean(i.assigned_to))
  const ativosTotal = ativosQuery.data?.total ?? 0

  // Se o total encolher (ex.: o último termo da página foi confirmado e mudou
  // de aba), volta para a última página que ainda existe.
  useEffect(() => {
    const max = Math.max(0, Math.ceil(pendentesTotal / PAGE_SIZE) - 1)
    if (pendentesQuery.data && pendentesPageIndex > max) setPendentesPageIndex(max)
  }, [pendentesQuery.data, pendentesTotal, pendentesPageIndex])
  useEffect(() => {
    const max = Math.max(0, Math.ceil(ativosTotal / PAGE_SIZE) - 1)
    if (ativosQuery.data && ativosPageIndex > max) setAtivosPageIndex(max)
  }, [ativosQuery.data, ativosTotal, ativosPageIndex])

  const isPendentes = filterTab === 'pendentes'
  const currentList = isPendentes ? pendentes : ativos
  const currentTotal = isPendentes ? pendentesTotal : ativosTotal
  const currentPageIndex = isPendentes ? pendentesPageIndex : ativosPageIndex
  const setCurrentPageIndex = isPendentes ? setPendentesPageIndex : setAtivosPageIndex
  const currentLoading = isPendentes ? pendentesQuery.isLoading : ativosQuery.isLoading
  const totalPages = Math.max(1, Math.ceil(currentTotal / PAGE_SIZE))

  function handleTabChange(tab: FilterTab) {
    if (tab === filterTab) return
    setFilterTab(tab)
    // O termo aberto pertence à outra aba — não deixa o visualizador mostrando
    // um documento que não está na lista visível.
    setSelectedItem(null)
  }

  return (
    <div className="space-y-6 max-w-7xl mx-auto">
      {/* Cabeçalho */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Governança Jurídica
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Documento Corporativo</span>
          </div>
          <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
            Termos de Responsabilidade
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Visualizador formal de custódia patrimonial, controle de assinaturas e arquivamento probatório.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-3">
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
        </div>
      </div>

      {/* Painel de confirmação com upload do PDF assinado (compartilhado com LoanPage) */}
      {confirmingId && (
        <ConfirmacaoTermo
          itemId={confirmingId}
          description="Faça o upload do termo de responsabilidade assinado pelo colaborador (PDF)."
          variant="blue"
          uploadLabel="Arraste ou clique para enviar o PDF assinado"
          errorMessage="Erro ao confirmar termo."
          showGenerateButton
          onConfirmed={() => {
            setConfirmingId(null)
            // O item mudou de aba (Pendente → Indisponível); fecha o documento
            // para não exibir um status desatualizado.
            setSelectedItem(null)
          }}
          onCancel={() => setConfirmingId(null)}
        />
      )}

      {/* Grade: seletor lateral de 340px + visualizador de documento */}
      <div className="grid grid-cols-1 lg:grid-cols-[340px_1fr] gap-6 items-start">
        {/* Painel lateral: seletor de termos */}
        <div className="rounded border border-border bg-surface p-4 space-y-4">
          {/* Abas de filtro */}
          <div className="grid grid-cols-2 gap-1 p-1 rounded bg-surface-alt border border-border">
            <button
              type="button"
              onClick={() => handleTabChange('pendentes')}
              className={`py-1.5 text-caption font-semibold rounded transition-colors ${
                isPendentes ? 'bg-foreground text-background' : 'text-muted-foreground hover:text-foreground'
              }`}
            >
              Pendentes (<span className="num">{pendentesTotal}</span>)
            </button>
            <button
              type="button"
              onClick={() => handleTabChange('assinados')}
              className={`py-1.5 text-caption font-semibold rounded transition-colors ${
                !isPendentes ? 'bg-foreground text-background' : 'text-muted-foreground hover:text-foreground'
              }`}
            >
              Assinados (<span className="num">{ativosTotal}</span>)
            </button>
          </div>

          {/* Busca (no servidor: colaborador, marca, modelo, serial e host) */}
          <div className="relative">
            <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
            <Input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Buscar colaborador, marca, modelo, serial…"
              className="pl-8 h-8 text-body-sm"
            />
          </div>

          {/* Lista de termos da página atual */}
          <div className="space-y-1.5 max-h-[600px] overflow-y-auto pr-1">
            {currentLoading ? (
              <div className="py-8 text-center text-caption text-muted-foreground">Carregando termos…</div>
            ) : currentList.length === 0 ? (
              <div className="py-8 text-center text-caption text-muted-foreground">Nenhum termo encontrado.</div>
            ) : (
              currentList.map((item) => {
                const isSelected = selectedItem?.id === item.id
                return (
                  <button
                    type="button"
                    key={item.id}
                    onClick={() => setSelectedItem(item)}
                    className={`block w-full text-left p-3 rounded border cursor-pointer transition-all ${
                      isSelected
                        ? 'border-foreground bg-surface-alt font-medium shadow-sm'
                        : 'border-border hover:bg-surface-alt'
                    }`}
                  >
                    <div className="flex items-center justify-between text-caption">
                      <span className="num font-semibold text-foreground">#{item.id}</span>
                      <span className="num text-muted-foreground">{formatDate(item.date_issued)}</span>
                    </div>
                    <p className="text-body-sm text-foreground truncate mt-0.5">
                      {item.assigned_to || 'Sem responsável'}
                    </p>
                    <div className="flex items-center justify-between text-caption text-muted-foreground mt-1">
                      <span className="truncate">
                        {[item.tipo, item.brand].filter(Boolean).join(' ')}
                      </span>
                      <span>{item.revenda || '—'}</span>
                    </div>
                  </button>
                )
              })
            )}
          </div>

          {/* Paginação no servidor */}
          {currentTotal > PAGE_SIZE && (
            <div className="flex items-center justify-between border-t border-border pt-3">
              <Button
                size="sm"
                variant="outline"
                disabled={currentPageIndex === 0}
                onClick={() => setCurrentPageIndex((p) => Math.max(0, p - 1))}
                className="text-xs h-7 px-3"
              >
                ← Anterior
              </Button>
              <span className="num text-caption text-muted-foreground">
                {currentPageIndex + 1} de {totalPages}
              </span>
              <Button
                size="sm"
                variant="outline"
                disabled={currentPageIndex >= totalPages - 1}
                onClick={() => setCurrentPageIndex((p) => p + 1)}
                className="text-xs h-7 px-3"
              >
                Próxima →
              </Button>
            </div>
          )}
        </div>

        {/* Visualizador de documento corporativo */}
        {selectedItem ? (
          <div className="rounded border border-border bg-surface p-8 space-y-6 shadow-sm">
            {/* Cabeçalho do documento */}
            <div className="border-b border-border pb-6 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
              <div>
                <span className="text-caption font-semibold uppercase tracking-widest text-muted-foreground">
                  Documento Oficial de Custódia de Ativo
                </span>
                <h2 className="text-heading-sm font-bold tracking-tight text-foreground mt-1">
                  TERMO DE RESPONSABILIDADE E GUARDA DE EQUIPAMENTO
                </h2>
                <p className="num text-caption text-muted-foreground mt-0.5">
                  REF: TR-{String(selectedItem.id).padStart(6, '0')} • EMISSÃO: {formatDate(selectedItem.date_issued)}
                </p>
              </div>

              <div className="flex items-center gap-2">
                <Button size="sm" variant="outline" onClick={() => generateAndDownloadLoanTerm(selectedItem.id)}>
                  <Printer className="mr-1.5 size-3.5" />
                  Imprimir Minuta
                </Button>

                {selectedItem.status === 'Pendente' ? (
                  <Button size="sm" onClick={() => setConfirmingId(selectedItem.id)}>
                    <Upload className="mr-1.5 size-3.5" />
                    Anexar Assinado
                  </Button>
                ) : (
                  <Button size="sm" variant="outline" onClick={() => viewSignedTerm(selectedItem.id)}>
                    <FileDown className="mr-1.5 size-3.5" />
                    Baixar PDF Assinado
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
                  <span className="font-semibold text-foreground">{selectedItem.assigned_to || '—'}</span>
                  <span className="num text-caption text-muted-foreground block mt-0.5">
                    CPF: {selectedItem.cpf || '—'}
                  </span>
                </div>
                <div>
                  <span className="text-caption text-muted-foreground block">Lotação e Unidade</span>
                  <span className="font-medium text-foreground">{selectedItem.revenda || '—'}</span>
                  <span className="text-caption text-muted-foreground block mt-0.5">
                    Setor: {selectedItem.setor || '—'}
                  </span>
                </div>
              </div>
            </div>

            {/* Cláusula 2: discriminação do ativo */}
            <div className="space-y-3">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                2. Especificação do Equipamento Entregue
              </span>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 rounded border border-border bg-surface-alt p-4 text-caption">
                <div>
                  <span className="text-muted-foreground block">ID Patrimonial</span>
                  <span className="num font-semibold text-foreground">#{selectedItem.id}</span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Tipo / Categoria</span>
                  <span className="font-medium text-foreground">{selectedItem.tipo || '—'}</span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Marca e Modelo</span>
                  <span className="font-medium text-foreground">
                    {[selectedItem.brand, selectedItem.model].filter(Boolean).join(' ') || '—'}
                  </span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Serial / Identificador</span>
                  <span className="num text-foreground">{selectedItem.identificador || '—'}</span>
                </div>
              </div>
            </div>

            {/* Cláusula 3: obrigações */}
            <div className="space-y-3 pt-2">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                3. Condições e Obrigações de Guarda
              </span>
              <div className="space-y-2 text-caption text-muted-foreground leading-relaxed border-l-2 border-border pl-4">
                <p>
                  <strong>3.1. Finalidade:</strong> O equipamento acima discriminado é de propriedade exclusiva da
                  Revalle e é cedido em comodato exclusivamente para execução das atividades funcionais.
                </p>
                <p>
                  <strong>3.2. Conservação:</strong> O colaborador compromete-se a zelar pelo perfeito estado de
                  funcionamento e conservação do bem, comunicando imediatamente ao departamento de TI qualquer falha,
                  dano ou extravio.
                </p>
                <p>
                  <strong>3.3. Restituição:</strong> O colaborador compromete-se a restituir o equipamento nas mesmas
                  condições em que o recebeu no caso de rescisão contratual, transferência ou solicitação expressa da
                  empresa.
                </p>
              </div>
            </div>

            {/* Bloco de assinatura probatória */}
            <div className="pt-6 border-t border-border">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-8 pt-4">
                <div className="border-t border-border-strong pt-2 text-center text-caption">
                  <span className="font-semibold text-foreground block">
                    {selectedItem.assigned_to || 'Colaborador'}
                  </span>
                  <span className="text-muted-foreground">Custodiante Responsável</span>
                </div>
                <div className="border-t border-border-strong pt-2 text-center text-caption">
                  <span className="font-semibold text-foreground block">Departamento de TI / Patrimônio</span>
                  <span className="text-muted-foreground">Revalle Enterprise Core</span>
                </div>
              </div>
            </div>
          </div>
        ) : (
          <div className="rounded border border-border bg-surface p-12 text-center text-caption text-muted-foreground">
            Selecione um termo na coluna lateral para abrir o documento corporativo.
          </div>
        )}
      </div>
    </div>
  )
}
