import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { toast } from '@/components/ui/toast'
import {
  FileDown,
  Search,
  Upload,
  Clock,
  Printer,
  ShieldCheck,
} from 'lucide-react'

import { listItemsPaginated } from '@/api/items'
import { downloadSignedTerm } from '@/api/loans'
import { getErrorMessage } from '@/lib/api-error'
import { formatDate } from '@/lib/utils'
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'

const FETCH_ALL_LIMIT = 500

export default function TermsPage() {
  const [selectedItemId, setSelectedItemId] = useState<number | null>(null)
  const [filterTab, setFilterTab] = useState<'pendentes' | 'assinados'>('pendentes')
  const [search, setSearch] = useState('')
  const [confirmingId, setConfirmingId] = useState<number | null>(null)

  const { data, refetch } = useQuery({
    queryKey: ['items'],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  })
  const items = data?.items ?? []
  const pendentes = items.filter((i) => i.status === 'Pendente')
  const ativos = items.filter((i) => i.status === 'Indisponível' && Boolean(i.assigned_to))

  const currentList = filterTab === 'pendentes' ? pendentes : ativos
  const filteredList = currentList.filter((i) => {
    if (!search.trim()) return true
    const q = search.toLowerCase()
    return (
      String(i.id).includes(q) ||
      (i.tipo ?? '').toLowerCase().includes(q) ||
      (i.brand ?? '').toLowerCase().includes(q) ||
      (i.assigned_to ?? '').toLowerCase().includes(q) ||
      (i.cpf ?? '').includes(q) ||
      (i.revenda ?? '').toLowerCase().includes(q)
    )
  })

  const selectedItem = items.find((i) => i.id === selectedItemId) || filteredList[0]

  async function handleViewSignedTerm(itemId: number) {
    try {
      const blob = await downloadSignedTerm(itemId)
      const url = URL.createObjectURL(blob)
      window.open(url, '_blank')
      window.setTimeout(() => URL.revokeObjectURL(url), 10000)
    } catch (err) {
      toast.error(getErrorMessage(err, 'Termo assinado não encontrado no repositório.'))
    }
  }

  return (
    <div className="space-y-6 max-w-7xl mx-auto">
      {/* Header */}
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

        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded border border-border bg-surface text-caption">
            <Clock className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Pendentes de Assinatura:</span>
            <span className="font-mono font-semibold text-foreground">{pendentes.length}</span>
          </div>
          <div className="flex items-center gap-2 px-3 py-1.5 rounded border border-border bg-surface text-caption">
            <ShieldCheck className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Vigentes / Assinados:</span>
            <span className="font-mono font-semibold text-foreground">{ativos.length}</span>
          </div>
        </div>
      </div>

      {confirmingId && (
        <ConfirmacaoTermo
          itemId={confirmingId}
          description="Faça o upload do termo de responsabilidade assinado pelo colaborador (PDF)."
          uploadLabel="Arraste ou clique para enviar o PDF assinado"
          errorMessage="Erro ao confirmar termo."
          showGenerateButton
          onConfirmed={() => {
            setConfirmingId(null)
            void refetch()
          }}
          onCancel={() => setConfirmingId(null)}
        />
      )}

      {/* Grid: 340px Seletor Lateral — Document Viewer Central */}
      <div className="grid grid-cols-1 lg:grid-cols-[340px_1fr] gap-6 items-start">
        {/* Painel Lateral: Seletor de Termos */}
        <div className="rounded border border-border bg-surface p-4 space-y-4">
          {/* Tabs de Filtro */}
          <div className="grid grid-cols-2 gap-1 p-1 rounded bg-surface-alt border border-border">
            <button
              type="button"
              onClick={() => setFilterTab('pendentes')}
              className={`py-1.5 text-caption font-semibold rounded transition-colors ${
                filterTab === 'pendentes'
                  ? 'bg-foreground text-background'
                  : 'text-muted-foreground hover:text-foreground'
              }`}
            >
              Pendentes ({pendentes.length})
            </button>
            <button
              type="button"
              onClick={() => setFilterTab('assinados')}
              className={`py-1.5 text-caption font-semibold rounded transition-colors ${
                filterTab === 'assinados'
                  ? 'bg-foreground text-background'
                  : 'text-muted-foreground hover:text-foreground'
              }`}
            >
              Assinados ({ativos.length})
            </button>
          </div>

          {/* Busca */}
          <div className="relative">
            <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
            <Input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Buscar colaborador, ID, CPF…"
              className="pl-8 h-8 text-body-sm"
            />
          </div>

          {/* Lista de Itens */}
          <div className="space-y-1.5 max-h-[600px] overflow-y-auto pr-1">
            {filteredList.length === 0 ? (
              <div className="py-8 text-center text-caption text-muted-foreground">
                Nenhum termo encontrado.
              </div>
            ) : (
              filteredList.map((item) => {
                const isSelected = selectedItem?.id === item.id
                return (
                  <div
                    key={item.id}
                    onClick={() => setSelectedItemId(item.id)}
                    className={`p-3 rounded border cursor-pointer transition-all ${
                      isSelected
                        ? 'border-foreground bg-surface-alt font-medium shadow-xs'
                        : 'border-border hover:bg-surface-alt'
                    }`}
                  >
                    <div className="flex items-center justify-between text-caption">
                      <span className="font-mono font-semibold text-foreground">#{item.id}</span>
                      <span className="text-muted-foreground font-mono">
                        {formatDate(item.date_issued)}
                      </span>
                    </div>
                    <p className="text-body-sm text-foreground truncate mt-0.5">
                      {item.assigned_to || 'Sem responsável'}
                    </p>
                    <div className="flex items-center justify-between text-caption text-muted-foreground mt-1">
                      <span className="truncate">
                        {item.tipo} {item.brand}
                      </span>
                      <span>{item.revenda || '—'}</span>
                    </div>
                  </div>
                )
              })
            )}
          </div>
        </div>

        {/* Visualizador de Documento Corporativo */}
        {selectedItem ? (
          <div className="rounded border border-border bg-surface p-8 space-y-6 shadow-sm">
            {/* Cabeçalho do Documento */}
            <div className="border-b border-border pb-6 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
              <div>
                <span className="text-caption font-semibold uppercase tracking-widest text-muted-foreground">
                  Documento Oficial de Custódia de Ativo
                </span>
                <h2 className="text-heading-sm font-bold tracking-tight text-foreground mt-1">
                  TERMO DE RESPONSABILIDADE E GUARDA DE EQUIPAMENTO
                </h2>
                <p className="font-mono text-caption text-muted-foreground mt-0.5">
                  REF: TR-{String(selectedItem.id).padStart(6, '0')} • EMISSÃO:{' '}
                  {formatDate(selectedItem.date_issued)}
                </p>
              </div>

              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => generateAndDownloadLoanTerm(selectedItem.id)}
                >
                  <Printer className="mr-1.5 size-3.5" />
                  Imprimir Minuta
                </Button>

                {selectedItem.status === 'Pendente' ? (
                  <Button
                    size="sm"
                    onClick={() => setConfirmingId(selectedItem.id)}
                  >
                    <Upload className="mr-1.5 size-3.5" />
                    Anexar Assinado
                  </Button>
                ) : (
                  <Button
                    size="sm"
                    variant="outline"
                    onClick={() => handleViewSignedTerm(selectedItem.id)}
                  >
                    <FileDown className="mr-1.5 size-3.5" />
                    Baixar PDF Assinado
                  </Button>
                )}
              </div>
            </div>

            {/* Cláusula 1: Partes */}
            <div className="space-y-3">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                1. Identificação das Partes
              </span>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 rounded border border-border bg-surface-alt p-4 text-body-sm">
                <div>
                  <span className="text-caption text-muted-foreground block">
                    Custodiante / Colaborador
                  </span>
                  <span className="font-semibold text-foreground">
                    {selectedItem.assigned_to || '—'}
                  </span>
                  <span className="text-caption font-mono text-muted-foreground block mt-0.5">
                    CPF: {selectedItem.cpf || '—'}
                  </span>
                </div>
                <div>
                  <span className="text-caption text-muted-foreground block">
                    Lotação e Unidade
                  </span>
                  <span className="font-medium text-foreground">{selectedItem.revenda || '—'}</span>
                  <span className="text-caption text-muted-foreground block mt-0.5">
                    Setor: {selectedItem.setor || '—'}{' '}
                    {selectedItem.cargo ? `• ${selectedItem.cargo}` : ''}
                  </span>
                </div>
              </div>
            </div>

            {/* Cláusula 2: Discriminação do Ativo */}
            <div className="space-y-3">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                2. Especificação do Equipamento Entregue
              </span>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 rounded border border-border bg-surface-alt p-4 text-caption">
                <div>
                  <span className="text-muted-foreground block">ID Patrimonial</span>
                  <span className="font-mono font-semibold text-foreground">
                    #{selectedItem.id}
                  </span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Tipo / Categoria</span>
                  <span className="font-medium text-foreground">{selectedItem.tipo || '—'}</span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Marca e Modelo</span>
                  <span className="font-medium text-foreground">
                    {selectedItem.brand} {selectedItem.model}
                  </span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Serial / Identificador</span>
                  <span className="font-mono text-foreground">
                    {selectedItem.identificador || '—'}
                  </span>
                </div>
              </div>
            </div>

            {/* Cláusula 3: Obrigações */}
            <div className="space-y-3 pt-2">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                3. Condições e Obrigações de Guarda
              </span>
              <div className="space-y-2 text-caption text-muted-foreground leading-relaxed border-l-2 border-border pl-4">
                <p>
                  <strong>3.1. Finalidade:</strong> O equipamento acima discriminado é de
                  propriedade exclusiva da Revalle e é cedido em comodato exclusivamente para
                  execução das atividades funcionais.
                </p>
                <p>
                  <strong>3.2. Conservação:</strong> O colaborador compromete-se a zelar pelo
                  perfeito estado de funcionamento e conservação do bem, comunicando imediatamente
                  ao departamento de TI qualquer falha, dano ou extravio.
                </p>
                <p>
                  <strong>3.3. Restituição:</strong> O colaborador compromete-se a restituir o
                  equipamento nas mesmas condições em que o recebeu no caso de rescisão contratual,
                  transferência ou solicitação expressa da empresa.
                </p>
              </div>
            </div>

            {/* Bloco de Assinatura Probatória */}
            <div className="pt-6 border-t border-border">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-8 pt-4">
                <div className="border-t border-border-strong pt-2 text-center text-caption">
                  <span className="font-semibold text-foreground block">
                    {selectedItem.assigned_to || 'Colaborador'}
                  </span>
                  <span className="text-muted-foreground">Custodiante Responsável</span>
                </div>
                <div className="border-t border-border-strong pt-2 text-center text-caption">
                  <span className="font-semibold text-foreground block">
                    Departamento de TI / Patrimônio
                  </span>
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
