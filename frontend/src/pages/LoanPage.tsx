import { useEffect, useState } from 'react'
import { useQuery, useMutation, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { listItemsPaginated } from '@/api/items'
import { initiateLoan } from '@/api/loans'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { StatusBadge } from '@/components/ui/badge'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { DataTable } from '@/components/ui/DataTable'
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { Stepper, type Step } from '@/components/ui/Stepper'
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { useConstants } from '@/hooks/useConstants'
import { formatDate, maskCpfInput, isValidCpf } from '@/lib/utils'
import { FileDown, ArrowLeft, ArrowRight, CheckCircle2 } from 'lucide-react'

// Esta tela tem dois blocos de dados, cada um com o uso certo (T3):
// - "Equipamento *" (abaixo) é uma lista de SELEÇÃO — o operador procura um
//   item "Disponível" para emprestar, não folheia página por página. Por
//   isso liga a busca (com debounce) ao parâmetro `search` do servidor, com
//   `limit` pequeno, em vez de buscar os 500 primeiros e filtrar no cliente.
// - "Empréstimos Pendentes de Confirmação" é uma tabela de NAVEGAÇÃO — lista
//   para o operador percorrer/procurar. Usa paginação real no servidor, no
//   mesmo padrão do DataTable que HistoryPage/StockPage já usam.
const ITEM_SEARCH_LIMIT = 20
const PENDENTES_PAGE_SIZE = 10

// Assistente de 4 passos (T6, garimpado do FLUXO da FT_STC — não do
// carregamento de dados dela, que buscava um lote fixo de itens e filtrava
// no cliente, ficando incorreto em silêncio assim que o total de itens
// passasse do teto de página do backend). `currentStep` mora aqui na página
// porque o `Stepper` (Task 1) é só a barra visual: recebe
// `steps`/`currentStep`/`onStepClick` e não sabe nada sobre empréstimo,
// equipamento ou validação.
const STEPS: Step[] = [
  { id: 'equipamento', label: 'Equipamento', description: 'Ativo disponível' },
  { id: 'colaborador', label: 'Colaborador', description: 'Responsável e CPF' },
  { id: 'condicoes', label: 'Condições', description: 'Revenda e setor' },
  { id: 'confirmacao', label: 'Confirmação', description: 'Revisão e envio' },
]

export default function LoanPage() {
  const queryClient = useQueryClient()
  const { centerCosts, setores, revendas, isLoading: constantsLoading } = useConstants()
  const [currentStep, setCurrentStep] = useState(0)
  const [selectedItemId, setSelectedItemId] = useState('')
  const [usuario, setUsuario] = useState('')
  const [cpf, setCpf] = useState('')
  const [centerCost, setCenterCost] = useState('')
  const [setor, setSetor] = useState('')
  const [cargo, setCargo] = useState('')
  const [revenda, setRevenda] = useState('')
  // Desmarcado por padrão: o termo assume "pessoa física" (o caso comum, já
  // que o campo ao lado no termo é CPF). Marcar apenas quando o recebedor for
  // de fato representante de pessoa jurídica.
  const [pessoaJuridica, setPessoaJuridica] = useState(false)
  const [dateIssue, setDateIssue] = useState(() => {
    const d = new Date()
    return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`
  })
  const [pendingItemId, setPendingItemId] = useState<number | null>(null)

  // Busca do equipamento disponível (lista de SELEÇÃO): `itemSearch` é o que
  // o operador digita, `itemSearchAplicado` é o que vai para o servidor, com
  // atraso, para não disparar uma requisição por tecla (mesmo padrão do
  // debounce de busca já usado em HistoryPage/StockPage).
  const [itemSearch, setItemSearch] = useState('')
  const [itemSearchAplicado, setItemSearchAplicado] = useState('')
  useEffect(() => {
    const timer = setTimeout(() => setItemSearchAplicado(itemSearch.trim()), 400)
    return () => clearTimeout(timer)
  }, [itemSearch])

  const { data: disponivelData } = useQuery({
    queryKey: ['items', 'select', 'Disponível', itemSearchAplicado],
    queryFn: () =>
      listItemsPaginated({ status: 'Disponível', search: itemSearchAplicado || undefined, limit: ITEM_SEARCH_LIMIT }),
    placeholderData: keepPreviousData,
  })
  const disponivel = disponivelData?.items ?? []
  const selectedItem = disponivel.find((i) => String(i.id) === selectedItemId)

  // Empréstimos pendentes de confirmação (tabela de NAVEGAÇÃO): paginação e
  // busca resolvidas no servidor, mesmo padrão do DataTable usado abaixo.
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
    queryKey: ['items', 'pendentes-confirmacao', pendentesPageIndex, pendentesSearchAplicado],
    queryFn: () =>
      listItemsPaginated({
        status: 'Pendente',
        search: pendentesSearchAplicado || undefined,
        limit: PENDENTES_PAGE_SIZE,
        offset: pendentesPageIndex * PENDENTES_PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })
  const pendentes = pendentesData?.items ?? []
  const pendentesTotal = pendentesData?.total ?? 0

  const loanMutation = useMutation({
    mutationFn: initiateLoan,
    onSuccess: (_, vars) => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setPendingItemId(vars.item_id)
      toast('Empréstimo iniciado! Agora gere e assine o termo.')
      // Reinicia o assistente para o próximo empréstimo — o painel de
      // confirmação (abaixo) segue seu próprio ciclo de vida independente,
      // via `pendingItemId`.
      setSelectedItemId('')
      setUsuario('')
      setCpf('')
      setCenterCost('')
      setSetor('')
      setCargo('')
      setRevenda('')
      setPessoaJuridica(false)
      setItemSearch('')
      setCurrentStep(0)
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao iniciar empréstimo.'), 'error')
    },
  })

  // Validação por etapa (T6): cada passo só libera o "Próximo" com seus
  // próprios campos válidos — CPF no passo 2 (Colaborador), revenda e setor
  // no passo 3 (Condições). O Stepper em si não valida nada, só desenha a
  // barra; a decisão de avançar mora aqui.
  function handleNext() {
    if (currentStep === 0) {
      if (!selectedItemId) { toast('Selecione um equipamento disponível.', 'error'); return }
      // Preenchimento automático da revenda a partir do item selecionado (T6).
      if (selectedItem?.revenda && !revenda) setRevenda(selectedItem.revenda)
      setCurrentStep(1)
      return
    }
    if (currentStep === 1) {
      if (!usuario.trim()) { toast('Informe o nome do colaborador.', 'error'); return }
      // isValidCpf confere os dígitos verificadores (T3) — a versão antiga só
      // checava a contagem de dígitos e aceitava sequências como
      // 111.111.111-11, que nunca são CPFs reais.
      if (!isValidCpf(cpf)) { toast('CPF inválido.', 'error'); return }
      setCurrentStep(2)
      return
    }
    if (currentStep === 2) {
      if (!revenda) { toast('Selecione a revenda.', 'error'); return }
      if (!setor) { toast('Selecione o setor.', 'error'); return }
      setCurrentStep(3)
    }
  }

  function handlePrev() {
    if (currentStep > 0) setCurrentStep(currentStep - 1)
  }

  function handleConfirmar() {
    // Reconfere tudo no envio final — mesma validação de sempre — caso o
    // operador tenha voltado e alterado algum campo já validado.
    if (!selectedItemId) { toast('Selecione um item.', 'error'); return }
    if (!isValidCpf(cpf)) { toast('CPF inválido.', 'error'); return }
    if (!revenda) { toast('Selecione a revenda.', 'error'); return }
    if (!setor) { toast('Selecione o setor.', 'error'); return }
    loanMutation.mutate({
      item_id: Number(selectedItemId), usuario, cpf, center_cost: centerCost, cargo, setor, revenda,
      date_issue: dateIssue, pessoa_juridica: pessoaJuridica,
    })
  }

  const pendingColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Data', cell: ({ getValue }) => <span className="num">{formatDate(getValue() as string)}</span> },
    {
      id: 'actions',
      header: 'Ações',
      cell: ({ row }) => (
        <div className="flex gap-2">
          <Button size="sm" variant="outline" onClick={() => generateAndDownloadLoanTerm(row.original.id)}>
            <FileDown size={14} />Termo
          </Button>
        </div>
      ),
    },
  ]

  const isLastStep = currentStep === STEPS.length - 1

  return (
    <div className="page-container-reading space-y-8">
      <PageHeader
        eyebrow="Movimentação de Ativos"
        eyebrowDetail="Concessão de Equipamento"
        title="Emprestar Equipamento"
        description="Siga as quatro etapas para preencher os dados e iniciar o empréstimo."
      />

      <div className="surface-panel p-6 space-y-6">
        <PanelHeader
          title="Dados do Empréstimo"
          description="Informe o equipamento e os dados do colaborador responsável."
        />

        <Stepper
          steps={STEPS}
          currentStep={currentStep}
          // Só permite voltar para uma etapa já concluída — nunca pular para
          // uma etapa futura ainda não validada.
          onStepClick={(step) => { if (step < currentStep) setCurrentStep(step) }}
        />

        {/* Grid de 2 colunas: painel da etapa atual (flexível) + coluna de
            contexto fixa de 340px com o resumo do que já foi escolhido. */}
        <div className="grid grid-cols-1 lg:grid-cols-[1fr_340px] gap-6 items-start">
          <div className="rounded border border-border bg-surface-alt p-5 space-y-5">
            {currentStep === 0 && (
              <div className="space-y-4">
                <h3 className="text-body-lg font-semibold text-foreground">1. Equipamento</h3>
                <div className="flex flex-col gap-1.5">
                  <Label>Equipamento *</Label>
                  {/* Busca no servidor (debounce de 400ms) controlando o próprio
                      campo de busca do dropdown via `search`/`onSearchChange`
                      — sem isso, o usuário veria uma segunda caixa de busca (a
                      interna do SearchableSelect) além desta, filtrando só os
                      20 já carregados. */}
                  <SearchableSelect
                    options={disponivel.map((i) => ({
                      value: String(i.id),
                      label: `#${i.id} — ${i.tipo} ${i.brand || ''} ${i.model || ''}`,
                      subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
                    }))}
                    value={selectedItemId}
                    onValueChange={setSelectedItemId}
                    placeholder="Selecione um equipamento disponível..."
                    searchPlaceholder="Buscar por marca, modelo ou identificador (patrimônio)..."
                    search={itemSearch}
                    onSearchChange={setItemSearch}
                  />
                </div>
              </div>
            )}

            {currentStep === 1 && (
              <div className="space-y-4">
                <h3 className="text-body-lg font-semibold text-foreground">2. Colaborador</h3>
                <div className="grid grid-cols-2 gap-4">
                  <div className="flex flex-col gap-1.5">
                    <Label>Funcionário *</Label>
                    <Input value={usuario} onChange={e => setUsuario(e.target.value)} required />
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <Label>CPF *</Label>
                    <Input value={cpf} onChange={e => setCpf(maskCpfInput(e.target.value))} placeholder="000.000.000-00" maxLength={14} required />
                    <label className="flex items-center gap-2 text-body-sm text-foreground cursor-pointer select-none mt-1">
                      <input
                        type="checkbox"
                        checked={pessoaJuridica}
                        onChange={(e) => setPessoaJuridica(e.target.checked)}
                        className="rounded-sm border-border-strong"
                      />
                      É pessoa jurídica
                    </label>
                  </div>
                </div>
              </div>
            )}

            {currentStep === 2 && (
              <div className="space-y-4">
                <h3 className="text-body-lg font-semibold text-foreground">3. Condições</h3>
                <div className="grid grid-cols-2 gap-4">
                  <div className="flex flex-col gap-1.5">
                    <Label>Cargo *</Label>
                    <Input value={cargo} onChange={e => setCargo(e.target.value)} required />
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <Label>Centro de Custo *</Label>
                    <Select value={centerCost} onValueChange={setCenterCost} required disabled={constantsLoading}>
                      <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
                      <SelectContent>{centerCosts.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent>
                    </Select>
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <Label>Setor *</Label>
                    <Select value={setor} onValueChange={setSetor} required disabled={constantsLoading}>
                      <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
                      <SelectContent>{setores.map(s => <SelectItem key={s} value={s}>{s}</SelectItem>)}</SelectContent>
                    </Select>
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <Label>Revenda *</Label>
                    {/* Pré-preenchida a partir do item ao sair da etapa 1 (T6)
                        — o operador ainda pode trocar aqui se necessário. */}
                    <Select value={revenda} onValueChange={setRevenda} required disabled={constantsLoading}>
                      <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
                      <SelectContent>{revendas.map(r => <SelectItem key={r} value={r}>{r}</SelectItem>)}</SelectContent>
                    </Select>
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <Label>Data do Empréstimo *</Label>
                    <Input value={dateIssue} onChange={e => setDateIssue(e.target.value)} placeholder="dd/mm/aaaa" required />
                  </div>
                </div>
              </div>
            )}

            {currentStep === 3 && (
              <div className="space-y-4">
                <h3 className="text-body-lg font-semibold text-foreground">4. Confirmação</h3>
                <p className="text-body-sm text-muted-foreground">
                  Revise os dados abaixo antes de iniciar o empréstimo.
                </p>
                <div className="rounded border border-border bg-surface p-4 space-y-2.5 text-body-sm">
                  <div className="flex justify-between border-b border-border pb-1.5">
                    <span className="text-muted-foreground">Equipamento</span>
                    <span className="font-semibold text-foreground">
                      #{selectedItem?.id} — {selectedItem?.tipo} {selectedItem?.brand} {selectedItem?.model}
                    </span>
                  </div>
                  <div className="flex justify-between border-b border-border pb-1.5">
                    <span className="text-muted-foreground">Colaborador</span>
                    <span className="font-semibold text-foreground">{usuario}{pessoaJuridica ? ' (Pessoa Jurídica)' : ''}</span>
                  </div>
                  <div className="flex justify-between border-b border-border pb-1.5">
                    <span className="text-muted-foreground">CPF</span>
                    <span className="font-mono text-foreground">{cpf}</span>
                  </div>
                  <div className="flex justify-between border-b border-border pb-1.5">
                    <span className="text-muted-foreground">Revenda / Setor</span>
                    <span className="text-foreground">{revenda} • {setor}</span>
                  </div>
                  <div className="flex justify-between border-b border-border pb-1.5">
                    <span className="text-muted-foreground">Data do Empréstimo</span>
                    <span className="font-mono text-foreground">{dateIssue}</span>
                  </div>
                  <div className="flex justify-between items-center">
                    <span className="text-muted-foreground">Status após iniciar</span>
                    <StatusBadge status="Pendente" />
                  </div>
                </div>
              </div>
            )}

            <div className="flex justify-between pt-4 border-t border-border">
              <Button type="button" variant="outline" onClick={handlePrev} disabled={currentStep === 0 || loanMutation.isPending}>
                <ArrowLeft size={14} />Voltar
              </Button>
              {isLastStep ? (
                <Button type="button" onClick={handleConfirmar} disabled={loanMutation.isPending}>
                  <CheckCircle2 size={14} />
                  {loanMutation.isPending ? 'Iniciando...' : 'Iniciar Empréstimo'}
                </Button>
              ) : (
                <Button type="button" onClick={handleNext}>
                  Próximo<ArrowRight size={14} />
                </Button>
              )}
            </div>
          </div>

          {/* Coluna de contexto: resumo do que já foi escolhido nas etapas
              anteriores, fixa em 340px nas telas grandes. */}
          <div className="rounded border border-border bg-surface p-5 space-y-4">
            <PanelHeader title="Resumo" description="Dados confirmados nas etapas concluídas." />
            {currentStep === 0 && (
              <p className="text-body-sm text-muted-foreground italic">
                Selecione um equipamento na etapa 1 para começar o resumo.
              </p>
            )}
            {currentStep > 0 && (
              <div className="space-y-1.5 pb-3 border-b border-border text-body-sm">
                <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">Equipamento</span>
                <p className="font-semibold text-foreground">
                  {selectedItem ? `#${selectedItem.id} — ${selectedItem.tipo} ${selectedItem.brand || ''}` : '—'}
                </p>
                <p className="text-caption text-muted-foreground">{selectedItem?.identificador || '—'}</p>
              </div>
            )}
            {currentStep > 1 && (
              <div className="space-y-1.5 pb-3 border-b border-border text-body-sm">
                <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">Colaborador</span>
                <p className="font-semibold text-foreground">{usuario || '—'}</p>
                <p className="font-mono text-caption text-muted-foreground">
                  {cpf}{pessoaJuridica ? ' • Pessoa jurídica' : ''}
                </p>
              </div>
            )}
            {currentStep > 2 && (
              <div className="space-y-1.5 text-body-sm">
                <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">Condições</span>
                <p className="text-foreground">{revenda} • {setor}</p>
                <p className="text-caption text-muted-foreground">{cargo || '—'} • {dateIssue}</p>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Confirmação de empréstimo pendente (T5: painel compartilhado com TermsPage) */}
      {pendingItemId && (
        <ConfirmacaoTermo
          itemId={pendingItemId}
          description={
            <>
              1. Clique em "Gerar Termo" para baixar o documento<br />
              2. Imprima, assine e escaneie como PDF<br />
              3. Faça o upload do PDF assinado abaixo e confirme
            </>
          }
          showGenerateButton
          onConfirmed={() => setPendingItemId(null)}
          onCancel={() => setPendingItemId(null)}
        />
      )}

      {/* Lista de empréstimos pendentes de confirmação (tabela de NAVEGAÇÃO,
          paginada no servidor) */}
      {pendentesTotal > 0 && (
        <div className="space-y-3">
          <h3 className="text-body-lg font-semibold text-foreground">Empréstimos Pendentes de Confirmação</h3>
          <DataTable
            data={pendentes}
            columns={pendingColumns}
            searchPlaceholder="Buscar por marca, modelo, usuário..."
            pagination={{
              total: pendentesTotal,
              pageIndex: pendentesPageIndex,
              pageSize: PENDENTES_PAGE_SIZE,
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
