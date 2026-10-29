import { useEffect, useState } from 'react'
import { useQuery, useMutation, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import {
  FileDown,
  PackageCheck,
  ClipboardCheck,
  ArrowRight,
  ArrowLeft,
  CheckCircle2,
} from 'lucide-react'

import { listItemsPaginated } from '@/api/items'
import { initiateLoan } from '@/api/loans'
import { useConstants } from '@/hooks/useConstants'
import { getErrorMessage } from '@/lib/api-error'
import { isValidCpf, maskCpfInput } from '@/lib/utils'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Checkbox } from '@/components/ui/checkbox'
import { Badge } from '@/components/ui/badge'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import { Stepper, type Step } from '@/components/ui/Stepper'
import { toast } from '@/components/ui/toast'

// Layout visual idêntico ao da FT_STC; a lógica de dados é a da main:
// - "Equipamento em Estoque *" é uma lista de SELEÇÃO — a busca (com
//   debounce) vai para o parâmetro `search` do servidor, com `limit` pequeno,
//   em vez de buscar um lote fixo de itens e filtrar no cliente (que ficava
//   incorreto em silêncio quando o total passava do teto de página).
// - "Empréstimos Pendentes de Termo" é uma tabela de NAVEGAÇÃO — paginação
//   real no servidor (limit/offset), com o visual de tabela da FT_STC.
// - Os contadores do cabeçalho usam o `total` devolvido pelo servidor.
const ITEM_SEARCH_LIMIT = 20
const PENDENTES_PAGE_SIZE = 7

// Assistente de 4 passos: `currentStep` mora aqui na página porque o
// `Stepper` é só a barra visual — não sabe nada de empréstimo nem valida.
const STEPS: Step[] = [
  { id: 'equipamento', label: 'Equipamento', description: 'Ativo disponível' },
  { id: 'usuario', label: 'Colaborador', description: 'Responsável e CPF' },
  { id: 'condicoes', label: 'Condições', description: 'Setor e centro de custo' },
  { id: 'confirmacao', label: 'Confirmação', description: 'Emissão de termo' },
]

export default function LoanPage() {
  const queryClient = useQueryClient()
  // Setores, revendas e centros de custo vêm do backend (inclui os setores
  // novos Frota, Logística e Controle) — nada de lista fixa no cliente.
  const { centerCosts, setores, revendas, isLoading: constantsLoading } = useConstants()

  const [currentStep, setCurrentStep] = useState(0)

  // Estado do formulário
  const [selectedItemId, setSelectedItemId] = useState('')
  const [usuario, setUsuario] = useState('')
  const [cpf, setCpf] = useState('')
  const [centerCost, setCenterCost] = useState('')
  const [setor, setSetor] = useState('')
  const [cargo, setCargo] = useState('')
  const [revenda, setRevenda] = useState('')
  // Desmarcado por padrão: o termo assume "pessoa física" (o caso comum, já
  // que o campo ao lado no termo é CPF). Marcar apenas quando o recebedor for
  // de fato representante de pessoa jurídica — muda o texto legal do termo.
  const [pessoaJuridica, setPessoaJuridica] = useState(false)
  const [dateIssue, setDateIssue] = useState(() => {
    const d = new Date()
    return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`
  })
  const [pendingItemId, setPendingItemId] = useState<number | null>(null)

  // Busca do equipamento disponível: `itemSearch` é o que o operador digita,
  // `itemSearchAplicado` é o que vai para o servidor, com atraso, para não
  // disparar uma requisição por tecla.
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

  // Contador "Disponíveis" do cabeçalho: total real no servidor, sem a busca
  // do seletor (que alteraria o número enquanto o operador digita).
  const { data: disponivelCountData } = useQuery({
    queryKey: ['items', 'contagem', 'Disponível'],
    queryFn: () => listItemsPaginated({ status: 'Disponível', limit: 1 }),
  })
  const disponivelTotal = disponivelCountData?.total ?? 0

  // Empréstimos pendentes de termo (tabela de NAVEGAÇÃO, paginada no servidor)
  const [pendentesPage, setPendentesPage] = useState(0)
  const { data: pendentesData } = useQuery({
    queryKey: ['items', 'pendentes-confirmacao', pendentesPage],
    queryFn: () =>
      listItemsPaginated({
        status: 'Pendente',
        limit: PENDENTES_PAGE_SIZE,
        offset: pendentesPage * PENDENTES_PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })
  const pendentes = pendentesData?.items ?? []
  const pendentesTotal = pendentesData?.total ?? 0
  const pendentesTotalPages = Math.ceil(pendentesTotal / PENDENTES_PAGE_SIZE)
  // Se a última página esvaziar (termo anexado em outro lugar), volta para a
  // última página que ainda existe em vez de mostrar uma tabela vazia.
  useEffect(() => {
    if (pendentesTotalPages > 0 && pendentesPage > pendentesTotalPages - 1) {
      setPendentesPage(pendentesTotalPages - 1)
    }
  }, [pendentesPage, pendentesTotalPages])

  const loanMutation = useMutation({
    mutationFn: initiateLoan,
    onSuccess: (_, vars) => {
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      setPendingItemId(vars.item_id)
      toast('Empréstimo iniciado! Agora gere e assine o termo.')
      // Reinicia o assistente para o próximo empréstimo — o painel de
      // confirmação segue seu próprio ciclo de vida, via `pendingItemId`.
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

  // Validação por etapa: cada passo só libera o "Próximo" com seus próprios
  // campos válidos — CPF no passo 2 (Colaborador), revenda e setor no passo 3
  // (Condições).
  function handleNext() {
    if (currentStep === 0) {
      if (!selectedItemId) { toast('Selecione um equipamento disponível.', 'error'); return }
      // Preenchimento automático da revenda a partir do item selecionado.
      if (selectedItem?.revenda && !revenda) setRevenda(selectedItem.revenda)
      setCurrentStep(1)
      return
    }
    if (currentStep === 1) {
      if (!usuario.trim()) { toast('Informe o nome do colaborador.', 'error'); return }
      // isValidCpf confere os dígitos verificadores (rejeita 111.111.111-11 etc.).
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
    // Reconfere tudo no envio final, caso o operador tenha voltado e
    // alterado algum campo já validado.
    if (!selectedItemId) { toast('Selecione um item.', 'error'); return }
    if (!isValidCpf(cpf)) { toast('CPF inválido.', 'error'); return }
    if (!revenda) { toast('Selecione a revenda.', 'error'); return }
    if (!setor) { toast('Selecione o setor.', 'error'); return }
    loanMutation.mutate({
      item_id: Number(selectedItemId),
      usuario,
      cpf,
      center_cost: centerCost,
      cargo,
      setor,
      revenda,
      date_issue: dateIssue,
      pessoa_juridica: pessoaJuridica,
    })
  }

  const isLastStep = currentStep === STEPS.length - 1

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
            <span className="text-caption text-foreground font-medium">Fluxo Guiado</span>
          </div>
          <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
            Empréstimo de Equipamento
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Alocação de ativo com validação de colaborador e emissão de termo de responsabilidade.
          </p>
        </div>

        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded border border-border bg-surface text-caption">
            <PackageCheck className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Disponíveis:</span>
            <span className="font-mono font-semibold text-foreground">{disponivelTotal}</span>
          </div>
          <div className="flex items-center gap-2 px-3 py-1.5 rounded border border-border bg-surface text-caption">
            <ClipboardCheck className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Pendentes:</span>
            <span className="font-mono font-semibold text-foreground">{pendentesTotal}</span>
          </div>
        </div>
      </div>

      {/* Painel de confirmação pós-empréstimo (compartilhado com TermsPage) */}
      {pendingItemId && (
        <ConfirmacaoTermo
          itemId={pendingItemId}
          description="Faça o download do termo de responsabilidade assinado pelo colaborador (PDF) ou anexe diretamente."
          uploadLabel="Arraste ou clique para anexar o termo assinado"
          errorMessage="Erro ao confirmar termo de empréstimo."
          showGenerateButton
          onConfirmed={() => setPendingItemId(null)}
          onCancel={() => setPendingItemId(null)}
        />
      )}

      {/* Stepper superior */}
      <div className="rounded border border-border bg-surface p-5">
        <Stepper
          steps={STEPS}
          currentStep={currentStep}
          // Só permite voltar para uma etapa já concluída — nunca pular para
          // uma etapa futura ainda não validada.
          onStepClick={(step) => {
            if (step < currentStep) setCurrentStep(step)
          }}
        />
      </div>

      {/* Grid de 2 colunas: formulário da etapa atual + contexto fixo */}
      <div className="grid grid-cols-1 lg:grid-cols-[1fr_340px] gap-6 items-start">
        {/* Painel da etapa atual */}
        <div className="rounded border border-border bg-surface p-6 space-y-6">
          {/* ETAPA 1: EQUIPAMENTO */}
          {currentStep === 0 && (
            <div className="space-y-4">
              <div>
                <h2 className="text-body-lg font-semibold text-foreground">
                  1. Seleção do Equipamento
                </h2>
                <p className="text-caption text-muted-foreground mt-0.5">
                  Escolha o ativo patrimonial disponível para alocação.
                </p>
              </div>

              <div className="space-y-2 pt-2 border-t border-border">
                <Label className="text-caption font-medium text-foreground">
                  Equipamento em Estoque *
                </Label>
                {/* Busca no servidor (debounce de 400ms) controlando o próprio
                    campo de busca do dropdown via `search`/`onSearchChange`. */}
                <SearchableSelect
                  options={disponivel.map((i) => ({
                    value: String(i.id),
                    label: `#${i.id} — ${i.tipo ?? ''} ${i.brand ?? ''} ${i.model ?? ''}`.trim(),
                    subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
                  }))}
                  value={selectedItemId}
                  onValueChange={setSelectedItemId}
                  placeholder="Buscar por ID, modelo, tipo ou serial…"
                  searchPlaceholder="Filtrar equipamentos disponíveis…"
                  search={itemSearch}
                  onSearchChange={setItemSearch}
                />
              </div>

              {selectedItem && (
                <div className="rounded border border-border bg-surface-alt p-4 space-y-2 text-caption">
                  <div className="flex justify-between border-b border-border pb-1.5">
                    <span className="text-muted-foreground">Tipo / Modelo</span>
                    <span className="font-semibold text-foreground">
                      {selectedItem.tipo} {selectedItem.brand} {selectedItem.model}
                    </span>
                  </div>
                  <div className="flex justify-between border-b border-border pb-1.5">
                    <span className="text-muted-foreground">Serial / Identificador</span>
                    <span className="font-mono text-foreground">
                      {selectedItem.identificador || '—'}
                    </span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted-foreground">Unidade de Origem</span>
                    <span className="text-foreground">{selectedItem.revenda || '—'}</span>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ETAPA 2: COLABORADOR */}
          {currentStep === 1 && (
            <div className="space-y-4">
              <div>
                <h2 className="text-body-lg font-semibold text-foreground">
                  2. Dados do Colaborador Responsável
                </h2>
                <p className="text-caption text-muted-foreground mt-0.5">
                  Informe quem receberá a custódia do equipamento.
                </p>
              </div>

              <div className="space-y-4 pt-2 border-t border-border">
                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">
                    Nome Completo do Colaborador *
                  </Label>
                  <Input
                    value={usuario}
                    onChange={(e) => setUsuario(e.target.value)}
                    placeholder="Ex: João da Silva"
                    required
                    className="h-9"
                  />
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <Label className="text-caption font-medium text-foreground">
                      CPF do Colaborador *
                    </Label>
                    <Input
                      value={cpf}
                      onChange={(e) => setCpf(maskCpfInput(e.target.value))}
                      placeholder="000.000.000-00"
                      maxLength={14}
                      required
                      className="font-mono h-9"
                    />
                  </div>

                  {/* Toggle "É pessoa jurídica": muda apenas o texto legal do
                      termo. O rótulo da main foi mantido (o da FT_STC citava
                      "terceirizado", que nem sempre é pessoa jurídica). */}
                  <div className="flex items-center gap-2 pt-6">
                    <Checkbox
                      id="pj"
                      checked={pessoaJuridica}
                      onCheckedChange={setPessoaJuridica}
                    />
                    <label
                      htmlFor="pj"
                      className="text-caption text-foreground cursor-pointer select-none"
                    >
                      É pessoa jurídica
                    </label>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ETAPA 3: CONDIÇÕES */}
          {currentStep === 2 && (
            <div className="space-y-4">
              <div>
                <h2 className="text-body-lg font-semibold text-foreground">
                  3. Condições Operacionais e Alocação
                </h2>
                <p className="text-caption text-muted-foreground mt-0.5">
                  Unidade de lotação, centro de custo e cargo do colaborador.
                </p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-2 border-t border-border">
                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">
                    Revenda / Filial *
                  </Label>
                  {/* Pré-preenchida a partir do item ao sair da etapa 1 — o
                      operador ainda pode trocar aqui se necessário. */}
                  <Select value={revenda} onValueChange={setRevenda} disabled={constantsLoading}>
                    <SelectTrigger className="h-9">
                      <SelectValue placeholder="Selecione a revenda" />
                    </SelectTrigger>
                    <SelectContent>
                      {revendas.map((r) => (
                        <SelectItem key={r} value={r}>
                          {r}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">Setor *</Label>
                  <Select value={setor} onValueChange={setSetor} disabled={constantsLoading}>
                    <SelectTrigger className="h-9">
                      <SelectValue placeholder="Selecione o setor" />
                    </SelectTrigger>
                    <SelectContent>
                      {setores.map((s) => (
                        <SelectItem key={s} value={s}>
                          {s}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">Cargo / Função</Label>
                  <Input
                    value={cargo}
                    onChange={(e) => setCargo(e.target.value)}
                    placeholder="Ex: Analista Comercial"
                    className="h-9"
                  />
                </div>

                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">
                    Centro de Custo
                  </Label>
                  <Select value={centerCost} onValueChange={setCenterCost} disabled={constantsLoading}>
                    <SelectTrigger className="h-9">
                      <SelectValue placeholder="Selecione o centro de custo" />
                    </SelectTrigger>
                    <SelectContent>
                      {centerCosts.map((c) => (
                        <SelectItem key={c} value={c}>
                          {c}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-1.5">
                  <Label className="text-caption font-medium text-foreground">
                    Data da Entrega
                  </Label>
                  <Input
                    value={dateIssue}
                    onChange={(e) => setDateIssue(e.target.value)}
                    placeholder="DD/MM/AAAA"
                    className="font-mono h-9"
                  />
                </div>
              </div>
            </div>
          )}

          {/* ETAPA 4: CONFIRMAÇÃO E REVISÃO */}
          {currentStep === 3 && (
            <div className="space-y-4">
              <div>
                <h2 className="text-body-lg font-semibold text-foreground">
                  4. Revisão e Emissão do Termo
                </h2>
                <p className="text-caption text-muted-foreground mt-0.5">
                  Confira as informações antes de formalizar a entrega patrimonial.
                </p>
              </div>

              <div className="rounded border border-border bg-surface-alt p-4 space-y-3 text-caption">
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">Equipamento</span>
                  <span className="font-semibold text-foreground">
                    #{selectedItem?.id} — {selectedItem?.tipo} {selectedItem?.brand}{' '}
                    {selectedItem?.model}
                  </span>
                </div>
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">Colaborador</span>
                  <span className="font-semibold text-foreground">{usuario}</span>
                </div>
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">CPF</span>
                  <span className="font-mono text-foreground">{cpf}</span>
                </div>
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">Revenda / Setor</span>
                  <span className="text-foreground">
                    {revenda} • {setor}
                  </span>
                </div>
                <div className="flex justify-between border-b border-border pb-1.5">
                  <span className="text-muted-foreground">Data Empréstimo</span>
                  <span className="font-mono text-foreground">{dateIssue}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-muted-foreground">Status Pós-Registro</span>
                  <Badge
                    variant="warning"
                    showDot
                    className="font-mono text-[10px] border-border text-foreground"
                  >
                    PENDENTE TERMO
                  </Badge>
                </div>
              </div>
            </div>
          )}

          {/* Controles de navegação do assistente */}
          <div className="pt-4 border-t border-border flex items-center justify-between">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={handlePrev}
              disabled={currentStep === 0 || loanMutation.isPending}
            >
              <ArrowLeft className="mr-1.5 size-3.5" />
              Anterior
            </Button>

            {!isLastStep ? (
              <Button type="button" size="sm" onClick={handleNext}>
                Próximo
                <ArrowRight className="ml-1.5 size-3.5" />
              </Button>
            ) : (
              <Button
                type="button"
                size="sm"
                onClick={handleConfirmar}
                disabled={loanMutation.isPending}
              >
                <CheckCircle2 className="mr-1.5 size-3.5" />
                {loanMutation.isPending ? 'Processando…' : 'Confirmar Empréstimo'}
              </Button>
            )}
          </div>
        </div>

        {/* Coluna lateral de contexto persistente */}
        <div className="rounded border border-border bg-surface p-5 space-y-4">
          <div className="border-b border-border pb-2.5">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Contexto do Ativo
            </span>
            <p className="text-body-sm font-semibold text-foreground mt-0.5">
              {selectedItem
                ? `#${selectedItem.id} — ${selectedItem.tipo}`
                : 'Nenhum ativo selecionado'}
            </p>
          </div>

          {selectedItem ? (
            <div className="space-y-2.5 text-caption">
              <div>
                <span className="text-muted-foreground block">Marca / Modelo</span>
                <span className="font-medium text-foreground">
                  {selectedItem.brand} {selectedItem.model}
                </span>
              </div>
              <div>
                <span className="text-muted-foreground block">Serial</span>
                <span className="font-mono text-foreground">
                  {selectedItem.identificador || '—'}
                </span>
              </div>
              <div>
                <span className="text-muted-foreground block">Filial Atual</span>
                <span className="text-foreground">{selectedItem.revenda || '—'}</span>
              </div>
              <div>
                <span className="text-muted-foreground block">Nota Fiscal</span>
                <span className="font-mono text-foreground">{selectedItem.nota_fiscal || '—'}</span>
              </div>
            </div>
          ) : (
            <p className="text-caption text-muted-foreground italic">
              Selecione um equipamento na etapa 1 para manter seus dados visíveis durante todo o processo.
            </p>
          )}

          <div className="pt-3 border-t border-border">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block mb-2">
              Regras do Fluxo
            </span>
            <ul className="space-y-1.5 text-caption text-muted-foreground leading-relaxed">
              <li>
                • O status passará para <strong>Pendente</strong> até que o termo assinado seja anexado.
              </li>
              <li>• O CPF é validado pelos dígitos verificadores oficiais.</li>
              <li>
                • O termo de responsabilidade pode ser gerado e baixado imediatamente após a confirmação.
              </li>
            </ul>
          </div>
        </div>
      </div>

      {/* Empréstimos pendentes de termo (paginados no servidor) */}
      {pendentesTotal > 0 && (
        <div className="rounded border border-border bg-surface p-5 space-y-3">
          <div className="flex items-center justify-between border-b border-border pb-2">
            <div>
              <h3 className="text-body font-semibold text-foreground">
                Empréstimos Pendentes de Termo ({pendentesTotal})
              </h3>
              <p className="text-caption text-muted-foreground">
                Equipamentos entregues que ainda aguardam a anexação do PDF assinado.
              </p>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-body-sm">
              <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground">
                <tr>
                  <th className="py-2 px-3 font-semibold">ID</th>
                  <th className="py-2 px-3 font-semibold">Tipo</th>
                  <th className="py-2 px-3 font-semibold">Colaborador</th>
                  <th className="py-2 px-3 font-semibold">Revenda</th>
                  <th className="py-2 px-3 text-right font-semibold">Ação</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {pendentes.map((item) => (
                  <tr key={item.id} className="hover:bg-surface-alt transition-colors">
                    <td className="py-2.5 px-3 font-mono font-medium">#{item.id}</td>
                    <td className="py-2.5 px-3">
                      {item.tipo} {item.brand}
                    </td>
                    <td className="py-2.5 px-3 font-medium">{item.assigned_to}</td>
                    <td className="py-2.5 px-3 text-muted-foreground">{item.revenda || '—'}</td>
                    <td className="py-2.5 px-3 text-right">
                      <div className="flex justify-end gap-2">
                        <Button
                          size="sm"
                          variant="outline"
                          onClick={() => generateAndDownloadLoanTerm(item.id)}
                          className="text-xs h-7 px-2.5"
                        >
                          <FileDown className="mr-1 size-3" />
                          Termo
                        </Button>
                        <Button
                          size="sm"
                          onClick={() => setPendingItemId(item.id)}
                          className="text-xs h-7 px-2.5"
                        >
                          Anexar Termo
                        </Button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Paginação dos pendentes */}
          {pendentesTotalPages > 1 && (
            <div className="flex items-center justify-between pt-2 border-t border-border text-xs text-muted-foreground">
              <span>
                Página {pendentesPage + 1} de {pendentesTotalPages}
              </span>
              <div className="flex gap-1">
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pendentesPage === 0}
                  onClick={() => setPendentesPage((p) => Math.max(0, p - 1))}
                  className="h-6 px-2 text-xs"
                >
                  ←
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pendentesPage >= pendentesTotalPages - 1}
                  onClick={() => setPendentesPage((p) => p + 1)}
                  className="h-6 px-2 text-xs"
                >
                  →
                </Button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  )
}
