import { useEffect, useState } from 'react'
import { useQuery, useMutation, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { listItemsPaginated } from '@/api/items'
import { initiateLoan } from '@/api/loans'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { toast } from '@/components/ui/toast'
import { DataTable } from '@/components/ui/DataTable'
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { useConstants } from '@/hooks/useConstants'
import { formatDate, maskCpfInput, isValidCpf } from '@/lib/utils'
import { FileDown } from 'lucide-react'

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

export default function LoanPage() {
  const queryClient = useQueryClient()
  const { centerCosts, setores, revendas, isLoading: constantsLoading } = useConstants()
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
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao iniciar empréstimo.'
      toast(msg, 'error')
    },
  })

  function handleLoanSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (!selectedItemId) { toast('Selecione um item.', 'error'); return }
    // isValidCpf agora confere os dígitos verificadores (T3) — a versão
    // antiga só checava a contagem de dígitos e aceitava sequências como
    // 111.111.111-11, que nunca são CPFs reais.
    if (!isValidCpf(cpf)) { toast('CPF inválido.', 'error'); return }
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

  return (
    <div className="page-container-reading space-y-8">
      <PageHeader
        eyebrow="Movimentação de Ativos"
        eyebrowDetail="Concessão de Equipamento"
        title="Emprestar Equipamento"
        description="Preencha os dados e inicie o processo de empréstimo."
      />

      <form onSubmit={handleLoanSubmit} className="surface-panel p-6 space-y-4">
        <PanelHeader
          title="Dados do Empréstimo"
          description="Informe o equipamento e os dados do colaborador responsável."
        />
        <div className="flex flex-col gap-1.5">
          <Label>Equipamento *</Label>
          {/* Busca no servidor (debounce de 400ms): digite para procurar entre
              todos os equipamentos "Disponível", não só os 20 exibidos abaixo. */}
          <Input
            value={itemSearch}
            onChange={(e) => setItemSearch(e.target.value)}
            placeholder="Buscar por marca, modelo ou identificador (patrimônio)..."
          />
          <SearchableSelect
            options={disponivel.map((i) => ({
              value: String(i.id),
              label: `#${i.id} — ${i.tipo} ${i.brand || ''} ${i.model || ''}`,
              subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
            }))}
            value={selectedItemId}
            onValueChange={setSelectedItemId}
            placeholder="Selecione um equipamento disponível..."
            searchPlaceholder="Filtrar nos resultados abaixo..."
          />
        </div>

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

        <div className="flex justify-end pt-2">
          <Button type="submit" disabled={loanMutation.isPending}>
            {loanMutation.isPending ? 'Iniciando...' : 'Iniciar Empréstimo'}
          </Button>
        </div>
      </form>

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
