import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
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
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { useConstants } from '@/hooks/useConstants'
import { formatDate, maskCpfInput, isValidCpf } from '@/lib/utils'
import { FileDown } from 'lucide-react'

// Sem paginação nesta tela (filtra "Disponível"/"Pendente" client-side a
// partir da lista completa) — usamos o teto de página do backend para não
// truncar em 50 itens (default de GET /api/items) como aconteceria chamando
// listItemsPaginated() sem limit.
const FETCH_ALL_LIMIT = 500

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

  const { data } = useQuery({
    queryKey: ['items'],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  })
  const items = data?.items ?? []

  const disponivel = items.filter(i => i.status === 'Disponível')
  const pendentes = items.filter(i => i.status === 'Pendente')

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
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Data', cell: ({ getValue }) => formatDate(getValue() as string) },
    {
      id: 'actions',
      header: 'Ações',
      cell: ({ row }) => (
        <div className="flex gap-2">
          <Button size="sm" variant="outline" onClick={() => generateAndDownloadLoanTerm(row.original.id)}>
            <FileDown size={14} className="mr-1" />Termo
          </Button>
        </div>
      ),
    },
  ]

  return (
    <div className="space-y-8 max-w-3xl">
      <div>
        <h2 className="text-xl font-semibold">Emprestar Equipamento</h2>
        <p className="text-sm text-muted-foreground">Preencha os dados e inicie o processo de empréstimo.</p>
      </div>

      <form onSubmit={handleLoanSubmit} className="bg-card rounded-xl border p-6 space-y-4">
        <div className="flex flex-col gap-1.5">
          <Label>Equipamento *</Label>
          <SearchableSelect
            options={disponivel.map((i) => ({
              value: String(i.id),
              label: `#${i.id} — ${i.tipo} ${i.brand || ''} ${i.model || ''}`,
              subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
            }))}
            value={selectedItemId}
            onValueChange={setSelectedItemId}
            placeholder="Selecione ou busque um equipamento disponível..."
            searchPlaceholder="Buscar por ID, tipo, marca, modelo, patrimônio..."
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
            <label className="flex items-center gap-2 text-sm text-muted-foreground cursor-pointer select-none mt-1">
              <input
                type="checkbox"
                checked={pessoaJuridica}
                onChange={(e) => setPessoaJuridica(e.target.checked)}
                className="rounded border-border"
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
        />
      )}

      {/* Lista de empréstimos pendentes de confirmação */}
      {pendentes.length > 0 && (
        <div className="space-y-3">
          <h3 className="font-medium text-foreground">Empréstimos Pendentes de Confirmação</h3>
          <DataTable data={pendentes} columns={pendingColumns} searchPlaceholder="Buscar..." />
        </div>
      )}
    </div>
  )
}
