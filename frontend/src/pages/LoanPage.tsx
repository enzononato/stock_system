import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listItems } from '@/api/items'
import { initiateLoan, generateLoanTerm, confirmLoan } from '@/api/loans'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { FileUpload } from '@/components/ui/FileUpload'
import { StatusBadge } from '@/components/ui/badge'
import { toast } from '@/components/ui/toast'
import { DataTable } from '@/components/ui/DataTable'
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { formatDate, maskCpfInput, isValidCpf } from '@/lib/utils'
import { FileDown, CheckCircle } from 'lucide-react'

const REVENDAS = ['Revalle Juazeiro','Revalle Bonfim','Revalle Petrolina','Revalle Ribeira','Revalle Paulo Afonso','Revalle Alagoinhas','Revalle Serrinha']
const CENTER_COSTS = ['101 - Puxada','202 - Armazém','301 - Administrativo','401 - Vendas','501 - Entrega','601 - CSC']
const SETORES = ['Contabilidade','Financeiro','Departamento pessoal','Cultura','Gente','Armazém','Distribuição','Segurança','Compras','TI','CME','Vendas','Puxada','Faturamento','Portaria','Auditório','Caixa','Diretoria']

export default function LoanPage() {
  const queryClient = useQueryClient()
  const [selectedItemId, setSelectedItemId] = useState('')
  const [usuario, setUsuario] = useState('')
  const [cpf, setCpf] = useState('')
  const [centerCost, setCenterCost] = useState('')
  const [setor, setSetor] = useState('')
  const [cargo, setCargo] = useState('')
  const [revenda, setRevenda] = useState('')
  const [dateIssue, setDateIssue] = useState(() => {
    const d = new Date()
    return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`
  })
  const [signedPdf, setSignedPdf] = useState<File | null>(null)
  const [pendingItemId, setPendingItemId] = useState<number | null>(null)

  const { data: items = [] } = useQuery({ queryKey: ['items'], queryFn: () => listItems() })

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

  const confirmMutation = useMutation({
    mutationFn: ({ itemId, pdf }: { itemId: number; pdf: File }) => confirmLoan(itemId, pdf),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setPendingItemId(null)
      setSignedPdf(null)
      toast('Empréstimo confirmado com sucesso!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao confirmar empréstimo.'
      toast(msg, 'error')
    },
  })

  async function handleGenerateTerm(itemId: number) {
    try {
      const blob = await generateLoanTerm(itemId)
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `termo_emprestimo_${itemId}.docx`
      a.click()
      URL.revokeObjectURL(url)
    } catch {
      toast('Erro ao gerar termo.', 'error')
    }
  }

  function handleLoanSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (!selectedItemId) { toast('Selecione um item.', 'error'); return }
    if (!isValidCpf(cpf)) { toast('CPF inválido. Deve conter 11 dígitos.', 'error'); return }
    loanMutation.mutate({ item_id: Number(selectedItemId), usuario, cpf, center_cost: centerCost, cargo, setor, revenda, date_issue: dateIssue })
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
          <Button size="sm" variant="outline" onClick={() => handleGenerateTerm(row.original.id)}>
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
        <p className="text-sm text-slate-500">Preencha os dados e inicie o processo de empréstimo.</p>
      </div>

      <form onSubmit={handleLoanSubmit} className="bg-white rounded-xl border p-6 space-y-4">
        <div className="flex flex-col gap-1.5">
          <Label>Equipamento *</Label>
          <Select value={selectedItemId} onValueChange={setSelectedItemId} required>
            <SelectTrigger><SelectValue placeholder="Selecione um item disponível" /></SelectTrigger>
            <SelectContent>
              {disponivel.map(i => (
                <SelectItem key={i.id} value={String(i.id)}>
                  #{i.id} — {i.tipo} {i.brand} {i.model}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>

        <div className="grid grid-cols-2 gap-4">
          <div className="flex flex-col gap-1.5">
            <Label>Funcionário *</Label>
            <Input value={usuario} onChange={e => setUsuario(e.target.value)} required />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>CPF *</Label>
            <Input value={cpf} onChange={e => setCpf(maskCpfInput(e.target.value))} placeholder="000.000.000-00" maxLength={14} required />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Cargo *</Label>
            <Input value={cargo} onChange={e => setCargo(e.target.value)} required />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Centro de Custo *</Label>
            <Select value={centerCost} onValueChange={setCenterCost} required>
              <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
              <SelectContent>{CENTER_COSTS.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent>
            </Select>
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Setor *</Label>
            <Select value={setor} onValueChange={setSetor} required>
              <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
              <SelectContent>{SETORES.map(s => <SelectItem key={s} value={s}>{s}</SelectItem>)}</SelectContent>
            </Select>
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Revenda *</Label>
            <Select value={revenda} onValueChange={setRevenda} required>
              <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
              <SelectContent>{REVENDAS.map(r => <SelectItem key={r} value={r}>{r}</SelectItem>)}</SelectContent>
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

      {/* Confirmação de empréstimo pendente */}
      {pendingItemId && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-6 space-y-4">
          <h3 className="font-medium text-amber-800">Confirmar Empréstimo — Item #{pendingItemId}</h3>
          <p className="text-sm text-amber-700">
            1. Clique em "Gerar Termo" para baixar o documento<br />
            2. Imprima, assine e escaneie como PDF<br />
            3. Faça o upload do PDF assinado abaixo e confirme
          </p>
          <Button variant="outline" onClick={() => handleGenerateTerm(pendingItemId)}>
            <FileDown size={14} className="mr-2" />Gerar Termo
          </Button>
          <FileUpload onFile={setSignedPdf} label="Upload do Termo Assinado (PDF)" />
          <Button
            disabled={!signedPdf || confirmMutation.isPending}
            onClick={() => signedPdf && confirmMutation.mutate({ itemId: pendingItemId, pdf: signedPdf })}
          >
            <CheckCircle size={14} className="mr-2" />
            {confirmMutation.isPending ? 'Confirmando...' : 'Confirmar Empréstimo'}
          </Button>
        </div>
      )}

      {/* Lista de empréstimos pendentes de confirmação */}
      {pendentes.length > 0 && (
        <div className="space-y-3">
          <h3 className="font-medium text-slate-700">Empréstimos Pendentes de Confirmação</h3>
          <DataTable data={pendentes} columns={pendingColumns} searchPlaceholder="Buscar..." />
        </div>
      )}
    </div>
  )
}
