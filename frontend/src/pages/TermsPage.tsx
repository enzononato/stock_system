import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listItems } from '@/api/items'
import { generateLoanTerm, confirmLoan } from '@/api/loans'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { FileUpload } from '@/components/ui/FileUpload'
import { StatusBadge } from '@/components/ui/badge'
import { toast } from '@/components/ui/toast'
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { formatDate } from '@/lib/utils'
import { FileDown, CheckCircle } from 'lucide-react'

export default function TermsPage() {
  const queryClient = useQueryClient()
  const [confirmingId, setConfirmingId] = useState<number | null>(null)
  const [signedPdf, setSignedPdf] = useState<File | null>(null)

  const { data: items = [] } = useQuery({ queryKey: ['items'], queryFn: () => listItems() })

  const pendentes = items.filter(i => i.status === 'Pendente')
  const ativos = items.filter(i => i.status === 'Indisponível')

  const confirmMutation = useMutation({
    mutationFn: ({ itemId, pdf }: { itemId: number; pdf: File }) => confirmLoan(itemId, pdf),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setConfirmingId(null)
      setSignedPdf(null)
      toast('Empréstimo confirmado com sucesso!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao confirmar.'
      toast(msg, 'error')
    },
  })

  async function handleDownloadTerm(itemId: number) {
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

  const pendingColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'model', header: 'Modelo' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'cpf', header: 'CPF' },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Data', cell: ({ getValue }) => formatDate(getValue() as string) },
    {
      id: 'actions',
      header: 'Ações',
      cell: ({ row }) => (
        <div className="flex gap-2">
          <Button size="sm" variant="outline" onClick={() => handleDownloadTerm(row.original.id)}>
            <FileDown size={13} className="mr-1" />Gerar Termo
          </Button>
          <Button size="sm" onClick={() => setConfirmingId(row.original.id)}>
            <CheckCircle size={13} className="mr-1" />Confirmar
          </Button>
        </div>
      ),
    },
  ]

  const activeColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'cpf', header: 'CPF' },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Data Empréstimo', cell: ({ getValue }) => formatDate(getValue() as string) },
    { accessorKey: 'status', header: 'Status', cell: ({ row }) => <StatusBadge status={row.original.status} /> },
  ]

  return (
    <div className="space-y-8 max-w-5xl">
      <div>
        <h2 className="text-xl font-semibold">Termos de Responsabilidade</h2>
        <p className="text-sm text-slate-500">Gerencie os termos de empréstimo pendentes e confirmados.</p>
      </div>

      {/* Pendentes de confirmação */}
      <div className="space-y-3">
        <div className="flex items-center gap-2">
          <h3 className="font-medium text-slate-700">Pendentes de Confirmação</h3>
          {pendentes.length > 0 && (
            <span className="text-xs bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full font-medium">
              {pendentes.length}
            </span>
          )}
        </div>
        <p className="text-sm text-slate-500">
          Gere o termo, imprima, colete a assinatura e confirme o empréstimo com o PDF assinado.
        </p>
        {pendentes.length === 0 ? (
          <p className="text-sm text-slate-400 py-6 text-center border rounded-xl bg-white">
            Nenhum empréstimo pendente de confirmação.
          </p>
        ) : (
          <DataTable data={pendentes} columns={pendingColumns} searchPlaceholder="Buscar..." />
        )}
      </div>

      {/* Painel de confirmação com upload */}
      {confirmingId && (
        <div className="bg-blue-50 border border-blue-200 rounded-xl p-5 space-y-4">
          <h3 className="font-medium text-blue-800">Confirmar Empréstimo — Item #{confirmingId}</h3>
          <p className="text-sm text-blue-700">Faça o upload do termo de responsabilidade assinado (PDF).</p>
          <FileUpload onFile={setSignedPdf} label="Arraste ou clique para enviar o PDF assinado" />
          <div className="flex gap-3">
            <Button
              disabled={!signedPdf || confirmMutation.isPending}
              onClick={() => signedPdf && confirmMutation.mutate({ itemId: confirmingId, pdf: signedPdf })}
            >
              <CheckCircle size={14} className="mr-2" />
              {confirmMutation.isPending ? 'Confirmando...' : 'Confirmar Empréstimo'}
            </Button>
            <Button variant="ghost" onClick={() => { setConfirmingId(null); setSignedPdf(null) }}>
              Cancelar
            </Button>
          </div>
        </div>
      )}

      {/* Empréstimos ativos (termos já confirmados) */}
      <div className="space-y-3">
        <h3 className="font-medium text-slate-700">Empréstimos Ativos ({ativos.length})</h3>
        {ativos.length === 0 ? (
          <p className="text-sm text-slate-400 py-6 text-center border rounded-xl bg-white">
            Nenhum empréstimo ativo no momento.
          </p>
        ) : (
          <DataTable data={ativos} columns={activeColumns} searchPlaceholder="Buscar por usuário, item..." />
        )}
      </div>
    </div>
  )
}
