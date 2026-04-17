import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listItems } from '@/api/items'
import { initiateReturn, confirmReturn } from '@/api/loans'
import { Button } from '@/components/ui/button'
import { FileUpload } from '@/components/ui/FileUpload'
import { DataTable } from '@/components/ui/DataTable'
import { toast } from '@/components/ui/toast'
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { formatDate } from '@/lib/utils'
import { FileDown, CheckCircle } from 'lucide-react'

export default function ReturnPage() {
  const queryClient = useQueryClient()
  const [pendingReturnId, setPendingReturnId] = useState<number | null>(null)
  const [signedPdf, setSignedPdf] = useState<File | null>(null)

  const { data: items = [] } = useQuery({ queryKey: ['items'], queryFn: () => listItems() })
  const indisponivel = items.filter(i => i.status === 'Indisponível')
  const pendenteDevolucao = items.filter(i => i.status === 'Pendente Devolução')

  const initiateMutation = useMutation({
    mutationFn: (itemId: number) => initiateReturn(itemId),
    onSuccess: (data, itemId) => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setPendingReturnId(itemId)
      // Download automático do termo
      window.open(data.download_url, '_blank')
      toast('Termo de devolução gerado! Faça a assinatura e confirme.')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao gerar termo.'
      toast(msg, 'error')
    },
  })

  const confirmMutation = useMutation({
    mutationFn: ({ itemId, pdf }: { itemId: number; pdf: File }) => confirmReturn(itemId, pdf),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setPendingReturnId(null)
      setSignedPdf(null)
      toast('Devolução confirmada com sucesso!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao confirmar devolução.'
      toast(msg, 'error')
    },
  })

  const activeColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'cpf', header: 'CPF' },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Empréstimo', cell: ({ getValue }) => formatDate(getValue() as string) },
    {
      id: 'actions',
      header: '',
      cell: ({ row }) => (
        <Button size="sm" variant="outline" onClick={() => initiateMutation.mutate(row.original.id)}>
          <FileDown size={14} className="mr-1" />Gerar Termo
        </Button>
      ),
    },
  ]

  const pendingColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'revenda', header: 'Revenda' },
    {
      id: 'actions',
      header: '',
      cell: ({ row }) => (
        <Button size="sm" variant="ghost" onClick={() => setPendingReturnId(row.original.id)}>
          Confirmar
        </Button>
      ),
    },
  ]

  return (
    <div className="space-y-8 max-w-4xl">
      <div>
        <h2 className="text-xl font-semibold">Devolver Equipamento</h2>
        <p className="text-sm text-slate-500">Gerencie as devoluções de equipamentos emprestados.</p>
      </div>

      {/* Empréstimos ativos */}
      <div className="space-y-3">
        <h3 className="font-medium text-slate-700">Empréstimos Ativos ({indisponivel.length})</h3>
        <p className="text-sm text-slate-500">Selecione um item para gerar o termo de devolução.</p>
        <DataTable data={indisponivel} columns={activeColumns} searchPlaceholder="Buscar por usuário, item..." />
      </div>

      {/* Confirmação de devolução */}
      {pendingReturnId && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-6 space-y-4">
          <h3 className="font-medium text-amber-800">Confirmar Devolução — Item #{pendingReturnId}</h3>
          <p className="text-sm text-amber-700">
            O termo de devolução foi gerado. Faça o upload do PDF assinado para confirmar.
          </p>
          <FileUpload onFile={setSignedPdf} label="Upload do Termo de Devolução Assinado (PDF)" />
          <div className="flex gap-3">
            <Button
              disabled={!signedPdf || confirmMutation.isPending}
              onClick={() => signedPdf && confirmMutation.mutate({ itemId: pendingReturnId, pdf: signedPdf })}
            >
              <CheckCircle size={14} className="mr-2" />
              {confirmMutation.isPending ? 'Confirmando...' : 'Confirmar Devolução'}
            </Button>
            <Button variant="ghost" onClick={() => setPendingReturnId(null)}>Cancelar</Button>
          </div>
        </div>
      )}

      {/* Devoluções pendentes de confirmação */}
      {pendenteDevolucao.length > 0 && (
        <div className="space-y-3">
          <h3 className="font-medium text-slate-700">Pendente de Confirmação ({pendenteDevolucao.length})</h3>
          <DataTable data={pendenteDevolucao} columns={pendingColumns} searchPlaceholder="Buscar..." />
        </div>
      )}
    </div>
  )
}
