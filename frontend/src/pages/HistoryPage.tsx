import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listHistory, reverseEntry, type HistoryEntry } from '@/api/history'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { toast } from '@/components/ui/toast'
import { useAuth } from '@/contexts/AuthContext'
import type { ColumnDef } from '@tanstack/react-table'
import { formatDate } from '@/lib/utils'
import { RotateCcw } from 'lucide-react'

const REVERSIBLE_OPS = ['Cadastro','Empréstimo','Confirmação Empréstimo','Devolução','Confirmação Devolução']

function OperationBadge({ op }: { op?: string }) {
  if (!op) return <Badge>-</Badge>
  if (op.includes('Empréstimo')) return <Badge variant="info">{op}</Badge>
  if (op.includes('Devolução')) return <Badge variant="success">{op}</Badge>
  if (op === 'Exclusão') return <Badge variant="danger">{op}</Badge>
  if (op === 'Estorno') return <Badge variant="warning">{op}</Badge>
  return <Badge>{op}</Badge>
}

export default function HistoryPage() {
  const { hasRole } = useAuth()
  const queryClient = useQueryClient()

  const { data: history = [], isLoading } = useQuery({
    queryKey: ['history'],
    queryFn: listHistory,
  })

  const reverseMutation = useMutation({
    mutationFn: (id: number) => reverseEntry(id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['history'] })
      queryClient.invalidateQueries({ queryKey: ['items'] })
      toast('Operação estornada com sucesso!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao estornar.'
      toast(msg, 'error')
    },
  })

  const columns: ColumnDef<HistoryEntry, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'item_id', header: 'Item', size: 60, cell: ({ getValue }) => getValue() as number ?? '-' },
    { accessorKey: 'operador', header: 'Operador' },
    { accessorKey: 'operation', header: 'Operação', cell: ({ row }) => <OperationBadge op={row.original.operation} /> },
    { accessorKey: 'tipo', header: 'Tipo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'marca', header: 'Marca', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'modelo', header: 'Modelo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'usuario', header: 'Usuário', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'revenda', header: 'Revenda', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'data_operacao', header: 'Data', cell: ({ getValue }) => formatDate(getValue() as string) },
    ...(hasRole('Gestor') ? [{
      id: 'actions',
      header: '',
      cell: ({ row }: { row: { original: HistoryEntry } }) =>
        REVERSIBLE_OPS.includes(row.original.operation ?? '') ? (
          <Button
            size="sm"
            variant="ghost"
            className="text-red-600 hover:text-red-700"
            onClick={() => reverseMutation.mutate(row.original.id)}
          >
            <RotateCcw size={14} className="mr-1" />Estornar
          </Button>
        ) : null,
    } as ColumnDef<HistoryEntry, unknown>] : []),
  ]

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-xl font-semibold">Histórico</h2>
        <p className="text-sm text-slate-500">Registro completo de todas as operações.</p>
      </div>
      {isLoading ? (
        <div className="py-8 text-center text-slate-400">Carregando...</div>
      ) : (
        <DataTable data={history} columns={columns} searchPlaceholder="Buscar por operador, item, usuário..." />
      )}
    </div>
  )
}
