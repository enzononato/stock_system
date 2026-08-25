import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { listItemsPaginated } from '@/api/items'
import { downloadSignedTerm } from '@/api/loans'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { toast } from '@/components/ui/toast'
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import type { ColumnDef } from '@tanstack/react-table'
import type { Item } from '@/api/items'
import { formatDate } from '@/lib/utils'
import { FileDown, CheckCircle } from 'lucide-react'

// Sem paginação nesta tela (filtra "Pendente"/"Indisponível" client-side a
// partir da lista completa) — usamos o teto de página do backend para não
// truncar em 50 itens (default de GET /api/items) como aconteceria chamando
// listItemsPaginated() sem limit.
const FETCH_ALL_LIMIT = 500

export default function TermsPage() {
  const [confirmingId, setConfirmingId] = useState<number | null>(null)

  const { data } = useQuery({
    queryKey: ['items'],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  })
  const items = data?.items ?? []

  const pendentes = items.filter(i => i.status === 'Pendente')
  const ativos = items.filter(i => i.status === 'Indisponível')

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
          <Button size="sm" variant="outline" onClick={() => generateAndDownloadLoanTerm(row.original.id)}>
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
    {
      id: 'actions',
      header: 'Termo',
      cell: ({ row }) => (
        <Button size="sm" variant="outline" onClick={async () => {
          try {
            const blob = await downloadSignedTerm(row.original.id)
            const url = URL.createObjectURL(blob)
            window.open(url, '_blank')
            setTimeout(() => URL.revokeObjectURL(url), 10000)
          } catch {
            toast('Termo assinado não encontrado.', 'error')
          }
        }}>
          <FileDown size={13} className="mr-1" />Ver Termo
        </Button>
      ),
    },
  ]

  return (
    <div className="space-y-8 max-w-6xl">
      <div>
        <h2 className="text-h1">Termos de Responsabilidade</h2>
        <p className="text-caption mt-1">Gerencie os termos de empréstimo pendentes e confirmados.</p>
      </div>

      {/* Pendentes de confirmação */}
      <div className="space-y-3">
        <div className="flex items-center gap-2">
          <h3 className="font-medium text-foreground">Pendentes de Confirmação</h3>
          {pendentes.length > 0 && (
            <span className="text-xs bg-warning/15 text-warning px-2 py-0.5 rounded-full font-medium">
              {pendentes.length}
            </span>
          )}
        </div>
        <p className="text-sm text-muted-foreground">
          Gere o termo, imprima, colete a assinatura e confirme o empréstimo com o PDF assinado.
        </p>
        {pendentes.length === 0 ? (
          <p className="text-sm text-muted-foreground py-10 text-center border border-border rounded-xl surface">
            Nenhum empréstimo pendente de confirmação.
          </p>
        ) : (
          <DataTable data={pendentes} columns={pendingColumns} searchPlaceholder="Buscar..." />
        )}
      </div>

      {/* Painel de confirmação com upload (T5: painel compartilhado com LoanPage) */}
      {confirmingId && (
        <ConfirmacaoTermo
          itemId={confirmingId}
          description="Faça o upload do termo de responsabilidade assinado (PDF)."
          variant="blue"
          uploadLabel="Arraste ou clique para enviar o PDF assinado"
          errorMessage="Erro ao confirmar."
          onConfirmed={() => setConfirmingId(null)}
          onCancel={() => setConfirmingId(null)}
        />
      )}

      {/* Empréstimos ativos (termos já confirmados) */}
      <div className="space-y-3">
        <h3 className="font-medium text-foreground">Empréstimos Ativos ({ativos.length})</h3>
        {ativos.length === 0 ? (
          <p className="text-sm text-muted-foreground py-10 text-center border border-border rounded-xl surface">
            Nenhum empréstimo ativo no momento.
          </p>
        ) : (
          <DataTable data={ativos} columns={activeColumns} searchPlaceholder="Buscar por usuário, item..." />
        )}
      </div>
    </div>
  )
}
