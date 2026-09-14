import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { listItemsPaginated } from '@/api/items'
import { downloadSignedTerm } from '@/api/loans'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { toast } from '@/components/ui/toast'
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
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
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'model', header: 'Modelo' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'cpf', header: 'CPF', cell: ({ getValue }) => <span className="num">{getValue() as string}</span> },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Data', cell: ({ getValue }) => <span className="num">{formatDate(getValue() as string)}</span> },
    {
      id: 'actions',
      header: 'Ações',
      cell: ({ row }) => (
        <div className="flex gap-2">
          <Button size="sm" variant="outline" onClick={() => generateAndDownloadLoanTerm(row.original.id)}>
            <FileDown size={13} />Gerar Termo
          </Button>
          <Button size="sm" onClick={() => setConfirmingId(row.original.id)}>
            <CheckCircle size={13} />Confirmar
          </Button>
        </div>
      ),
    },
  ]

  const activeColumns: ColumnDef<Item, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'assigned_to', header: 'Usuário' },
    { accessorKey: 'cpf', header: 'CPF', cell: ({ getValue }) => <span className="num">{getValue() as string}</span> },
    { accessorKey: 'revenda', header: 'Revenda' },
    { accessorKey: 'date_issued', header: 'Data Empréstimo', cell: ({ getValue }) => <span className="num">{formatDate(getValue() as string)}</span> },
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
          <FileDown size={13} />Ver Termo
        </Button>
      ),
    },
  ]

  return (
    <div className="space-y-8">
      <PageHeader
        eyebrow="Conformidade Documental"
        eyebrowDetail="Assinaturas Pendentes"
        title="Termos de Responsabilidade"
        description="Gerencie os termos de empréstimo pendentes e confirmados."
      />

      {/* Pendentes de confirmação */}
      <div className="space-y-3">
        <PanelHeader
          title={<>Pendentes de Confirmação{pendentes.length > 0 && <> (<span className="num">{pendentes.length}</span>)</>}</>}
          description="Gere o termo, imprima, colete a assinatura e confirme o empréstimo com o PDF assinado."
        />
        {pendentes.length === 0 ? (
          <p className="text-body-sm text-muted-foreground py-10 text-center">
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
        <PanelHeader title={<>Empréstimos Ativos (<span className="num">{ativos.length}</span>)</>} />
        {ativos.length === 0 ? (
          <p className="text-body-sm text-muted-foreground py-10 text-center">
            Nenhum empréstimo ativo no momento.
          </p>
        ) : (
          <DataTable data={ativos} columns={activeColumns} searchPlaceholder="Buscar por usuário, item..." />
        )}
      </div>
    </div>
  )
}
