import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { getMonthlyReport, exportMonthlyReportCsv, type ReportRow } from '@/api/reports'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { toast } from '@/components/ui/toast'
import { formatDateTime } from '@/lib/utils'
import type { ColumnDef } from '@tanstack/react-table'
import { Download } from 'lucide-react'

const MONTHS = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

export default function ReportPage() {
  const currentDate = new Date()
  const [year, setYear] = useState(String(currentDate.getFullYear()))
  const [month, setMonth] = useState(String(currentDate.getMonth() + 1))
  const [queryParams, setQueryParams] = useState({ year: currentDate.getFullYear(), month: currentDate.getMonth() + 1 })
  const [isExporting, setIsExporting] = useState(false)
  const [filterRevenda, setFilterRevenda] = useState('all')

  const { data: report = [], isLoading, refetch } = useQuery({
    queryKey: ['report', queryParams.year, queryParams.month],
    queryFn: () => getMonthlyReport(queryParams.year, queryParams.month),
  })

  // Lista de revendas presentes no relatório do mês, para popular o filtro
  // sem precisar de outro endpoint.
  const revendaOptions = [...new Set(report.map((r) => r.revenda).filter(Boolean))] as string[]
  const filteredReport = filterRevenda === 'all' ? report : report.filter((r) => r.revenda === filterRevenda)

  function handleGenerate() {
    setQueryParams({ year: Number(year), month: Number(month) })
    refetch()
  }

  async function handleExport() {
    setIsExporting(true)
    try {
      await exportMonthlyReportCsv(queryParams.year, queryParams.month)
    } catch (err) {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao exportar relatório.'
      toast(msg, 'error')
    } finally {
      setIsExporting(false)
    }
  }

  const columns: ColumnDef<ReportRow, unknown>[] = [
    { accessorKey: 'item_id', header: 'ID Item', size: 70 },
    { accessorKey: 'operador', header: 'Operador' },
    { accessorKey: 'operation_type', header: 'Operação' },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'model', header: 'Modelo' },
    { accessorKey: 'identificador', header: 'Identificador', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'nota_fiscal', header: 'Nota Fiscal', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'fornecedor', header: 'Fornecedor', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'usuario', header: 'Usuário', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'cpf', header: 'CPF', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'cargo', header: 'Cargo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'setor', header: 'Setor', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'revenda', header: 'Revenda', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'center_cost', header: 'C. Custo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'data_emprestimo', header: 'Data', cell: ({ getValue }) => formatDateTime(getValue() as string) },
    { accessorKey: 'data_confirmacao', header: 'Confirmação', cell: ({ getValue }) => formatDateTime(getValue() as string) },
    { accessorKey: 'data_devolucao', header: 'Devolução', cell: ({ getValue }) => formatDateTime(getValue() as string) },
    { accessorKey: 'details', header: 'Detalhes', cell: ({ getValue }) => getValue() as string || '-' },
  ]

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-h1">Relatório Mensal</h2>
        <p className="text-caption mt-1">Visualize todas as operações de um determinado mês.</p>
      </div>

      <div className="flex items-end gap-4 bg-card rounded-xl border p-4">
        <div className="flex flex-col gap-1.5">
          <Label>Ano</Label>
          <Input value={year} onChange={e => setYear(e.target.value)} className="w-24" />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label>Mês</Label>
          <Select value={month} onValueChange={setMonth}>
            <SelectTrigger className="w-40"><SelectValue /></SelectTrigger>
            <SelectContent>
              {MONTHS.map((m, i) => <SelectItem key={i+1} value={String(i+1)}>{m}</SelectItem>)}
            </SelectContent>
          </Select>
        </div>
        <Button onClick={handleGenerate}>Gerar Relatório</Button>
        <Button variant="outline" onClick={handleExport} disabled={isExporting}>
          <Download size={14} className="mr-2" />
          {isExporting ? 'Exportando...' : 'Exportar CSV'}
        </Button>
        <div className="flex flex-col gap-1.5 ml-auto">
          <Label>Revenda</Label>
          <Select value={filterRevenda} onValueChange={setFilterRevenda}>
            <SelectTrigger className="w-48"><SelectValue placeholder="Todas as revendas" /></SelectTrigger>
            <SelectContent>
              <SelectItem value="all">Todas as revendas</SelectItem>
              {revendaOptions.map((r) => <SelectItem key={r} value={r}>{r}</SelectItem>)}
            </SelectContent>
          </Select>
        </div>
      </div>

      {isLoading ? (
        <div className="py-8 text-center text-muted-foreground">Gerando relatório...</div>
      ) : (
        <DataTable data={filteredReport} columns={columns} searchPlaceholder="Buscar no relatório..." />
      )}
    </div>
  )
}
