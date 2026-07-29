import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { getMonthlyReport, getMonthlyReportExportUrl, type ReportRow } from '@/api/reports'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { formatDateTime } from '@/lib/utils'
import type { ColumnDef } from '@tanstack/react-table'
import { Download } from 'lucide-react'

const MONTHS = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

export default function ReportPage() {
  const currentDate = new Date()
  const [year, setYear] = useState(String(currentDate.getFullYear()))
  const [month, setMonth] = useState(String(currentDate.getMonth() + 1))
  const [queryParams, setQueryParams] = useState({ year: currentDate.getFullYear(), month: currentDate.getMonth() + 1 })

  const { data: report = [], isLoading, refetch } = useQuery({
    queryKey: ['report', queryParams.year, queryParams.month],
    queryFn: () => getMonthlyReport(queryParams.year, queryParams.month),
  })

  function handleGenerate() {
    setQueryParams({ year: Number(year), month: Number(month) })
    refetch()
  }

  const columns: ColumnDef<ReportRow, unknown>[] = [
    { accessorKey: 'item_id', header: 'ID Item', size: 70 },
    { accessorKey: 'operador', header: 'Operador' },
    { accessorKey: 'operation_type', header: 'Operação' },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'model', header: 'Modelo' },
    { accessorKey: 'usuario', header: 'Usuário', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'cpf', header: 'CPF', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'revenda', header: 'Revenda', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'center_cost', header: 'C. Custo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'data_emprestimo', header: 'Data', cell: ({ getValue }) => formatDateTime(getValue() as string) },
    { accessorKey: 'data_devolucao', header: 'Devolução', cell: ({ getValue }) => formatDateTime(getValue() as string) },
  ]

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-semibold">Relatório Mensal</h2>
        <p className="text-sm text-slate-500">Visualize todas as operações de um determinado mês.</p>
      </div>

      <div className="flex items-end gap-4 bg-white rounded-xl border p-4">
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
        <a
          href={getMonthlyReportExportUrl(queryParams.year, queryParams.month)}
          download
        >
          <Button variant="outline">
            <Download size={14} className="mr-2" />Exportar CSV
          </Button>
        </a>
      </div>

      {isLoading ? (
        <div className="py-8 text-center text-slate-400">Gerando relatório...</div>
      ) : (
        <DataTable data={report} columns={columns} searchPlaceholder="Buscar no relatório..." />
      )}
    </div>
  )
}
