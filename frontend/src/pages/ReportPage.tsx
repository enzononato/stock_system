import { useMemo, useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { Download, RefreshCw } from 'lucide-react'

import { getMonthlyReport, exportMonthlyReportCsv } from '@/api/reports'
import { listUnidades } from '@/api/unidades'
import { useConstants } from '@/hooks/useConstants'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Badge } from '@/components/ui/badge'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import { formatDateTime } from '@/lib/utils'
import { getErrorMessage } from '@/lib/api-error'
import { toast } from '@/components/ui/toast'

const MONTHS = [
  'Janeiro',
  'Fevereiro',
  'Março',
  'Abril',
  'Maio',
  'Junho',
  'Julho',
  'Agosto',
  'Setembro',
  'Outubro',
  'Novembro',
  'Dezembro',
]

export default function ReportPage() {
  const currentDate = new Date()
  const [year, setYear] = useState(String(currentDate.getFullYear()))
  const [month, setMonth] = useState(String(currentDate.getMonth() + 1))
  const [queryParams, setQueryParams] = useState({
    year: currentDate.getFullYear(),
    month: currentDate.getMonth() + 1,
  })
  const [isExporting, setIsExporting] = useState(false)
  const [filterRevenda, setFilterRevenda] = useState('all')
  const [filterOp, setFilterOp] = useState('all')
  const [currentPage, setCurrentPage] = useState(0)
  const PAGE_SIZE = 7

  const { revendas = [] } = useConstants()
  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades-report'],
    queryFn: () => listUnidades(),
  })

  const {
    data: report = [],
    isLoading,
    refetch,
  } = useQuery({
    queryKey: ['report', queryParams.year, queryParams.month],
    queryFn: () => getMonthlyReport(queryParams.year, queryParams.month),
  })

  const revendaOptions = useMemo(() => {
    const set = new Set<string>()
    revendas.forEach((r) => r && set.add(r))
    unidades.forEach((u) => u.nome && set.add(u.nome))
    report.forEach((r) => r.revenda && set.add(r.revenda))
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))
  }, [revendas, unidades, report])

  const opTypes = useMemo(() => {
    const set = new Set<string>()
    report.forEach((r) => r.operation_type && set.add(r.operation_type))
    return Array.from(set).sort()
  }, [report])

  const filteredReport = useMemo(() => {
    return report.filter((r) => {
      if (filterRevenda !== 'all' && r.revenda !== filterRevenda) return false
      if (filterOp !== 'all' && r.operation_type !== filterOp) return false
      return true
    })
  }, [report, filterRevenda, filterOp])

  const paginatedReport = filteredReport.slice(
    currentPage * PAGE_SIZE,
    (currentPage + 1) * PAGE_SIZE
  )
  const totalPages = Math.ceil(filteredReport.length / PAGE_SIZE)

  function handleGenerate() {
    setQueryParams({ year: Number(year), month: Number(month) })
    setCurrentPage(0)
    void refetch()
  }

  async function handleExport() {
    setIsExporting(true)
    try {
      await exportMonthlyReportCsv(queryParams.year, queryParams.month)
    } catch (err) {
      toast.error(getErrorMessage(err, 'Erro ao exportar relatório.'))
    } finally {
      setIsExporting(false)
    }
  }

  // Aggregated stats
  const emprestimos = report.filter((r) => r.operation_type === 'Empréstimo').length
  const devolucoes = report.filter(
    (r) => r.operation_type === 'Devolução' || r.operation_type?.includes('Devol')
  ).length
  const cadastros = report.filter((r) => r.operation_type === 'Cadastro').length
  const uniqueUsers = new Set(report.map((r) => r.usuario).filter(Boolean)).size

  return (
    <div className="space-y-6 max-w-7xl mx-auto">
      {/* Barra Superior de Configuração */}
      <div className="rounded border border-border bg-surface p-5">
        <div className="flex flex-col lg:flex-row lg:items-end justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 mb-1">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
                Relatórios Corporativos
              </span>
            </div>
            <h1 className="text-heading-lg font-semibold tracking-tight text-foreground">
              Relatório Mensal de Operações
            </h1>
            <p className="text-body-sm text-muted-foreground mt-0.5">
              Consolidação de movimentações, ativos alocados e histórico do período.
            </p>
          </div>

          <div className="flex flex-wrap items-end gap-3">
            <div className="space-y-1">
              <Label className="text-caption font-medium">Ano</Label>
              <Input
                value={year}
                onChange={(e) => setYear(e.target.value)}
                className="w-20 h-8 font-mono text-body-sm"
                maxLength={4}
              />
            </div>

            <div className="space-y-1">
              <Label className="text-caption font-medium">Mês</Label>
              <Select value={month} onValueChange={setMonth}>
                <SelectTrigger className="w-36 h-8 text-body-sm">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {MONTHS.map((m, i) => (
                    <SelectItem key={i + 1} value={String(i + 1)}>
                      {m}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-1">
              <Label className="text-caption font-medium">Revenda / Filial</Label>
              <Select value={filterRevenda} onValueChange={setFilterRevenda}>
                <SelectTrigger className="w-44 h-8 text-body-sm">
                  <SelectValue placeholder="Todas" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">Todas as revendas</SelectItem>
                  {revendaOptions.map((r) => (
                    <SelectItem key={r} value={r}>
                      {r}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-1">
              <Label className="text-caption font-medium">Operação</Label>
              <Select value={filterOp} onValueChange={setFilterOp}>
                <SelectTrigger className="w-36 h-8 text-body-sm">
                  <SelectValue placeholder="Todas" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">Todas</SelectItem>
                  {opTypes.map((op) => (
                    <SelectItem key={op} value={op}>
                      {op}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <Button
              size="sm"
              onClick={handleGenerate}
              className="h-8"
            >
              <RefreshCw className="mr-1.5 size-3.5" />
              Gerar
            </Button>

            <Button
              size="sm"
              variant="outline"
              onClick={handleExport}
              disabled={isExporting || report.length === 0}
              className="h-8"
            >
              <Download className="mr-1.5 size-3.5" />
              {isExporting ? 'Exportando…' : 'CSV'}
            </Button>
          </div>
        </div>
      </div>

      {/* Estatísticas do Período */}
      {report.length > 0 && (
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
          {[
            { label: 'Total de Registros', value: report.length, symbol: '#' },
            { label: 'Empréstimos', value: emprestimos, symbol: '→' },
            { label: 'Devoluções', value: devolucoes, symbol: '←' },
            { label: 'Colaboradores Únicos', value: uniqueUsers, symbol: '👥' },
          ].map((stat) => (
            <div key={stat.label} className="rounded border border-border bg-surface p-4">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                {stat.label}
              </span>
              <div className="flex items-baseline gap-2 mt-1">
                <span className="font-mono text-heading-md font-bold text-foreground">
                  {stat.value}
                </span>
                <span className="font-mono text-body-lg text-muted-foreground">{stat.symbol}</span>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Preview de Dados Denso */}
      <div className="rounded border border-border bg-surface overflow-hidden">
        <div className="flex items-center justify-between px-5 py-3 border-b border-border">
          <div>
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Preview do Relatório
            </span>
            <p className="text-body-sm font-medium text-foreground mt-0.5">
              {MONTHS[queryParams.month - 1]} {queryParams.year} — {filteredReport.length} registro
              {filteredReport.length !== 1 ? 's' : ''}
              {filterRevenda !== 'all' && ` • ${filterRevenda}`}
              {filterOp !== 'all' && ` • ${filterOp}`}
            </p>
          </div>
          <Badge variant="default" className="font-mono text-[11px] border-border">
            {totalPages > 0 ? `Pág ${currentPage + 1}/${totalPages}` : 'VAZIO'}
          </Badge>
        </div>

        {isLoading ? (
          <div className="py-12 text-center text-caption text-muted-foreground">
            Gerando relatório…
          </div>
        ) : filteredReport.length === 0 ? (
          <div className="py-12 text-center text-caption text-muted-foreground">
            Nenhum registro encontrado para o período e filtros selecionados.
          </div>
        ) : (
          <>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-body-sm min-w-[900px]">
                <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground bg-surface-alt">
                  <tr>
                    <th className="py-2.5 px-3 font-semibold">ID</th>
                    <th className="py-2.5 px-3 font-semibold">Operador</th>
                    <th className="py-2.5 px-3 font-semibold">Operação</th>
                    <th className="py-2.5 px-3 font-semibold">Tipo</th>
                    <th className="py-2.5 px-3 font-semibold">Marca / Modelo</th>
                    <th className="py-2.5 px-3 font-semibold">Serial</th>
                    <th className="py-2.5 px-3 font-semibold">Nota Fiscal</th>
                    <th className="py-2.5 px-3 font-semibold">Colaborador</th>
                    <th className="py-2.5 px-3 font-semibold">Setor</th>
                    <th className="py-2.5 px-3 font-semibold">Revenda</th>
                    <th className="py-2.5 px-3 font-semibold">Data</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border">
                  {paginatedReport.map((row, idx) => (
                    <tr
                      key={row.history_id ?? `${idx}-${row.item_id}-${row.data_emprestimo}`}
                      className="hover:bg-surface-alt transition-colors"
                    >
                      <td className="py-2 px-3 font-mono text-caption text-muted-foreground">
                        {row.item_id ? `#${row.item_id}` : '—'}
                      </td>
                      <td className="py-2 px-3 font-medium">{row.operador || '—'}</td>
                      <td className="py-2 px-3">
                        <span className="font-mono text-caption text-foreground">
                          {row.operation_type || '—'}
                        </span>
                      </td>
                      <td className="py-2 px-3 text-muted-foreground">{row.tipo || '—'}</td>
                      <td className="py-2 px-3 text-muted-foreground">
                        {row.brand} {row.model}
                      </td>
                      <td className="py-2 px-3 font-mono text-caption text-foreground">
                        {row.identificador || '—'}
                      </td>
                      <td className="py-2 px-3 font-mono text-caption text-foreground">
                        {row.nota_fiscal || '—'}
                      </td>
                      <td className="py-2 px-3 font-medium">{row.usuario || '—'}</td>
                      <td className="py-2 px-3 text-muted-foreground">{row.setor || '—'}</td>
                      <td className="py-2 px-3 text-muted-foreground">{row.revenda || '—'}</td>
                      <td className="py-2 px-3 font-mono text-caption text-muted-foreground">
                        {formatDateTime(
                          row.data_emprestimo || row.data_confirmacao || row.data_devolucao
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {totalPages > 1 && (
              <div className="flex items-center justify-between px-5 py-3 border-t border-border">
                <Button
                  size="sm"
                  variant="outline"
                  disabled={currentPage === 0}
                  onClick={() => setCurrentPage((p) => p - 1)}
                  className="text-xs h-7 px-3"
                >
                  ← Anterior
                </Button>
                <span className="font-mono text-caption text-muted-foreground">
                  {currentPage + 1} de {totalPages} — {filteredReport.length} registros
                </span>
                <Button
                  size="sm"
                  variant="outline"
                  disabled={currentPage >= totalPages - 1}
                  onClick={() => setCurrentPage((p) => p + 1)}
                  className="text-xs h-7 px-3"
                >
                  Próxima →
                </Button>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  )
}
