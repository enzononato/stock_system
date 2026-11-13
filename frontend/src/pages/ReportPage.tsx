import { useMemo, useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { Download, RefreshCw, Search } from 'lucide-react'

import { getMonthlyReport, exportMonthlyReportCsv } from '@/api/reports'
import { listUnidades } from '@/api/unidades'
import { useConstants } from '@/hooks/useConstants'
import { usePageSize } from '@/hooks/usePageSize'
import { PageHeader } from '@/components/layout/PageHeader'
import { Pagination } from '@/components/ui/Pagination'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { formatDateTime } from '@/lib/utils'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
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

// Sentinela de "sem filtro" para os <Select> — o Radix não aceita value="".
const ALL = 'all'

// O endpoint /reports/monthly devolve o mês inteiro numa única resposta, sem
// limit/offset — a paginação abaixo é feita no cliente, sobre o array já
// carregado (tamanho 25/50/100 por `usePageSize('relatorio')`). Os filtros de
// unidade/operação e a busca também são locais.

export default function ReportPage() {
  const currentDate = new Date()
  const [year, setYear] = useState(String(currentDate.getFullYear()))
  const [month, setMonth] = useState(String(currentDate.getMonth() + 1))
  const [queryParams, setQueryParams] = useState({
    year: currentDate.getFullYear(),
    month: currentDate.getMonth() + 1,
  })
  const [isExporting, setIsExporting] = useState(false)
  const [filterRevenda, setFilterRevenda] = useState(ALL)
  const [filterOp, setFilterOp] = useState(ALL)
  const [search, setSearch] = useState('')
  const [currentPage, setCurrentPage] = useState(1)
  const [pageSize, setPageSize] = usePageSize('relatorio')

  // Revendas oficiais (constantes) e unidades cadastradas enriquecem as opções
  // do filtro, para que unidades sem movimentação no mês também apareçam.
  const { revendas = [] } = useConstants()
  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades-report'],
    queryFn: () => listUnidades(),
  })

  const {
    data: report = [],
    isLoading,
    isFetching,
    refetch,
  } = useQuery({
    queryKey: ['report', queryParams.year, queryParams.month],
    queryFn: () => getMonthlyReport(queryParams.year, queryParams.month),
  })

  const revendaOptions = useMemo(() => {
    const set = new Set<string>()
    revendas.forEach((r) => { if (r) set.add(r) })
    unidades.forEach((u) => { if (u.nome) set.add(u.nome) })
    report.forEach((r) => { if (r.revenda) set.add(r.revenda) })
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))
  }, [revendas, unidades, report])

  const opTypes = useMemo(() => {
    const set = new Set<string>()
    report.forEach((r) => { if (r.operation_type) set.add(r.operation_type) })
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))
  }, [report])

  const filteredReport = useMemo(() => {
    const term = search.trim().toLowerCase()
    return report.filter((r) => {
      if (filterRevenda !== ALL && r.revenda !== filterRevenda) return false
      if (filterOp !== ALL && r.operation_type !== filterOp) return false
      if (term && !Object.values(r).some((v) => v != null && String(v).toLowerCase().includes(term))) return false
      return true
    })
  }, [report, filterRevenda, filterOp, search])

  const lastPage = Math.max(1, Math.ceil(filteredReport.length / pageSize))
  // Nunca aponta para uma página que deixou de existir após filtrar.
  const safePage = Math.min(currentPage, lastPage)
  const paginatedReport = filteredReport.slice((safePage - 1) * pageSize, safePage * pageSize)

  const hasActiveFilters = filterRevenda !== ALL || filterOp !== ALL || search.trim() !== ''

  // Toda mudança de filtro/busca volta para a primeira página.
  function changeRevenda(v: string) {
    setFilterRevenda(v)
    setCurrentPage(1)
  }
  function changeOp(v: string) {
    setFilterOp(v)
    setCurrentPage(1)
  }
  function changeSearch(v: string) {
    setSearch(v)
    setCurrentPage(1)
  }
  function handleClearFilters() {
    setFilterRevenda(ALL)
    setFilterOp(ALL)
    setSearch('')
    setCurrentPage(1)
  }

  function handleGenerate() {
    setQueryParams({ year: Number(year), month: Number(month) })
    // Um mês novo pode não ter as mesmas unidades/operações do anterior —
    // zera os filtros para não esconder o mês novo atrás de um filtro velho.
    handleClearFilters()
    void refetch()
  }

  async function handleExport() {
    setIsExporting(true)
    try {
      await exportMonthlyReportCsv(queryParams.year, queryParams.month)
    } catch (err) {
      toast(getErrorMessage(err, 'Erro ao exportar relatório.'), 'error')
    } finally {
      setIsExporting(false)
    }
  }

  // Totais sempre sobre o mês inteiro (`report`), não sobre o recorte filtrado.
  //
  // "Empréstimos" conta por operation_type === 'Empréstimo': o UNION do backend
  // reaproveita data_emprestimo para toda operação, então só operation_type
  // distingue um empréstimo de fato.
  //
  // Devoluções: o UNION nunca gera linha com operation_type 'Devolução' — a
  // devolução vem só como data_devolucao embutida na linha 'Empréstimo'. Por
  // isso a métrica é "empréstimos deste mês que já foram devolvidos" (linhas
  // Empréstimo com data_devolucao), e o rótulo diz exatamente isso. A FT_STC
  // contava operation_type 'Devolução', que nunca bate — o contador ficava
  // sempre zerado.
  const emprestimos = report.filter((r) => r.operation_type === 'Empréstimo').length
  const devolvidos = report.filter((r) => r.operation_type === 'Empréstimo' && r.data_devolucao).length
  const uniqueUsers = new Set(report.map((r) => r.usuario).filter(Boolean)).size

  const periodo = `${MONTHS[queryParams.month - 1]}/${queryParams.year}`

  return (
    <div className="space-y-6 max-w-7xl mx-auto">
      {/* Barra superior de configuração */}
      <PageHeader
        eyebrow="Relatórios Corporativos"
        title="Relatório Mensal de Operações"
        description="Consolidação de movimentações, ativos alocados e histórico do período."
      />

      <div className="rounded-md border border-border bg-surface p-5 space-y-3">
        <div>
          <div className="grid grid-cols-1 gap-3 sm:flex sm:flex-wrap sm:items-end">
            <div className="space-y-1">
              <Label className="text-caption font-medium">Ano</Label>
              <Input
                value={year}
                onChange={(e) => setYear(e.target.value)}
                className="w-full sm:w-20 h-8 num text-body-sm"
                maxLength={4}
              />
            </div>

            <div className="space-y-1">
              <Label className="text-caption font-medium">Mês</Label>
              <Select value={month} onValueChange={setMonth}>
                <SelectTrigger aria-label="Mês" className="w-full sm:w-44 h-8 text-body-sm">
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
              <Select value={filterRevenda} onValueChange={changeRevenda} disabled={report.length === 0}>
                <SelectTrigger aria-label="Revenda" className="w-full sm:w-44 h-8 text-body-sm">
                  <SelectValue placeholder="Todas" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value={ALL}>Todas as revendas</SelectItem>
                  {revendaOptions.map((r) => (
                    <SelectItem key={r} value={r}>
                      {upper(r)}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-1">
              <Label className="text-caption font-medium">Operação</Label>
              <Select value={filterOp} onValueChange={changeOp} disabled={report.length === 0}>
                <SelectTrigger aria-label="Operação" className="w-full sm:w-44 h-8 text-body-sm">
                  <SelectValue placeholder="Todas" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value={ALL}>Todas</SelectItem>
                  {opTypes.map((op) => (
                    <SelectItem key={op} value={op}>
                      {op}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <Button size="sm" onClick={handleGenerate} loading={isFetching}>
              {!isFetching && <RefreshCw />}
              Gerar
            </Button>

            <Button
              size="sm"
              variant="outline"
              onClick={handleExport}
              loading={isExporting}
              disabled={report.length === 0}
            >
              {!isExporting && <Download />}
              CSV
            </Button>
          </div>
        </div>

        {/* Aviso de escopo do CSV: a exportação só recebe ano/mês — os filtros
            locais nunca chegam ao servidor. Avisamos em vez de bloquear. */}
        <p className="text-caption text-muted-foreground">
          {`O CSV exportado sempre traz ${periodo} inteiro`}
          {hasActiveFilters
            ? ' — os filtros de revenda, operação e busca não são aplicados ao arquivo.'
            : ', sem os filtros de revenda e operação.'}
        </p>

        {/* O ramo de periféricos do relatório não carrega revenda: filtrar por
            unidade esconde essas operações. Não dá para corrigir sem mexer no
            backend — só sinalizar. */}
        {filterRevenda !== ALL && (
          <p className="rounded-md bg-brand-soft p-3 text-body-sm text-brand-soft-foreground">
            Operações de periférico (Cadastro Periférico, Vínculo, Desvínculo, Substituição Periférico) não têm
            unidade registrada nesta consulta e por isso somem da lista ao filtrar por revenda — isso não significa
            que não houve movimentação de periféricos no mês.
          </p>
        )}
      </div>

      {/* Estatísticas do período */}
      {!isLoading && report.length > 0 && (
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
          {[
            { label: 'Total de Registros', value: report.length, symbol: '#' },
            { label: 'Empréstimos', value: emprestimos, symbol: '→' },
            { label: 'Empréstimos já devolvidos', value: devolvidos, symbol: '←', hint: 'dos emitidos no período' },
            { label: 'Colaboradores Únicos', value: uniqueUsers, symbol: '◯' },
          ].map((stat) => (
            <div key={stat.label} className="rounded-md bg-brand-soft p-4">
              <span className="text-caption font-semibold uppercase tracking-wider text-brand-soft-foreground block">
                {stat.label}
              </span>
              <div className="flex items-baseline gap-2 mt-1">
                <span className="num text-heading font-bold text-brand-soft-foreground">{stat.value}</span>
                <span className="num text-body-lg text-brand-soft-foreground">{stat.symbol}</span>
              </div>
              {stat.hint && <span className="text-body-sm text-brand-soft-foreground block mt-0.5">{stat.hint}</span>}
            </div>
          ))}
        </div>
      )}

      {/* Preview de dados denso */}
      <div className="rounded-md border border-border bg-surface overflow-hidden">
        <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-3 border-b border-border">
          <div>
            <span className="text-caption font-semibold uppercase tracking-wider text-brand-text">
              Preview do Relatório
            </span>
            <p className="text-body-sm font-medium text-foreground mt-0.5">
              {MONTHS[queryParams.month - 1]} {queryParams.year} — {filteredReport.length} registro
              {filteredReport.length !== 1 ? 's' : ''}
              {filterRevenda !== ALL && ` • ${upper(filterRevenda)}`}
              {filterOp !== ALL && ` • ${filterOp}`}
            </p>
          </div>
          <div className="flex items-center gap-2">
            <div className="relative">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
              <Input
                value={search}
                onChange={(e) => changeSearch(e.target.value)}
                placeholder="Buscar no relatório…"
                aria-label="Buscar no relatório"
                className="pl-8 h-8 w-full sm:w-56 text-body-sm"
                disabled={report.length === 0}
              />
            </div>
          </div>
        </div>

        {isLoading ? (
          <div className="py-12 text-center text-caption text-muted-foreground">Gerando relatório…</div>
        ) : report.length === 0 ? (
          <div className="py-12 text-center text-caption text-muted-foreground">
            {`Nenhuma movimentação registrada em ${periodo}.`}
          </div>
        ) : filteredReport.length === 0 ? (
          <div className="py-12 flex flex-col items-center gap-3 text-center">
            <span className="text-caption text-muted-foreground">
              Nenhum registro encontrado para o período e filtros selecionados.
            </span>
            <Button size="sm" variant="outline" onClick={handleClearFilters} className="text-xs h-7 px-3">
              Limpar filtros
            </Button>
          </div>
        ) : (
          <>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-body-sm min-w-[900px]">
                <thead className="bg-brand-soft text-brand-soft-foreground text-caption uppercase tracking-wider">
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
                      key={row.history_id ?? `${safePage}-${idx}-${row.item_id}-${row.data_emprestimo}`}
                      className="hover:bg-surface-alt transition-colors"
                    >
                      <td className="py-2 px-3 num text-caption text-muted-foreground">
                        {row.item_id ? `#${row.item_id}` : '—'}
                      </td>
                      <td className="py-2 px-3 font-medium">{row.operador || '—'}</td>
                      <td className="py-2 px-3">
                        <span className="num text-caption text-foreground">{row.operation_type || '—'}</span>
                      </td>
                      <td className="py-2 px-3 text-muted-foreground">{row.tipo || '—'}</td>
                      <td className="py-2 px-3 text-muted-foreground">
                        {[row.brand, row.model].filter(Boolean).join(' ') || '—'}
                      </td>
                      <td className="py-2 px-3 num text-caption text-foreground">{row.identificador || '—'}</td>
                      <td className="py-2 px-3 num text-caption text-foreground">{row.nota_fiscal || '—'}</td>
                      <td className="py-2 px-3 font-medium">{row.usuario || '—'}</td>
                      <td className="py-2 px-3 text-muted-foreground">{row.setor || '—'}</td>
                      <td className="py-2 px-3 text-muted-foreground">{upper(row.revenda) || '—'}</td>
                      <td className="py-2 px-3 num text-caption text-muted-foreground">
                        {formatDateTime(row.data_emprestimo || row.data_confirmacao || row.data_devolucao)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="px-5 py-3 border-t border-border">
              <Pagination
                page={safePage}
                pageSize={pageSize}
                total={filteredReport.length}
                onPageChange={setCurrentPage}
                onPageSizeChange={(n) => {
                  setPageSize(n)
                  setCurrentPage(1)
                }}
                isFetching={isFetching}
              />
            </div>
          </>
        )}
      </div>
    </div>
  )
}
