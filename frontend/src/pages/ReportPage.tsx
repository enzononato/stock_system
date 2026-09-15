import { useEffect, useMemo, useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { getMonthlyReport, exportMonthlyReportCsv, type ReportRow } from '@/api/reports'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { StatBlock } from '@/components/ui/StatBlock'
import { toast } from '@/components/ui/toast'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { formatDateTime } from '@/lib/utils'
import type { ColumnDef } from '@tanstack/react-table'
import { Download, Loader2 } from 'lucide-react'

const MONTHS = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

// Sentinela de "sem filtro" para os <Select> — o Radix não aceita value="".
const ALL = 'all'

// O endpoint /reports/monthly devolve o mês inteiro numa única resposta, sem
// limit/offset (ver backend/app/routers/reports.py) — por isso a paginação
// abaixo é feita inteiramente no cliente, sobre o array já carregado.
const PAGE_SIZE = 15

export default function ReportPage() {
  const currentDate = new Date()
  const [year, setYear] = useState(String(currentDate.getFullYear()))
  const [month, setMonth] = useState(String(currentDate.getMonth() + 1))
  const [queryParams, setQueryParams] = useState({ year: currentDate.getFullYear(), month: currentDate.getMonth() + 1 })
  const [isExporting, setIsExporting] = useState(false)

  // Filtros de Unidade e Tipo de Operação (Step 1): aplicados no cliente
  // sobre as linhas do mês já buscadas — não disparam nova requisição.
  const [filterUnidade, setFilterUnidade] = useState(ALL)
  const [filterOperationType, setFilterOperationType] = useState(ALL)
  const [search, setSearch] = useState('')
  const [pageIndex, setPageIndex] = useState(0)

  const { data: report = [], isLoading, refetch } = useQuery({
    queryKey: ['report', queryParams.year, queryParams.month],
    queryFn: () => getMonthlyReport(queryParams.year, queryParams.month),
  })

  function handleGenerate() {
    setQueryParams({ year: Number(year), month: Number(month) })
    // Um mês novo pode não ter as mesmas unidades/tipos/termo buscado do
    // anterior — zera os filtros para não esconder o mês novo atrás de um
    // filtro que não faz mais sentido para ele.
    setFilterUnidade(ALL)
    setFilterOperationType(ALL)
    setSearch('')
    setPageIndex(0)
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

  // Opções dos filtros derivadas dos próprios dados do mês (não de uma lista
  // fixa) — sempre a partir de `report` (não do já-filtrado), para trocar de
  // unidade ou tipo sem que as demais opções desapareçam da lista.
  const unidadeOptions = useMemo(() => {
    const set = new Set<string>()
    report.forEach((r) => { if (r.revenda) set.add(r.revenda) })
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))
  }, [report])

  const operationTypeOptions = useMemo(() => {
    const set = new Set<string>()
    report.forEach((r) => { if (r.operation_type) set.add(r.operation_type) })
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))
  }, [report])

  const filteredReport = useMemo(() => {
    return report.filter((r) => {
      if (filterUnidade !== ALL && r.revenda !== filterUnidade) return false
      if (filterOperationType !== ALL && r.operation_type !== filterOperationType) return false
      return true
    })
  }, [report, filterUnidade, filterOperationType])

  // Busca textual própria: como a paginação é no cliente (ver PAGE_SIZE
  // acima), quem filtra por texto precisa ser este componente — não o
  // DataTable — para o total de `pagination` continuar batendo com o que é
  // efetivamente mostrado.
  const searchedReport = useMemo(() => {
    const term = search.trim().toLowerCase()
    if (!term) return filteredReport
    return filteredReport.filter((r) =>
      Object.values(r).some((v) => v != null && String(v).toLowerCase().includes(term))
    )
  }, [filteredReport, search])

  // Volta para a primeira página sempre que um filtro ou a busca muda o
  // conjunto de linhas — senão a página em que o usuário estava pode ficar
  // além do novo total.
  useEffect(() => {
    setPageIndex(0)
  }, [filterUnidade, filterOperationType, search])

  // Se o total encolher (troca de filtro, nova busca) a ponto da página atual
  // deixar de existir, volta para a última página válida.
  useEffect(() => {
    if (searchedReport.length === 0) return
    const maxPageIndex = Math.max(0, Math.ceil(searchedReport.length / PAGE_SIZE) - 1)
    if (pageIndex > maxPageIndex) setPageIndex(maxPageIndex)
  }, [searchedReport.length, pageIndex])

  const pageRows = useMemo(
    () => searchedReport.slice(pageIndex * PAGE_SIZE, (pageIndex + 1) * PAGE_SIZE),
    [searchedReport, pageIndex]
  )

  const hasActiveFilters = filterUnidade !== ALL || filterOperationType !== ALL || search.trim() !== ''

  function handleClearFilters() {
    setFilterUnidade(ALL)
    setFilterOperationType(ALL)
    setSearch('')
    setPageIndex(0)
  }

  // Blocos de totais (Step 2): sempre sobre o mês inteiro buscado (`report`),
  // não sobre o resultado filtrado — respondem "o que aconteceu no mês"; os
  // filtros abaixo só recortam a tabela de detalhe. Mesmo critério da
  // referência (origin/redesign-frontend, linhas ~101-104).
  //
  // "Empréstimos" conta por operation_type === 'Empréstimo'. O UNION do
  // backend (generate_monthly_report, backend/app/db/inventory_manager_db.py
  // ~L960) reaproveita a coluna data_emprestimo para TODA operação do
  // relatório — cadastro, periférico, exclusão também preenchem essa coluna.
  // Contar por ela ser truthy infla o total; operation_type é o único campo
  // que distingue de fato um empréstimo das demais operações.
  //
  // "Empréstimos já devolvidos" (rótulo deliberadamente não é "Devoluções") não
  // pode ser contado por operation_type: o UNION nunca gera uma linha com
  // operation_type='Devolução' — a devolução aparece só como o campo
  // data_devolucao embutido na própria linha 'Empréstimo' (subquery MIN(...)
  // dentro da função). Por isso o critério aqui é "empréstimos deste mês que
  // já foram devolvidos": linhas Empréstimo com data_devolucao preenchida. A
  // referência tenta `operation_type === 'Devolução' ||
  // operation_type?.includes('Devol')`, que nunca bate com nenhuma linha real
  // — o contador dela fica sempre zerado; corrigido aqui.
  //
  // O rótulo evita a palavra "Devoluções" de propósito: essa métrica NÃO é
  // "devoluções ocorridas neste mês" — é "dos empréstimos emitidos neste mês,
  // quantos já foram devolvidos (em qualquer data)". As duas grandezas
  // divergem nos dois sentidos: uma devolução em março de um empréstimo de
  // fevereiro não entra aqui, e um empréstimo de março devolvido em maio
  // entra. Rotular como "Devoluções" prometeria algo que o número não
  // entrega, sem o usuário ter como perceber — o mesmo tipo de armadilha
  // sinalizada nos avisos do CSV e do filtro de unidade.
  const totalRegistros = report.length
  const totalEmprestimos = report.filter((r) => r.operation_type === 'Empréstimo').length
  const totalDevolucoes = report.filter((r) => r.operation_type === 'Empréstimo' && r.data_devolucao).length
  const totalColaboradores = new Set(report.map((r) => r.usuario).filter(Boolean)).size

  const columns: ColumnDef<ReportRow, unknown>[] = [
    { accessorKey: 'item_id', header: 'ID Item', size: 70 },
    { accessorKey: 'operador', header: 'Operador' },
    { accessorKey: 'operation_type', header: 'Operação' },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca' },
    { accessorKey: 'model', header: 'Modelo' },
    { accessorKey: 'identificador', header: 'Identificador', cell: ({ getValue }) => <span className="num">{(getValue() as string) || '-'}</span> },
    { accessorKey: 'nota_fiscal', header: 'Nota Fiscal', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'fornecedor', header: 'Fornecedor', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'usuario', header: 'Usuário', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'cpf', header: 'CPF', cell: ({ getValue }) => <span className="num">{(getValue() as string) || '-'}</span> },
    { accessorKey: 'cargo', header: 'Cargo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'setor', header: 'Setor', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'revenda', header: 'Revenda', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'center_cost', header: 'C. Custo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'data_emprestimo', header: 'Data', cell: ({ getValue }) => <span className="num">{formatDateTime(getValue() as string)}</span> },
    { accessorKey: 'data_confirmacao', header: 'Confirmação', cell: ({ getValue }) => <span className="num">{formatDateTime(getValue() as string)}</span> },
    { accessorKey: 'data_devolucao', header: 'Devolução', cell: ({ getValue }) => <span className="num">{formatDateTime(getValue() as string)}</span> },
    { accessorKey: 'details', header: 'Detalhes', cell: ({ getValue }) => getValue() as string || '-' },
  ]

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Inteligência de Negócio"
        eyebrowDetail="Relatórios Gerenciais"
        title="Relatório Mensal"
        description="Visualize todas as operações de um determinado mês."
      />

      <div className="surface-panel p-4 space-y-3">
        <PanelHeader title="Filtros de Período" />
        <div className="flex flex-wrap items-end gap-4">
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
          <div className="flex flex-col gap-1">
            <Button variant="outline" onClick={handleExport} disabled={isExporting}>
              <Download size={14} />
              {isExporting ? 'Exportando...' : 'Exportar CSV'}
            </Button>
          </div>
        </div>
        {/*
         * Aviso 4(a): handleExport chama /reports/monthly/export só com
         * ano/mês — os filtros de unidade e tipo abaixo, por serem
         * client-side, nunca chegam ao servidor. Em vez de desabilitar a
         * exportação enquanto houver filtro ativo (o que impediria exportar
         * o mês inteiro justamente quando o usuário está com a tela filtrada
         * para conferir algo), deixamos a exportação sempre disponível e
         * avisamos explicitamente o escopo — para não repetir o problema da
         * referência, que filtra, exporta e entrega o mês inteiro sem avisar.
         */}
        <p className="text-caption text-muted-foreground">
          {`O CSV exportado sempre traz ${MONTHS[queryParams.month - 1]}/${queryParams.year} inteiro`}
          {hasActiveFilters
            ? ' — os filtros de unidade e tipo de operação abaixo não são aplicados ao arquivo.'
            : ', sem os filtros de unidade e tipo de operação abaixo.'}
        </p>
      </div>

      {!isLoading && report.length > 0 && (
        <div className="surface-panel p-4">
          <PanelHeader title="Totais do Mês" />
          <div className="grid grid-cols-2 gap-6 sm:grid-cols-4 mt-3">
            <StatBlock label="Total" value={totalRegistros} symbol="#" />
            <StatBlock label="Empréstimos" value={totalEmprestimos} symbol="↑" />
            <StatBlock
              label="Empréstimos já devolvidos"
              value={totalDevolucoes}
              hint="dos emitidos no período"
              symbol="↓"
            />
            <StatBlock label="Colaboradores Únicos" value={totalColaboradores} symbol="◯" />
          </div>
        </div>
      )}

      {!isLoading && report.length > 0 && (
        <div className="surface-panel p-4 space-y-3">
          <PanelHeader
            title="Filtros de Resultado"
            actions={hasActiveFilters ? (
              <Button variant="ghost" size="sm" onClick={handleClearFilters}>Limpar filtros</Button>
            ) : undefined}
          />
          <div className="flex flex-wrap items-end gap-4">
            <div className="flex flex-col gap-1.5">
              <Label>Unidade</Label>
              <Select value={filterUnidade} onValueChange={setFilterUnidade}>
                <SelectTrigger className="w-52"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value={ALL}>Todas as unidades</SelectItem>
                  {unidadeOptions.map((u) => <SelectItem key={u} value={u}>{u}</SelectItem>)}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Tipo de Operação</Label>
              <Select value={filterOperationType} onValueChange={setFilterOperationType}>
                <SelectTrigger className="w-56"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value={ALL}>Todos os tipos</SelectItem>
                  {operationTypeOptions.map((o) => <SelectItem key={o} value={o}>{o}</SelectItem>)}
                </SelectContent>
              </Select>
            </div>
          </div>

          {/*
           * Aviso 4(b): o ramo de periféricos do UNION do relatório
           * (backend/app/db/inventory_manager_db.py ~L980) seleciona NULL na
           * coluna de revenda — só os ramos de Empréstimo, Cadastro e
           * Exclusão carregam revenda de verdade. Filtrar por unidade
           * descarta, portanto, toda operação de periférico (Cadastro
           * Periférico, Vínculo, Desvínculo, Substituição), mesmo quando ela
           * de fato ocorreu na unidade selecionada. Não dá para corrigir sem
           * alterar o backend — só sinalizar.
           */}
          {filterUnidade !== ALL && (
            <p className="border border-border-strong bg-surface-alt p-3 rounded text-body-sm">
              Operações de periférico (Cadastro Periférico, Vínculo, Desvínculo, Substituição Periférico) não têm
              unidade registrada nesta consulta e por isso somem da lista ao filtrar por unidade — isso não significa
              que não houve movimentação de periféricos no mês.
            </p>
          )}
        </div>
      )}

      {isLoading ? (
        <div className="flex flex-col items-center gap-4 py-8">
          <Loader2 className="h-8 w-8 animate-spin text-muted-foreground" />
          <p className="text-body-sm text-muted-foreground">Gerando relatório...</p>
        </div>
      ) : report.length === 0 ? (
        // Estado vazio ciente dos filtros (Step 3), primeiro caso: o mês em
        // si não teve nenhuma movimentação — não há filtro a sugerir limpar.
        <div className="surface-panel p-8 text-center">
          <p className="text-body-sm text-muted-foreground">
            {`Nenhuma movimentação registrada em ${MONTHS[queryParams.month - 1]}/${queryParams.year}.`}
          </p>
        </div>
      ) : searchedReport.length === 0 ? (
        // Segundo caso: o mês tem dados, mas os filtros de unidade/tipo (ou a
        // busca) eliminaram todas as linhas — aqui faz sentido sugerir limpar.
        <div className="surface-panel p-8 flex flex-col items-center gap-3 text-center">
          <p className="text-body-sm text-muted-foreground">Nenhuma linha passou nos filtros selecionados.</p>
          <Button variant="outline" size="sm" onClick={handleClearFilters}>Limpar filtros</Button>
        </div>
      ) : (
        <DataTable
          data={pageRows}
          columns={columns}
          searchPlaceholder="Buscar no relatório..."
          pagination={{
            total: searchedReport.length,
            pageIndex,
            pageSize: PAGE_SIZE,
            onPageChange: setPageIndex,
            search,
            onSearchChange: setSearch,
          }}
        />
      )}
    </div>
  )
}
