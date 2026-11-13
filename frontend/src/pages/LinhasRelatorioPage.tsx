import { useState } from 'react'
import { Link } from 'react-router-dom'
import { useQuery } from '@tanstack/react-query'
import { isAxiosError } from 'axios'
import { ArrowLeft, FileSpreadsheet, FileText } from 'lucide-react'
import { exportarRelatorioCusto, getRelatorioCusto } from '@/api/linhas'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { formatMoeda, upper } from '@/lib/format'
import { formatDateTime } from '@/lib/utils'

const MSG_EXPORTAR = 'Não foi possível exportar o relatório.'

/** O download usa `responseType: 'blob'`: o `detail` do erro vem dentro do Blob, não no objeto. */
async function mensagemDeErro(err: unknown): Promise<string> {
  if (isAxiosError(err) && err.response?.data instanceof Blob) {
    try {
      const corpo = JSON.parse(await err.response.data.text()) as { detail?: unknown }
      if (typeof corpo.detail === 'string' && corpo.detail.trim()) return corpo.detail
    } catch {
      /* corpo que não é JSON: usa a mensagem padrão abaixo */
    }
    return MSG_EXPORTAR
  }
  return getErrorMessage(err, MSG_EXPORTAR)
}

export default function LinhasRelatorioPage() {
  const [exportando, setExportando] = useState<'xlsx' | 'pdf' | null>(null)

  const { data, isLoading, isError, error, refetch } = useQuery({
    queryKey: ['linhas', 'relatorio-custo'],
    queryFn: getRelatorioCusto,
  })

  async function exportar(formato: 'xlsx' | 'pdf') {
    setExportando(formato)
    try {
      await exportarRelatorioCusto(formato)
    } catch (err) {
      toast(await mensagemDeErro(err), 'error')
    } finally {
      setExportando(null)
    }
  }

  return (
    <div className="space-y-6">
      <Link
        to="/linhas"
        className="inline-flex items-center gap-1.5 rounded-sm text-sm text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
      >
        <ArrowLeft size={14} aria-hidden="true" />
        Voltar para linhas
      </Link>

      <PageHeader
        eyebrow="Telefonia"
        title="Custo mensal de linhas"
        description="Somente linhas ativas, somadas por unidade."
        actions={
          <>
            <Button variant="outline" loading={exportando === 'xlsx'} disabled={exportando === 'pdf'} onClick={() => exportar('xlsx')}>
              <FileSpreadsheet aria-hidden="true" />
              Exportar Excel
            </Button>
            <Button variant="outline" loading={exportando === 'pdf'} disabled={exportando === 'xlsx'} onClick={() => exportar('pdf')}>
              <FileText aria-hidden="true" />
              Exportar PDF
            </Button>
          </>
        }
      />

      {isLoading ? (
        <LoadingState label="Carregando relatório..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : !data || data.unidades.length === 0 ? (
        <EmptyState
          title="Nenhuma linha ativa"
          description="O relatório lista as unidades que têm pelo menos uma linha ativa."
        />
      ) : (
        <div className="space-y-2">
          <div className="surface-panel overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="border-b border-border bg-surface-alt">
                  <tr>
                    <th className="whitespace-nowrap px-4 py-2.5 text-left text-caption text-muted-foreground">Unidade</th>
                    <th className="whitespace-nowrap px-4 py-2.5 text-right text-caption text-muted-foreground">
                      Linhas ativas
                    </th>
                    <th className="whitespace-nowrap px-4 py-2.5 text-right text-caption text-muted-foreground">
                      Custo mensal
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {data.unidades.map((u) => (
                    <tr key={u.unidade_id} className="border-b border-border last:border-0">
                      <td className="whitespace-nowrap px-4 py-2.5 text-body-sm font-medium">{upper(u.unidade_nome)}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-right text-body-sm tabular-nums">{u.quantidade}</td>
                      <td className="whitespace-nowrap px-4 py-2.5 text-right text-body-sm tabular-nums">
                        {formatMoeda(u.custo_total)}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot className="border-t border-border bg-surface-alt">
                  <tr>
                    <th scope="row" className="whitespace-nowrap px-4 py-2.5 text-left text-body-sm font-semibold">
                      Total
                    </th>
                    <td className="whitespace-nowrap px-4 py-2.5 text-right text-body-sm font-semibold tabular-nums">
                      {data.total_linhas}
                    </td>
                    <td className="whitespace-nowrap px-4 py-2.5 text-right text-body-sm font-semibold tabular-nums">
                      {formatMoeda(data.custo_total)}
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>
          <p className="text-caption text-muted-foreground">Gerado em {formatDateTime(data.gerado_em)}</p>
        </div>
      )}
    </div>
  )
}
