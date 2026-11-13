import { useEffect, useState } from 'react'
import { Link } from 'react-router-dom'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import type { Colaborador } from '@/api/colaboradores'
import { listLinhas, type Linha } from '@/api/linhas'
import { LinhaHistoricoDialog } from '@/components/linhas/LinhaHistoricoDialog'
import { Button } from '@/components/ui/button'
import { Pagination } from '@/components/ui/Pagination'
import { StatusPill } from '@/components/ui/StatusPill'
import { descreverAparelho } from '@/features/linhas/format'
import { statusLinhaInfo } from '@/features/linhas/status'
import { usePageSize } from '@/hooks/usePageSize'
import { formatMoeda, formatTelefone } from '@/lib/format'
import type { ColaboradorDetailTab } from '../types'

function LinhasTab({ colaborador }: { colaborador: Colaborador }) {
  const [pageSize, setStoredPageSize] = usePageSize('colaborador-linhas')
  const [page, setPage] = useState(1)
  const [historico, setHistorico] = useState<Linha | null>(null)

  // Trocar o tamanho da página volta para a primeira no mesmo evento, para não
  // disparar uma requisição com a página antiga e o tamanho novo.
  const setPageSize = (n: number) => {
    setStoredPageSize(n)
    setPage(1)
  }

  // Volta para a primeira página ao trocar de colaborador (ex.: link do superior).
  useEffect(() => {
    setPage(1)
  }, [colaborador.id])

  const { data, isLoading, isError, isFetching } = useQuery({
    queryKey: ['linhas', 'colaborador', colaborador.id, page, pageSize],
    queryFn: () => listLinhas({ colaborador_id: colaborador.id, limit: pageSize, offset: (page - 1) * pageSize }),
    placeholderData: keepPreviousData,
  })

  const linhas = data?.items ?? []
  const total = data?.total ?? 0
  const lastPage = Math.max(1, Math.ceil(total / pageSize))

  // Se o total encolheu e a página atual passou da última, ajusta o estado (e,
  // com ele, o offset da consulta) em vez de só corrigir o que é exibido.
  useEffect(() => {
    if (data && page > lastPage) setPage(lastPage)
  }, [data, page, lastPage])

  if (isLoading) return <p className="text-sm text-muted-foreground">Carregando linhas...</p>
  if (isError) return <p className="text-sm text-muted-foreground">Não foi possível carregar as linhas.</p>
  if (total === 0) return <p className="text-sm text-muted-foreground">Nenhuma linha com este colaborador.</p>

  return (
    <div className="space-y-3">
      <ul className="divide-y divide-border rounded-lg border border-border">
        {linhas.map((l) => {
          const info = statusLinhaInfo(l.status)
          const numero = formatTelefone(l.numero)
          const aparelho = descreverAparelho({
            tipo: l.item_tipo,
            brand: l.item_brand,
            model: l.item_model,
            identificador: l.item_identificador,
          })
          return (
            <li key={l.id} className="flex flex-wrap items-center justify-between gap-3 px-4 py-3">
              <div className="min-w-0 space-y-0.5">
                <p className="text-sm font-semibold tabular-nums text-foreground">{numero}</p>
                <p className="text-xs text-muted-foreground">
                  <span>{[l.operadora, l.plano].filter(Boolean).join(' · ')}</span>
                  <span> · {formatMoeda(l.custo_mensal)}</span>
                </p>
                <p className="text-xs">
                  {l.item_id !== null ? (
                    <Link
                      to={`/items/${l.item_id}`}
                      className="rounded-sm text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
                    >
                      {aparelho || `Equipamento ${l.item_id}`}
                    </Link>
                  ) : (
                    <span className="text-muted-foreground">Sem aparelho</span>
                  )}
                  {l.item_ativo === false && <span className="text-muted-foreground"> (baixado)</span>}
                </p>
              </div>
              <div className="flex items-center gap-2">
                <StatusPill status={info.label} tone={info.tone} />
                <Button size="sm" variant="ghost" aria-label={`Histórico ${numero}`} onClick={() => setHistorico(l)}>
                  Histórico
                </Button>
              </div>
            </li>
          )
        })}
      </ul>
      <Pagination
        page={Math.min(page, lastPage)}
        pageSize={pageSize}
        total={total}
        onPageChange={setPage}
        onPageSizeChange={setPageSize}
        isFetching={isFetching}
      />
      <LinhaHistoricoDialog
        linha={historico}
        open={historico !== null}
        onOpenChange={(open) => {
          if (!open) setHistorico(null)
        }}
      />
    </div>
  )
}

const tab: ColaboradorDetailTab = { id: 'linhas', label: 'Linhas', order: 40, Component: LinhasTab }
export default tab
