import { useEffect, useState } from 'react'
import { Link } from 'react-router-dom'
import { useQuery, keepPreviousData } from '@tanstack/react-query'
import type { Colaborador } from '@/api/colaboradores'
import { listItemsPaginated } from '@/api/items'
import { Pagination } from '@/components/ui/Pagination'
import { StatusPill } from '@/components/ui/StatusPill'
import { usePageSize } from '@/hooks/usePageSize'
import type { ColaboradorDetailTab } from '../types'

function EquipamentosTab({ colaborador }: { colaborador: Colaborador }) {
  const [pageSize, setStoredPageSize] = usePageSize('colaborador-equipamentos')
  const [page, setPage] = useState(1)

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
    queryKey: ['colaborador-equipamentos', colaborador.id, page, pageSize],
    queryFn: () =>
      listItemsPaginated({ colaborador_id: colaborador.id, limit: pageSize, offset: (page - 1) * pageSize }),
    placeholderData: keepPreviousData,
  })

  const itens = data?.items ?? []
  const total = data?.total ?? 0
  const lastPage = Math.max(1, Math.ceil(total / pageSize))

  // Se o total encolheu e a página atual passou da última, ajusta o estado (e,
  // com ele, o offset da consulta) em vez de só corrigir o que é exibido.
  useEffect(() => {
    if (data && page > lastPage) setPage(lastPage)
  }, [data, page, lastPage])

  if (isLoading) return <p className="text-sm text-muted-foreground">Carregando equipamentos...</p>
  if (isError) return <p className="text-sm text-muted-foreground">Não foi possível carregar os equipamentos.</p>
  if (total === 0) {
    return <p className="text-sm text-muted-foreground">Nenhum equipamento com este colaborador.</p>
  }

  return (
    <div className="space-y-3">
      <ul className="divide-y divide-border rounded-lg border border-border">
        {itens.map((item) => {
          const nome = [item.brand, item.model].filter(Boolean).join(' ') || item.tipo || `Equipamento ${item.id}`
          return (
            <li key={item.id} className="flex flex-wrap items-center justify-between gap-2 px-4 py-3">
              <div className="min-w-0">
                <Link
                  to={`/items/${item.id}`}
                  className="rounded-sm text-sm font-semibold text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2"
                >
                  {nome}
                </Link>
                <p className="text-xs text-muted-foreground">
                  {[item.tipo, item.identificador].filter(Boolean).join(' · ') || '—'}
                </p>
              </div>
              <StatusPill status={item.status} />
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
    </div>
  )
}

const tab: ColaboradorDetailTab = { id: 'equipamentos', label: 'Equipamentos', order: 20, Component: EquipamentosTab }
export default tab
