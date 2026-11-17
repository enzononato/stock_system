import { useEffect, useState } from 'react'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import type { Colaborador } from '@/api/colaboradores'
import { getTermosDoColaborador } from '@/api/termos'
import { TermosLista } from '@/components/termos/TermosLista'
import { Pagination } from '@/components/ui/Pagination'
import { usePageSize } from '@/hooks/usePageSize'
import type { ColaboradorDetailTab } from '../types'

function TermosTab({ colaborador }: { colaborador: Colaborador }) {
  const [pageSize, setStoredPageSize] = usePageSize('colaborador-termos')
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
    queryKey: ['colaborador-termos', colaborador.id, page, pageSize],
    queryFn: () => getTermosDoColaborador(colaborador.id, { limit: pageSize, offset: (page - 1) * pageSize }),
    placeholderData: keepPreviousData,
  })

  const termos = data?.items ?? []
  const total = data?.total ?? 0
  const lastPage = Math.max(1, Math.ceil(total / pageSize))

  // Se o total encolheu e a página atual passou da última, ajusta o estado (e,
  // com ele, o offset da consulta) em vez de só corrigir o que é exibido.
  useEffect(() => {
    if (data && page > lastPage) setPage(lastPage)
  }, [data, page, lastPage])

  if (isLoading) return <p className="text-sm text-muted-foreground">Carregando termos...</p>
  if (isError) return <p className="text-sm text-muted-foreground">Não foi possível carregar os termos.</p>
  if (total === 0) {
    return <p className="text-sm text-muted-foreground">Nenhum termo registrado para este colaborador.</p>
  }

  return (
    <div className="space-y-3">
      <TermosLista termos={termos} mostrarItem />
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

const tab: ColaboradorDetailTab = {
  id: 'termos',
  label: 'Termos',
  order: 30,
  roles: ['Gestor', 'Técnico'],
  Component: TermosTab,
}
export default tab
