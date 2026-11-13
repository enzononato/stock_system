import { useEffect, useState } from 'react'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import type { Item } from '@/api/items'
import { getTermosDoItem } from '@/api/termos'
import { TermoVigenteCard } from '@/components/termos/TermoVigenteCard'
import { TermosLista } from '@/components/termos/TermosLista'
import { Pagination } from '@/components/ui/Pagination'
import { usePageSize } from '@/hooks/usePageSize'
import type { ItemDetailTab } from '../types'

function TermosTab({ item }: { item: Item }) {
  const [pageSize, setStoredPageSize] = usePageSize('item-termos')
  const [page, setPage] = useState(1)

  // Trocar o tamanho da página volta para a primeira no mesmo evento, para não
  // disparar uma requisição com a página antiga e o tamanho novo.
  const setPageSize = (n: number) => {
    setStoredPageSize(n)
    setPage(1)
  }

  // Volta para a primeira página ao trocar de equipamento.
  useEffect(() => {
    setPage(1)
  }, [item.id])

  const { data, isLoading, isError, isFetching } = useQuery({
    queryKey: ['item-termos', item.id, page, pageSize],
    queryFn: () => getTermosDoItem(item.id, { limit: pageSize, offset: (page - 1) * pageSize }),
    placeholderData: keepPreviousData,
  })

  const historico = data?.historico.items ?? []
  const total = data?.historico.total ?? 0
  const lastPage = Math.max(1, Math.ceil(total / pageSize))

  // Se o total encolheu e a página atual passou da última, ajusta o estado (e,
  // com ele, o offset da consulta) em vez de só corrigir o que é exibido.
  useEffect(() => {
    if (data && page > lastPage) setPage(lastPage)
  }, [data, page, lastPage])

  if (isLoading) return <p className="text-sm text-muted-foreground">Carregando termos...</p>
  if (isError || !data) return <p className="text-sm text-muted-foreground">Não foi possível carregar os termos.</p>

  return (
    <div className="space-y-6">
      <TermoVigenteCard
        itemId={item.id}
        itemStatus={item.status}
        vigente={data.vigente}
        revsignAtivo={data.revsign_ativo}
      />

      <section className="space-y-3">
        <h3 className="text-heading-sm text-foreground">Histórico de solicitações</h3>
        {total === 0 ? (
          <p className="text-sm text-muted-foreground">Nenhuma solicitação de assinatura registrada.</p>
        ) : (
          <>
            <TermosLista termos={historico} />
            <Pagination
              page={Math.min(page, lastPage)}
              pageSize={pageSize}
              total={total}
              onPageChange={setPage}
              onPageSizeChange={setPageSize}
              isFetching={isFetching}
            />
          </>
        )}
      </section>
    </div>
  )
}

const tab: ItemDetailTab = {
  id: 'termos',
  label: 'Termos',
  order: 30,
  roles: ['Gestor', 'Técnico'],
  Component: TermosTab,
}
export default tab
