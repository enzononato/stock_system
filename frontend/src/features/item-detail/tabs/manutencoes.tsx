import { useEffect, useState } from 'react'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import { Link } from 'react-router-dom'
import type { Item } from '@/api/items'
import { listManutencoes } from '@/api/manutencoes'
import { Button } from '@/components/ui/button'
import { Pagination } from '@/components/ui/Pagination'
import { useAuth } from '@/contexts/AuthContext'
import { AbrirManutencaoDialog } from '@/features/manutencoes/AbrirManutencaoDialog'
import { ManutencoesLista } from '@/features/manutencoes/ManutencoesLista'
import { podeAbrirManutencao } from '@/features/manutencoes/regras'
import { usePageSize } from '@/hooks/usePageSize'
import { formatMoeda } from '@/lib/format'
import type { ItemDetailTab } from '../types'

function ManutencoesTab({ item }: { item: Item }) {
  const { hasRole } = useAuth()
  const [pageSize, setStoredPageSize] = usePageSize('item-manutencoes')
  const [page, setPage] = useState(1)
  const [abrirAberto, setAbrirAberto] = useState(false)

  // Trocar o tamanho da página volta para a primeira no mesmo evento.
  const setPageSize = (n: number) => {
    setStoredPageSize(n)
    setPage(1)
  }

  // Volta para a primeira página ao trocar de equipamento.
  useEffect(() => {
    setPage(1)
  }, [item.id])

  const { data, isLoading, isError, isFetching } = useQuery({
    queryKey: ['manutencoes', { itemId: item.id, page, pageSize }],
    queryFn: () => listManutencoes({ item_id: item.id, limit: pageSize, offset: (page - 1) * pageSize }),
    placeholderData: keepPreviousData,
  })

  // Consulta própria (não depende da página aberta): existe ordem em andamento?
  const { data: andamento } = useQuery({
    queryKey: ['manutencoes', { itemId: item.id, status: 'em_andamento' }],
    queryFn: () => listManutencoes({ item_id: item.id, status: 'em_andamento', limit: 1 }),
  })
  const temEmAndamento = (andamento?.total ?? 0) > 0

  const total = data?.total ?? 0
  const lastPage = Math.max(1, Math.ceil(total / pageSize))
  useEffect(() => {
    if (data && page > lastPage) setPage(lastPage)
  }, [data, page, lastPage])

  const regra = podeAbrirManutencao(item, temEmAndamento)
  const semConserto =
    item.status === 'Em Manutenção' &&
    !temEmAndamento &&
    (data?.items ?? []).some((m) => m.status === 'sem_conserto')

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <p className="text-sm font-semibold text-foreground">
          {`Custo acumulado: ${formatMoeda(data?.custo_total ?? 0)}`}
        </p>
        <div className="flex flex-wrap items-center gap-3">
          {!regra.pode && regra.motivo && <p className="text-xs text-muted-foreground">{regra.motivo}</p>}
          <Button type="button" size="sm" disabled={!regra.pode} onClick={() => setAbrirAberto(true)}>
            Abrir manutenção
          </Button>
        </div>
      </div>

      {semConserto && (
        <div className="flex flex-wrap items-center gap-3 rounded border border-border bg-surface-alt px-3 py-2 text-sm text-foreground">
          <span>Equipamento sem conserto. Avalie a baixa.</span>
          {hasRole('Gestor') && (
            <Link to={`/remove?item=${item.id}`} className="font-medium text-brand-text hover:underline">
              Dar baixa
            </Link>
          )}
        </div>
      )}

      {isLoading && <p className="text-sm text-muted-foreground">Carregando manutenções...</p>}
      {isError && <p className="text-sm text-muted-foreground">Não foi possível carregar as manutenções.</p>}
      {data && total === 0 && (
        <p className="text-sm text-muted-foreground">Nenhuma manutenção registrada para este equipamento.</p>
      )}
      {data && total > 0 && (
        <div className="space-y-3">
          <ManutencoesLista manutencoes={data.items} mostrarEquipamento={false} />
          <Pagination
            page={Math.min(page, lastPage)}
            pageSize={pageSize}
            total={total}
            onPageChange={setPage}
            onPageSizeChange={setPageSize}
            isFetching={isFetching}
          />
        </div>
      )}

      <AbrirManutencaoDialog item={item} open={abrirAberto} onOpenChange={setAbrirAberto} />
    </div>
  )
}

const tab: ItemDetailTab = {
  id: 'manutencoes',
  label: 'Manutenções',
  order: 40,
  roles: ['Gestor', 'Técnico'],
  Component: ManutencoesTab,
}
export default tab
