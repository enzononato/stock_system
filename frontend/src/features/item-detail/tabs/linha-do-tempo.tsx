import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { ArrowLeftRight, Circle, History, Trash2, Wrench, type LucideIcon } from 'lucide-react'
import { getLinhaDoTempo, type Item } from '@/api/items'
import { Button } from '@/components/ui/button'
import { CompraResumo } from '@/features/ciclo-vida/CompraResumo'
import { TransferirItemDialog } from '@/features/ciclo-vida/TransferirItemDialog'
import { formatDateTime } from '@/lib/utils'
import type { ItemDetailTab } from '../types'

// Ícone por origem. Origens desconhecidas (as fontes de termos, linha telefônica e inventário
// entram depois) caem no círculo, para os eventos aparecerem sem mudar esta aba.
const ICONE_POR_ORIGEM: Record<string, LucideIcon> = {
  historico: History,
  manutencao: Wrench,
  transferencia: ArrowLeftRight,
  baixa: Trash2,
}

function LinhaDoTempoTab({ item }: { item: Item }) {
  const [transferirAberto, setTransferirAberto] = useState(false)
  const { data, isLoading, isError } = useQuery({
    queryKey: ['item-linha-do-tempo', item.id],
    queryFn: () => getLinhaDoTempo(item.id),
  })

  return (
    <div className="space-y-4">
      <CompraResumo item={item} />

      <div className="flex justify-end">
        <Button type="button" variant="outline" size="sm" onClick={() => setTransferirAberto(true)}>
          <ArrowLeftRight className="mr-1.5 size-3.5" />
          Transferir unidade
        </Button>
      </div>

      {isLoading && <p className="text-sm text-muted-foreground">Carregando linha do tempo...</p>}
      {isError && <p className="text-sm text-muted-foreground">Não foi possível carregar a linha do tempo.</p>}

      {data && data.fontes_com_erro.length > 0 && (
        <p role="alert" className="rounded border border-border bg-surface-alt px-3 py-2 text-sm text-foreground">
          Parte da linha do tempo não pôde ser carregada. Tente novamente mais tarde.
        </p>
      )}

      {data && data.eventos.length === 0 && (
        <p className="text-sm text-muted-foreground">Nenhum evento registrado para este equipamento.</p>
      )}

      {data && data.eventos.length > 0 && (
        <ol aria-label="Linha do tempo" className="divide-y divide-border rounded-lg border border-border">
          {data.eventos.map((e, i) => {
            const Icone = ICONE_POR_ORIGEM[e.origem] ?? Circle
            return (
              <li key={`${e.origem}-${e.referencia_id ?? 'x'}-${i}`} className="flex items-start gap-3 px-4 py-3">
                <Icone aria-hidden className="mt-0.5 size-4 shrink-0 text-muted-foreground" />
                <div className="min-w-0 flex-1">
                  <p className="text-sm font-semibold text-foreground">{e.titulo}</p>
                  {e.descricao && <p className="mt-0.5 text-xs text-muted-foreground">{e.descricao}</p>}
                  {e.usuario && <p className="mt-0.5 text-xs text-muted-foreground">{`por ${e.usuario}`}</p>}
                </div>
                <time dateTime={e.ocorrido_em} className="shrink-0 text-xs text-muted-foreground">
                  {formatDateTime(e.ocorrido_em)}
                </time>
              </li>
            )
          })}
        </ol>
      )}

      <TransferirItemDialog item={item} open={transferirAberto} onOpenChange={setTransferirAberto} />
    </div>
  )
}

const tab: ItemDetailTab = {
  id: 'linha-do-tempo',
  label: 'Linha do tempo',
  order: 20,
  roles: ['Gestor', 'Técnico'],
  Component: LinhaDoTempoTab,
}
export default tab
