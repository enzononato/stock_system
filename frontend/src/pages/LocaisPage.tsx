import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { listUnidades } from '@/api/unidades'
import { LocaisManager } from '@/components/locais/LocaisManager'
import { PageHeader } from '@/components/layout/PageHeader'
import { EmptyState } from '@/components/ui/StateBlocks'
import { upper } from '@/lib/format'

const selectClass =
  'h-9 min-w-[240px] rounded border border-input bg-surface-alt px-3 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring'

/** Gestão de locais por unidade para Gestor e Técnico (a tela de Unidades é só do Gestor). */
export default function LocaisPage() {
  const [unidadeId, setUnidadeId] = useState('')

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades', { includeInactive: false }],
    queryFn: () => listUnidades(),
  })

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Inventário"
        eyebrowDetail="Locais por unidade"
        title="Locais"
        description="Cadastre e organize os locais físicos (salas, almoxarifados...) de cada unidade."
      />

      <div className="surface-panel flex flex-wrap items-end gap-3 p-3">
        <label className="flex flex-col gap-1 text-caption text-muted-foreground">
          Unidade
          <select
            aria-label="Unidade"
            className={selectClass}
            value={unidadeId}
            onChange={(e) => setUnidadeId(e.target.value)}
          >
            <option value="">Selecione uma unidade...</option>
            {unidades.map((u) => (
              <option key={u.id} value={u.id}>
                {upper(u.nome)}
              </option>
            ))}
          </select>
        </label>
      </div>

      {unidadeId === '' ? (
        <EmptyState title="Selecione uma unidade" description="Escolha uma unidade para ver e gerenciar os locais dela." />
      ) : (
        <LocaisManager key={unidadeId} unidade={{ id: Number(unidadeId) }} />
      )}
    </div>
  )
}
