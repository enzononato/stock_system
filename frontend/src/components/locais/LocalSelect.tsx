import { useQuery } from '@tanstack/react-query'
import { listLocais } from '@/api/locais'

export interface LocalSelectProps {
  /** Unidade cujos locais são oferecidos; `null` desabilita o seletor. */
  unidadeId: number | null
  /** `id` do local selecionado (ou `null` = "Sem local"). */
  value: number | null
  onChange: (id: number | null) => void
  disabled?: boolean
}

const selectClass =
  'flex h-9 w-full rounded border border-input bg-surface-alt px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-40'

/**
 * Seletor de local de uma unidade. Oferece só locais ativos, mas mantém
 * visível (marcado "inativo") o local já selecionado que foi desativado depois,
 * para o rótulo não sumir e o valor não ser trocado sem o usuário perceber.
 */
export function LocalSelect({ unidadeId, value, onChange, disabled }: LocalSelectProps) {
  // Mesma chave do diálogo de locais: as duas telas compartilham o cache.
  const { data: locais = [] } = useQuery({
    queryKey: ['locais', { unidadeId, includeInactive: true }],
    queryFn: () => listLocais({ unidade_id: unidadeId as number, include_inactive: true }),
    enabled: unidadeId !== null,
  })

  const opcoes = unidadeId === null ? [] : locais.filter((l) => l.ativo || l.id === value)

  return (
    <select
      aria-label="Local"
      className={selectClass}
      value={value === null ? '' : String(value)}
      disabled={disabled || unidadeId === null}
      onChange={(e) => onChange(e.target.value === '' ? null : Number(e.target.value))}
    >
      <option value="">Sem local</option>
      {opcoes.map((l) => (
        <option key={l.id} value={l.id}>
          {l.ativo ? l.nome : `${l.nome} (inativo)`}
        </option>
      ))}
    </select>
  )
}
