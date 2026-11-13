import { useEffect } from 'react'
import { useQuery } from '@tanstack/react-query'
import { listLocais } from '@/api/locais'

interface LocalEncontradoSelectProps {
  unidadeId: number
  value: number | null
  onChange: (id: number | null) => void
}

const selectClass =
  'flex h-11 w-full rounded border border-input bg-surface-alt px-3 py-2 text-base text-foreground focus-visible:ring-1 focus-visible:ring-ring'

/**
 * Seletor "Estou no local" da tela de leitura. Mesmo comportamento do `LocalSelect`
 * compartilhado (locais ativos da unidade, mesma chave de cache, limpa o valor que
 * não pertence à unidade), mas sem o `aria-label` fixo "Local": o nome acessível
 * vem do rótulo visível que envolve o campo, para os dois coincidirem.
 */
export function LocalEncontradoSelect({ unidadeId, value, onChange }: LocalEncontradoSelectProps) {
  const { data: locais = [], isSuccess } = useQuery({
    queryKey: ['locais', { unidadeId, includeInactive: true }],
    queryFn: () => listLocais({ unidade_id: unidadeId, include_inactive: true }),
  })

  useEffect(() => {
    if (value !== null && isSuccess && !locais.some((l) => l.id === value)) onChange(null)
    // `onChange` fica fora das dependências de propósito: só reagimos a mudança de dados.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value, isSuccess, locais])

  return (
    <select
      className={selectClass}
      value={value === null ? '' : String(value)}
      onChange={(e) => onChange(e.target.value === '' ? null : Number(e.target.value))}
    >
      <option value="">Sem local</option>
      {locais
        .filter((l) => l.ativo || l.id === value)
        .map((l) => (
          <option key={l.id} value={l.id}>
            {l.ativo ? l.nome : `${l.nome} (inativo)`}
          </option>
        ))}
    </select>
  )
}
