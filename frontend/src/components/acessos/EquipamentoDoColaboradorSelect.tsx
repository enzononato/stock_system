import { useQuery } from '@tanstack/react-query'
import { listEquipamentosDoColaborador } from '@/api/acessos'
import { descreverEquipamento } from '@/components/acessos/rotulos'

const selectClass =
  'flex h-9 w-full rounded border border-input bg-surface-alt px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-40'

interface EquipamentoDoColaboradorSelectProps {
  colaboradorId: number
  value: number | null
  onChange: (itemId: number | null) => void
  id?: string
  disabled?: boolean
}

/** Equipamentos que estão com o colaborador (para dizer onde o acesso foi instalado). */
export function EquipamentoDoColaboradorSelect({
  colaboradorId,
  value,
  onChange,
  id,
  disabled,
}: EquipamentoDoColaboradorSelectProps) {
  const { data: equipamentos = [] } = useQuery({
    queryKey: ['acessos', 'equipamentos', colaboradorId],
    queryFn: () => listEquipamentosDoColaborador(colaboradorId),
  })
  const foraDaLista = value !== null && !equipamentos.some((e) => e.id === value)

  return (
    <select
      id={id}
      className={selectClass}
      disabled={disabled}
      value={value === null ? '' : String(value)}
      onChange={(e) => onChange(e.target.value === '' ? null : Number(e.target.value))}
    >
      <option value="">Nenhum</option>
      {foraDaLista && <option value={String(value)}>{`Equipamento #${value}`}</option>}
      {equipamentos.map((e) => (
        <option key={e.id} value={String(e.id)}>
          {descreverEquipamento(e)}
        </option>
      ))}
    </select>
  )
}
