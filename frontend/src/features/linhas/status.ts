import type { StatusLinha } from '@/api/linhas'
import type { StatusTone } from '@/lib/status'

export type AcaoStatusLinha = 'suspender' | 'reativar' | 'cancelar'

export const STATUS_LINHA_OPCOES: { value: StatusLinha; label: string; tone: StatusTone }[] = [
  { value: 'ativa', label: 'Ativa', tone: 'available' },
  { value: 'suspensa', label: 'Suspensa', tone: 'pending' },
  { value: 'cancelada', label: 'Cancelada', tone: 'retired' },
]

export function statusLinhaInfo(status: StatusLinha): { label: string; tone: StatusTone } {
  const opcao = STATUS_LINHA_OPCOES.find((o) => o.value === status)
  return opcao ? { label: opcao.label, tone: opcao.tone } : { label: status, tone: 'neutral' }
}

/** Ações que o backend aceita a partir de cada status (cancelada só volta a ficar ativa). */
export function acoesDisponiveis(status: StatusLinha): AcaoStatusLinha[] {
  if (status === 'ativa') return ['suspender', 'cancelar']
  if (status === 'suspensa') return ['reativar', 'cancelar']
  return ['reativar']
}

export const STATUS_DA_ACAO: Record<AcaoStatusLinha, StatusLinha> = {
  suspender: 'suspensa',
  reativar: 'ativa',
  cancelar: 'cancelada',
}
