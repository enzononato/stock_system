import { STATUS_TERMO_ROTULO, type TermoStatus } from '@/api/termos'
import { StatusPill } from '@/components/ui/StatusPill'

/** Pílula do status de uma solicitação de termo ("manual" aparece como pendência em andamento). */
export function TermoStatusPill({ status }: { status: TermoStatus }) {
  return <StatusPill status={STATUS_TERMO_ROTULO[status]} tone={status === 'manual' ? 'pending' : undefined} />
}
