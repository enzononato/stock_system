import type { StatusTone } from '@/lib/status'
import type { StatusManutencao } from '@/api/manutencoes'

export const ROTULO_STATUS_MANUTENCAO: Record<StatusManutencao, string> = {
  aberta: 'Aberta',
  enviada: 'Enviada',
  em_reparo: 'Em reparo',
  retornada: 'Retornada',
  sem_conserto: 'Sem conserto',
  cancelada: 'Cancelada',
}

/** Passado sempre como `tone` do `StatusPill`: não mexe em `lib/status.ts`. */
export const TOM_STATUS_MANUTENCAO: Record<StatusManutencao, StatusTone> = {
  aberta: 'pending',
  enviada: 'maintenance',
  em_reparo: 'maintenance',
  retornada: 'available',
  sem_conserto: 'danger',
  cancelada: 'retired',
}

export const STATUS_EM_ANDAMENTO: StatusManutencao[] = ['aberta', 'enviada', 'em_reparo']
