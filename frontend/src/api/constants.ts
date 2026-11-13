import api from './client'

export interface Constants {
  center_costs: string[]
  revendas: string[]
  setores: string[]
  equipment_types: string[]
  peripheral_types: string[]
  removal_reasons: string[]
  removal_reasons_attachment: Record<string, boolean>
  /** Status possíveis de um equipamento (inclui `Em Manutenção`). */
  item_statuses?: string[]
  /** Tipos cuja baixa exige registro de apagamento de dados (LGPD). */
  tipos_exigem_apagamento?: string[]
  /** Motivos de baixa em que "não se aplica" vale (roubo, perda). */
  motivos_sem_posse?: string[]
  /** Chave → rótulo dos métodos de apagamento. */
  metodos_apagamento?: Record<string, string>
}

export async function getConstants() {
  const res = await api.get('/constants')
  return res.data as Constants
}
