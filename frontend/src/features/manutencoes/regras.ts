import type { Item } from '@/api/items'

/**
 * Se dá para abrir uma manutenção para o equipamento (espelha o backend, que recusa
 * pelos mesmos motivos). Equipamento `Em Manutenção` sem ordem em andamento é o caso
 * "sem conserto": uma nova manutenção (segunda assistência) pode ser aberta.
 */
export function podeAbrirManutencao(
  item: Item,
  temEmAndamento: boolean
): { pode: boolean; motivo?: string } {
  if (item.colaborador_id != null || item.assigned_to) {
    return { pode: false, motivo: 'Equipamento com colaborador: faça a devolução antes de abrir a manutenção.' }
  }
  if (temEmAndamento) {
    return { pode: false, motivo: 'Já existe uma manutenção em andamento.' }
  }
  if (item.status !== 'Disponível' && item.status !== 'Em Manutenção') {
    return { pode: false, motivo: 'Só é possível abrir manutenção para equipamento disponível.' }
  }
  return { pode: true }
}
