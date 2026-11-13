import type { Alerta, SeveridadeAlerta, TipoAlerta } from '@/api/painel'
import type { StatusTone } from '@/lib/status'

export const ORDEM_TIPOS_ALERTA: TipoAlerta[] = [
  'colaborador_desligado',
  'devolucao_pendente',
  'termo_pendente',
  'manutencao_parada',
  'linha_aparelho_baixado',
  'garantia_vencendo',
]

export const ROTULO_TIPO_ALERTA: Record<TipoAlerta, string> = {
  colaborador_desligado: 'Colaborador desligado com equipamento',
  devolucao_pendente: 'Devolução pendente',
  termo_pendente: 'Termo pendente',
  manutencao_parada: 'Manutenção parada',
  linha_aparelho_baixado: 'Linha ativa em aparelho baixado',
  garantia_vencendo: 'Garantia vencendo',
}

export const ROTULO_SEVERIDADE: Record<SeveridadeAlerta, string> = {
  alta: 'Alta',
  media: 'Média',
  baixa: 'Baixa',
}

export const TOM_SEVERIDADE: Record<SeveridadeAlerta, StatusTone> = {
  alta: 'danger',
  media: 'pending',
  baixa: 'neutral',
}

/** Agrupa por tipo na ordem de `ORDEM_TIPOS_ALERTA`; só os grupos que têm alertas. Mantém a ordem recebida dentro do grupo. */
export function agruparAlertas(alertas: Alerta[]): { tipo: TipoAlerta; alertas: Alerta[] }[] {
  return ORDEM_TIPOS_ALERTA.map((tipo) => ({
    tipo,
    alertas: alertas.filter((a) => a.tipo === tipo),
  })).filter((grupo) => grupo.alertas.length > 0)
}
