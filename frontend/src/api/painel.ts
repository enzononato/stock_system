import api from './client'

/** Espelho de `app/schemas/painel.py`; decimais e datas chegam como texto. */
export type StatusEquipamento =
  | 'Disponível'
  | 'Indisponível'
  | 'Pendente'
  | 'Pendente Devolução'
  | 'Em Manutenção'

export interface EquipamentosPorUnidade {
  /** `null` = equipamentos sem unidade. */
  unidade: string | null
  total: number
  por_status: Record<string, number>
}

export interface ResumoPainel {
  gerado_em: string
  equipamentos: {
    total: number
    por_status: Record<string, number>
    por_unidade: EquipamentosPorUnidade[]
  }
  manutencao: { equipamentos_em_manutencao: number; ordens_em_andamento: number }
  termos: {
    pendentes: number
    por_status: Record<string, number>
    sem_solicitacao: number
  }
  linhas: {
    ativas: number
    custo_mensal: string
    por_unidade: { unidade_nome: string; quantidade: number; custo_total: string }[]
  }
}

export type TipoAlerta =
  | 'colaborador_desligado'
  | 'devolucao_pendente'
  | 'termo_pendente'
  | 'manutencao_parada'
  | 'linha_aparelho_baixado'
  | 'garantia_vencendo'

export type SeveridadeAlerta = 'alta' | 'media' | 'baixa'

export interface Alerta {
  chave: string
  tipo: TipoAlerta
  severidade: SeveridadeAlerta
  titulo: string
  descricao: string
  /** Caminho interno do registro (ex.: `/items/9?aba=termos`). */
  link: string
  item_id: number | null
  colaborador_id: number | null
  desde: string | null
  dias: number | null
}

export interface AlertasPainel {
  gerado_em: string
  limites: {
    garantia_dias: number
    termo_pendente_dias: number
    devolucao_pendente_dias: number
    manutencao_dias: number
  }
  total: number
  por_tipo: Record<string, number>
  alertas: Alerta[]
}

export async function getResumoPainel(): Promise<ResumoPainel> {
  const res = await api.get('/painel/resumo')
  return res.data as ResumoPainel
}

export async function getAlertasPainel(): Promise<AlertasPainel> {
  const res = await api.get('/painel/alertas')
  return res.data as AlertasPainel
}
