import type { AcaoTarefa, StatusAcesso, StatusEvento, StatusTarefa, TipoEvento } from '@/api/acessos'
import type { StatusTone } from '@/lib/status'

export const ROTULO_TIPO_EVENTO: Record<TipoEvento, string> = {
  admissao: 'Admissão',
  desligamento: 'Desligamento',
  mudanca_funcao: 'Mudança de função',
}

export const STATUS_EVENTO: Record<StatusEvento, { label: string; tone: StatusTone }> = {
  aberto: { label: 'Aberto', tone: 'pending' },
  concluido: { label: 'Concluído', tone: 'available' },
  cancelado: { label: 'Cancelado', tone: 'retired' },
}

export const STATUS_TAREFA: Record<StatusTarefa, { label: string; tone: StatusTone }> = {
  pendente: { label: 'Pendente', tone: 'pending' },
  feito: { label: 'Feito', tone: 'available' },
  nao_se_aplica: { label: 'Não se aplica', tone: 'neutral' },
}

export const ROTULO_ACAO_TAREFA: Record<AcaoTarefa, string> = {
  criar: 'Criar',
  desativar: 'Desativar',
}

export const STATUS_ACESSO: Record<StatusAcesso, { label: string; tone: StatusTone }> = {
  ativo: { label: 'Ativo', tone: 'available' },
  desativado: { label: 'Desativado', tone: 'retired' },
}

/** Mesmo formato do backend (`descrever_item`): "Notebook Dell X · NB-1"; sem dados, "Equipamento #<id>". */
export function descreverEquipamento(e: {
  id: number
  tipo?: string | null
  brand?: string | null
  model?: string | null
  identificador?: string | null
}): string {
  const base = [e.tipo, e.brand, e.model]
    .map((p) => (p ?? '').trim())
    .filter(Boolean)
    .join(' ')
  const identificador = (e.identificador ?? '').trim()
  const texto = identificador ? (base ? `${base} · ${identificador}` : identificador) : base
  return texto || `Equipamento #${e.id}`
}
