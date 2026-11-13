import type { LinhaEvento, TipoEventoLinha } from '@/api/linhas'
import { formatMoeda, formatTelefone, upper } from '@/lib/format'

export const ROTULO_EVENTO: Record<TipoEventoLinha, string> = {
  ativacao: 'Ativação',
  troca_aparelho: 'Troca de aparelho',
  troca_usuario: 'Troca de usuário',
  portabilidade: 'Portabilidade',
  suspensao: 'Suspensão',
  reativacao: 'Reativação',
  cancelamento: 'Cancelamento',
  edicao: 'Edição',
}

export const ROTULO_CAMPO: Record<string, string> = {
  numero: 'Número',
  plano: 'Plano',
  custo_mensal: 'Custo mensal',
  unidade: 'Unidade',
  iccid: 'ICCID',
  observacoes: 'Observações',
}

function valorDaEdicao(campo: string | null, valor: string | null): string {
  if (!valor) return '—'
  if (campo === 'custo_mensal') return formatMoeda(valor) || valor
  if (campo === 'numero') return formatTelefone(valor)
  if (campo === 'unidade') return upper(valor)
  return valor
}

/** Texto da linha de detalhe de um evento ("de → para" ou frase fixa do tipo). */
export function descreverEvento(e: LinhaEvento): string {
  switch (e.tipo) {
    case 'ativacao':
      return 'Linha ativada'
    case 'troca_aparelho':
      return `${e.de_valor ?? 'sem aparelho'} → ${e.para_valor ?? 'sem aparelho'}`
    case 'troca_usuario':
      return `${e.de_valor ?? 'sem usuário'} → ${e.para_valor ?? 'sem usuário'}`
    case 'portabilidade':
      return `${e.de_valor ?? '—'} → ${e.para_valor ?? '—'}`
    case 'suspensao':
      return 'Linha suspensa'
    case 'reativacao':
      return 'Linha reativada'
    case 'cancelamento':
      return 'Linha cancelada'
    case 'edicao': {
      const rotulo = (e.campo && ROTULO_CAMPO[e.campo]) ?? e.campo ?? 'Campo'
      return `${rotulo}: ${valorDaEdicao(e.campo, e.de_valor)} → ${valorDaEdicao(e.campo, e.para_valor)}`
    }
  }
}
