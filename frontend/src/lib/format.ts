import { formatDate } from './utils'

/** Maiúsculas respeitando pt-BR; vazio vira ''. */
export function upper(v: string | null | undefined): string {
  return v ? v.toLocaleUpperCase('pt-BR') : ''
}

/** `(74) 99999-0000` (celular) ou `(74) 3611-0000` (fixo). Outros formatos voltam como vieram. */
export function formatTelefone(v: string | null | undefined): string {
  if (!v) return ''
  const d = v.replace(/\D/g, '')
  if (d.length === 11) return `(${d.slice(0, 2)}) ${d.slice(2, 7)}-${d.slice(7)}`
  if (d.length === 10) return `(${d.slice(0, 2)}) ${d.slice(2, 6)}-${d.slice(6)}`
  return v
}

const BRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' })

/** `R$ 1.234,56`; vazio ou não numérico vira ''. */
export function formatMoeda(v: number | string | null | undefined): string {
  if (v === null || v === undefined || v === '') return ''
  const n = typeof v === 'number' ? v : Number(v)
  return Number.isFinite(n) ? BRL.format(n) : ''
}

export const formatDataBR = formatDate
