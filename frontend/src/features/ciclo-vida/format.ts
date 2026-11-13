/**
 * Valores e datas digitados nas telas do ciclo de vida (compra, garantia,
 * manutenção, apagamento de dados).
 *
 * Datas só-dia chegam do backend como `AAAA-MM-DD`. Passar isso por `new Date()`
 * interpreta como UTC e, no fuso de Brasília, mostra um dia a menos — por isso
 * `isoParaBR` converte só por texto.
 */

/** Limite do anexo (NF, anexo da manutenção, comprovante de apagamento), igual ao do backend. */
export const MAX_ANEXO_BYTES = 10 * 1024 * 1024

export function anexoGrandeDemais(file: File): boolean {
  return file.size > MAX_ANEXO_BYTES
}

/**
 * Valor em reais digitado no padrão brasileiro. `4.599,90`, `4599,9` e `4599.90`
 * valem 4599.9. Vazio, negativo, não numérico ou com mais de 2 casas → `null`.
 * Sem vírgula, `^\d+\.\d{1,2}$` é decimal; senão os pontos são de milhar.
 */
export function parseMoedaBR(texto: string): number | null {
  const t = texto.trim()
  if (!t || !/^[\d.,]+$/.test(t)) return null

  let inteiro: string
  let centavos = ''
  if (t.includes(',')) {
    const partes = t.split(',')
    if (partes.length !== 2) return null
    inteiro = partes[0].replace(/\./g, '')
    centavos = partes[1]
    if (!/^\d{1,2}$/.test(centavos)) return null
  } else if (/^\d+\.\d{1,2}$/.test(t)) {
    ;[inteiro, centavos] = t.split('.')
  } else if (/^\d{1,3}(\.\d{3})+$/.test(t) || /^\d+$/.test(t)) {
    inteiro = t.replace(/\./g, '')
  } else {
    return null
  }
  if (!/^\d+$/.test(inteiro)) return null
  const n = Number(centavos ? `${inteiro}.${centavos}` : inteiro)
  return Number.isFinite(n) ? n : null
}

const MOEDA_INPUT = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

/** `"4599.90"` → `"4.599,90"` (sem o símbolo, para o campo de digitação). Vazio ou inválido → `''`. */
export function formatMoedaInput(v: string | number | null | undefined): string {
  if (v === null || v === undefined || v === '') return ''
  const n = typeof v === 'number' ? v : Number(v)
  return Number.isFinite(n) ? MOEDA_INPUT.format(n) : ''
}

/** Deixa só dígitos, ponto e vírgula. */
export function maskMoedaInput(v: string): string {
  return v.replace(/[^\d.,]/g, '')
}

/** `AAAA-MM-DD` (ou `AAAA-MM-DDTHH:MM:SS`) → `DD/MM/AAAA`, só por texto. Vazio ou fora do padrão → `''`. */
export function isoParaBR(iso: string | null | undefined): string {
  if (!iso) return ''
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso)
  return m ? `${m[3]}/${m[2]}/${m[1]}` : ''
}

/** Data de calendário válida em `dd/mm/aaaa`. */
export function isDataBR(texto: string): boolean {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(texto.trim())
  if (!m) return false
  const [dia, mes, ano] = [Number(m[1]), Number(m[2]), Number(m[3])]
  const d = new Date(ano, mes - 1, dia)
  return d.getFullYear() === ano && d.getMonth() === mes - 1 && d.getDate() === dia
}

/** Dígitos → `dd/mm/aaaa` (no máximo 8 dígitos), à medida que se digita. */
export function maskDataBR(v: string): string {
  const d = v.replace(/\D/g, '').slice(0, 8)
  if (d.length <= 2) return d
  if (d.length <= 4) return `${d.slice(0, 2)}/${d.slice(2)}`
  return `${d.slice(0, 2)}/${d.slice(2, 4)}/${d.slice(4)}`
}

/** Hoje em `dd/mm/aaaa`, no fuso do navegador. */
export function hojeBR(): string {
  const d = new Date()
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`
}
