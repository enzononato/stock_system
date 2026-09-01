import { type ClassValue, clsx } from 'clsx'
import { twMerge } from 'tailwind-merge'

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs))
}

export function formatDate(dateStr: string | null | undefined): string {
  if (!dateStr) return '-'
  const d = new Date(dateStr)
  if (isNaN(d.getTime())) return dateStr
  return d.toLocaleDateString('pt-BR')
}

export function formatDateTime(dateStr: string | null | undefined): string {
  if (!dateStr) return '-'
  const d = new Date(dateStr)
  if (isNaN(d.getTime())) return dateStr
  return d.toLocaleString('pt-BR')
}

export function formatCpf(cpf: string | null | undefined): string {
  if (!cpf) return '-'
  const digits = cpf.replace(/\D/g, '')
  if (digits.length === 11) {
    return `${digits.slice(0, 3)}.${digits.slice(3, 6)}.${digits.slice(6, 9)}-${digits.slice(9)}`
  }
  return cpf
}

/**
 * Valida um CPF conferindo os dois dígitos verificadores (módulo 11).
 * Rejeita sequências com os 11 dígitos repetidos (ex: 111.111.111-11), que
 * passariam pelo cálculo dos dígitos verificadores mas nunca são CPFs reais.
 */
export function isValidCpf(cpf: string | null | undefined): boolean {
  if (!cpf) return false
  const digits = cpf.replace(/\D/g, '')
  if (digits.length !== 11) return false
  if (/^(\d)\1{10}$/.test(digits)) return false

  const calcCheckDigit = (base: string, startFactor: number): number => {
    let total = 0
    let factor = startFactor
    for (const char of base) {
      total += Number(char) * factor
      factor -= 1
    }
    const remainder = total % 11
    return remainder < 2 ? 0 : 11 - remainder
  }

  const dv1 = calcCheckDigit(digits.slice(0, 9), 10)
  const dv2 = calcCheckDigit(digits.slice(0, 9) + String(dv1), 11)

  return dv1 === Number(digits[9]) && dv2 === Number(digits[10])
}

/** Máscara progressiva de CPF (000.000.000-00) aplicada enquanto o usuário digita. */
export function maskCpfInput(value: string): string {
  return value
    .replace(/\D/g, '')
    .slice(0, 11)
    .replace(/(\d{3})(\d)/, '$1.$2')
    .replace(/(\d{3})(\d)/, '$1.$2')
    .replace(/(\d{3})(\d{1,2})$/, '$1-$2')
}

/** Nota fiscal: exatamente 9 dígitos numéricos (sem formatação). */
export function isValidNotaFiscal(nf: string | null | undefined): boolean {
  if (!nf) return false
  const digits = nf.replace(/\D/g, '')
  return digits.length === 9
}

/** Restringe a digitação da nota fiscal a no máximo 9 dígitos numéricos. */
export function maskNotaFiscalInput(value: string): string {
  return value.replace(/\D/g, '').slice(0, 9)
}

/**
 * Valida um endereço MAC nos formatos AA:BB:CC:DD:EE:FF, AA-BB-CC-DD-EE-FF
 * ou AABBCCDDEEFF (sem separador). Não aceita separadores misturados.
 */
export function isValidMac(mac: string | null | undefined): boolean {
  if (!mac) return false
  const trimmed = mac.trim()
  return (
    /^[0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5}$/.test(trimmed) ||
    /^[0-9A-Fa-f]{2}(-[0-9A-Fa-f]{2}){5}$/.test(trimmed) ||
    /^[0-9A-Fa-f]{12}$/.test(trimmed)
  )
}

/** Máscara progressiva de MAC (AA:BB:CC:DD:EE:FF) aplicada enquanto o usuário digita. */
export function maskMacInput(value: string): string {
  const hex = value.replace(/[^0-9A-Fa-f]/g, '').slice(0, 12).toUpperCase()
  return hex.match(/.{1,2}/g)?.join(':') ?? hex
}

/** Valida um endereço IPv4 (cada octeto entre 0 e 255, sem zeros à esquerda). */
export function isValidIp(ip: string | null | undefined): boolean {
  if (!ip) return false
  const parts = ip.trim().split('.')
  if (parts.length !== 4) return false
  return parts.every((part) => {
    if (!/^\d{1,3}$/.test(part)) return false
    if (part.length > 1 && part.startsWith('0')) return false
    const n = Number(part)
    return n >= 0 && n <= 255
  })
}

/**
 * Exporta um array de objetos como CSV, gerado inteiramente no cliente — usa
 * os dados já carregados na tela (sem chamada extra ao backend). Só serve
 * para listas que já estão 100% no estado do React; para relatórios paginados
 * no servidor, prefira o endpoint de exportação dedicado quando ele existir.
 */
export function exportToCsv<T extends Record<string, unknown>>(
  filename: string,
  rows: T[],
  columns?: { key: keyof T; label: string }[]
): void {
  const first = rows[0]
  if (!first) return
  const cols = columns ?? (Object.keys(first) as (keyof T)[]).map((key) => ({ key, label: String(key) }))

  const escape = (value: unknown) => {
    const s = value == null ? '' : String(value)
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
  }

  const header = cols.map((c) => escape(c.label)).join(',')
  const body = rows.map((row) => cols.map((c) => escape(row[c.key])).join(',')).join('\n')
  const csv = '\uFEFF' + header + '\n' + body // BOM para acentuação abrir certo no Excel

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename.endsWith('.csv') ? filename : `${filename}.csv`
  document.body.appendChild(a)
  a.click()
  a.remove()
  URL.revokeObjectURL(url)
}
