import { describe, it, expect } from 'vitest'
import { upper, formatTelefone, formatMoeda, formatDataBR } from './format'
import { formatDate } from './utils'

const norm = (s: string) => s.replace(/\s/g, ' ')

describe('upper', () => {
  it('converte para maiúsculas em pt-BR', () => {
    expect(upper('açaí')).toBe('AÇAÍ')
    expect(upper('i')).toBe('I')
  })
  it('vazio vira string vazia', () => {
    expect(upper(null)).toBe('')
    expect(upper(undefined)).toBe('')
    expect(upper('')).toBe('')
  })
})

describe('formatTelefone', () => {
  it('celular', () => {
    expect(formatTelefone('74999990000')).toBe('(74) 99999-0000')
    expect(formatTelefone('(74) 99999-0000')).toBe('(74) 99999-0000')
  })
  it('fixo', () => {
    expect(formatTelefone('7436110000')).toBe('(74) 3611-0000')
  })
  it('vazio ou formato desconhecido', () => {
    expect(formatTelefone(null)).toBe('')
    expect(formatTelefone('')).toBe('')
    expect(formatTelefone('123')).toBe('123')
  })
})

describe('formatMoeda', () => {
  it('formata em reais', () => {
    expect(norm(formatMoeda(1234.56))).toBe('R$ 1.234,56')
    expect(norm(formatMoeda('1234.5'))).toBe('R$ 1.234,50')
    expect(norm(formatMoeda(0))).toBe('R$ 0,00')
  })
  it('vazio ou inválido', () => {
    expect(formatMoeda(null)).toBe('')
    expect(formatMoeda(undefined)).toBe('')
    expect(formatMoeda('')).toBe('')
    expect(formatMoeda('abc')).toBe('')
  })
})

describe('formatDataBR', () => {
  it('é o formatDate de utils', () => {
    expect(formatDataBR).toBe(formatDate)
    expect(formatDataBR(null)).toBe('-')
  })
})
