import { describe, expect, it } from 'vitest'
import {
  anexoGrandeDemais,
  formatMoedaInput,
  hojeBR,
  isDataBR,
  isoParaBR,
  maskDataBR,
  maskMoedaInput,
  MAX_ANEXO_BYTES,
  parseMoedaBR,
} from './format'

describe('parseMoedaBR', () => {
  it.each(['4.599,90', '4599,9', '4599.90'])('%s vale 4599.9', (texto) => {
    expect(parseMoedaBR(texto)).toBe(4599.9)
  })

  it('pontos de milhar sem vírgula', () => {
    expect(parseMoedaBR('1.234')).toBe(1234)
    expect(parseMoedaBR('1.234.567')).toBe(1234567)
    expect(parseMoedaBR('150')).toBe(150)
    expect(parseMoedaBR('0,5')).toBe(0.5)
  })

  it.each(['', '   ', 'abc', '-5', '1,234', '1.2345', '1,5,5', 'R$ 10'])('%j é recusado', (texto) => {
    expect(parseMoedaBR(texto)).toBeNull()
  })
})

describe('formatMoedaInput', () => {
  it('formata no padrão brasileiro', () => {
    expect(formatMoedaInput('4599.90')).toBe('4.599,90')
    expect(formatMoedaInput(150)).toBe('150,00')
  })
  it('vazio vira vazio', () => {
    expect(formatMoedaInput(null)).toBe('')
    expect(formatMoedaInput(undefined)).toBe('')
    expect(formatMoedaInput('')).toBe('')
  })
})

describe('maskMoedaInput', () => {
  it('aceita só dígitos, ponto e vírgula', () => {
    expect(maskMoedaInput('R$ 4.599,90abc')).toBe('4.599,90')
  })
})

describe('isoParaBR', () => {
  it('converte por texto, sem deslocar o dia', () => {
    expect(isoParaBR('2026-01-10')).toBe('10/01/2026')
    expect(isoParaBR('2026-01-10T00:00:00')).toBe('10/01/2026')
  })
  it('vazio vira vazio', () => {
    expect(isoParaBR(null)).toBe('')
    expect(isoParaBR(undefined)).toBe('')
    expect(isoParaBR('')).toBe('')
  })
})

describe('datas dd/mm/aaaa', () => {
  it('isDataBR confere o calendário', () => {
    expect(isDataBR('31/02/2026')).toBe(false)
    expect(isDataBR('29/02/2028')).toBe(true)
    expect(isDataBR('29/02/2026')).toBe(false)
    expect(isDataBR('10/01/26')).toBe(false)
    expect(isDataBR('')).toBe(false)
  })
  it('maskDataBR coloca as barras', () => {
    expect(maskDataBR('10012026')).toBe('10/01/2026')
    expect(maskDataBR('1001')).toBe('10/01')
    expect(maskDataBR('1')).toBe('1')
    expect(maskDataBR('10/01/2026999')).toBe('10/01/2026')
  })
  it('hojeBR devolve a data de hoje', () => {
    expect(isDataBR(hojeBR())).toBe(true)
    const d = new Date()
    expect(hojeBR()).toBe(
      `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`
    )
  })
})

describe('anexoGrandeDemais', () => {
  it('limite de 10 MB', () => {
    const arquivo = (size: number) => {
      const f = new File(['x'], 'nf.pdf')
      Object.defineProperty(f, 'size', { value: size })
      return f
    }
    expect(anexoGrandeDemais(arquivo(MAX_ANEXO_BYTES))).toBe(false)
    expect(anexoGrandeDemais(arquivo(MAX_ANEXO_BYTES + 1))).toBe(true)
  })
})
