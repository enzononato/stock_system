import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  CHAVE_LAYOUT,
  LAYOUT_PADRAO,
  calcularPosicoes,
  carregarLayout,
  etiquetasPorFolha,
  salvarLayout,
  totalFolhas,
  validarLayout,
  validarPosicaoInicial,
} from '@/features/etiquetas/layout'

beforeEach(() => localStorage.clear())
afterEach(() => vi.restoreAllMocks())

describe('validarLayout', () => {
  it('o padrão (3 × 8, 70 × 37 mm) é válido e dá 24 por folha', () => {
    expect(validarLayout(LAYOUT_PADRAO)).toBeNull()
    expect(etiquetasPorFolha(LAYOUT_PADRAO)).toBe(24)
  })

  it('layout que não cabe na A4 é recusado', () => {
    expect(validarLayout({ ...LAYOUT_PADRAO, colunas: 4 })).toBe('O layout não cabe numa folha A4 (210 × 297 mm).')
    expect(validarLayout({ ...LAYOUT_PADRAO, linhas: 9 })).toBe('O layout não cabe numa folha A4 (210 × 297 mm).')
    expect(validarLayout({ ...LAYOUT_PADRAO, margemEsquerdaMm: 5 })).toBe(
      'O layout não cabe numa folha A4 (210 × 297 mm).'
    )
  })

  it('devolve a primeira falha na ordem dos campos', () => {
    expect(validarLayout({ ...LAYOUT_PADRAO, colunas: 0 })).toBe('Colunas: informe um número inteiro de 1 a 10.')
    expect(validarLayout({ ...LAYOUT_PADRAO, colunas: 1.5 })).toBe('Colunas: informe um número inteiro de 1 a 10.')
    expect(validarLayout({ ...LAYOUT_PADRAO, linhas: 31 })).toBe('Linhas: informe um número inteiro de 1 a 30.')
    expect(validarLayout({ ...LAYOUT_PADRAO, larguraMm: 9 })).toBe('Largura: informe de 10 a 210 mm.')
    expect(validarLayout({ ...LAYOUT_PADRAO, alturaMm: 298 })).toBe('Altura: informe de 10 a 297 mm.')
    expect(validarLayout({ ...LAYOUT_PADRAO, margemSuperiorMm: -1 })).toBe(
      'Margens e espaçamentos: informe de 0 a 50 mm.'
    )
    expect(validarLayout({ ...LAYOUT_PADRAO, espacoVerticalMm: 51 })).toBe(
      'Margens e espaçamentos: informe de 0 a 50 mm.'
    )
    expect(validarLayout({ ...LAYOUT_PADRAO, larguraMm: Number.NaN })).toBe('Largura: informe de 10 a 210 mm.')
  })
})

describe('posições', () => {
  it('posição inicial precisa estar entre 1 e o total da folha', () => {
    expect(validarPosicaoInicial(LAYOUT_PADRAO, 25)).toBe('A posição inicial deve estar entre 1 e 24.')
    expect(validarPosicaoInicial(LAYOUT_PADRAO, 0)).toBe('A posição inicial deve estar entre 1 e 24.')
    expect(validarPosicaoInicial(LAYOUT_PADRAO, 1.5)).toBe('A posição inicial deve estar entre 1 e 24.')
    expect(validarPosicaoInicial(LAYOUT_PADRAO, 24)).toBeNull()
  })

  it('preenche linha a linha, da esquerda para a direita, a partir da posição inicial', () => {
    const [a, b] = calcularPosicoes(LAYOUT_PADRAO, 2, 5)
    expect(a).toMatchObject({ folha: 0, linha: 1, coluna: 1 })
    expect(a.esquerdaMm).toBeCloseTo(70)
    expect(a.topoMm).toBeCloseTo(37.5)
    expect(b).toMatchObject({ folha: 0, linha: 1, coluna: 2 })
    expect(b.esquerdaMm).toBeCloseTo(140)
    expect(b.topoMm).toBeCloseTo(37.5)
  })

  it('a 25ª etiqueta cai na primeira célula da segunda folha', () => {
    const posicoes = calcularPosicoes(LAYOUT_PADRAO, 25, 1)
    expect(posicoes).toHaveLength(25)
    expect(posicoes[24]).toMatchObject({ folha: 1, linha: 0, coluna: 0 })
  })

  it('conta as folhas considerando a posição inicial', () => {
    expect(totalFolhas(LAYOUT_PADRAO, 20, 10)).toBe(2)
    expect(totalFolhas(LAYOUT_PADRAO, 24, 1)).toBe(1)
    expect(totalFolhas(LAYOUT_PADRAO, 24, 2)).toBe(2)
    expect(totalFolhas(LAYOUT_PADRAO, 0, 1)).toBe(0)
  })

  it('respeita margens e espaços entre as células', () => {
    const layout = { ...LAYOUT_PADRAO, colunas: 2, larguraMm: 50, margemEsquerdaMm: 10, espacoHorizontalMm: 5 }
    const [, segunda] = calcularPosicoes(layout, 2, 1)
    expect(segunda.esquerdaMm).toBeCloseTo(10 + 55)
  })
})

describe('persistência do layout', () => {
  it('JSON inválido, campo faltando ou layout inválido voltam ao padrão', () => {
    localStorage.setItem(CHAVE_LAYOUT, '{x')
    expect(carregarLayout()).toEqual(LAYOUT_PADRAO)
    localStorage.setItem(CHAVE_LAYOUT, JSON.stringify({ colunas: 2 }))
    expect(carregarLayout()).toEqual(LAYOUT_PADRAO)
    localStorage.setItem(CHAVE_LAYOUT, JSON.stringify({ ...LAYOUT_PADRAO, colunas: 7 }))
    expect(carregarLayout()).toEqual(LAYOUT_PADRAO)
  })

  it('layout válido salvo é restaurado', () => {
    salvarLayout({ ...LAYOUT_PADRAO, colunas: 2, larguraMm: 100 })
    expect(carregarLayout().colunas).toBe(2)
  })

  it('armazenamento indisponível não quebra', () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new Error('bloqueado')
    })
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new Error('bloqueado')
    })
    expect(carregarLayout()).toEqual(LAYOUT_PADRAO)
    expect(() => salvarLayout(LAYOUT_PADRAO)).not.toThrow()
  })
})
