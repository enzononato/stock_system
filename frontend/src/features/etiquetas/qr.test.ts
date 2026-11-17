import { describe, expect, it } from 'vitest'
import { gerarQrSvg, gerarQrs } from '@/features/etiquetas/qr'

describe('qr', () => {
  it('gera um SVG vetorial para o endereço', async () => {
    const svg = await gerarQrSvg('https://patrimonio.revalle.com.br/e/12')
    expect(svg.startsWith('<svg')).toBe(true)
    expect(svg).toContain('viewBox')
  })

  it('gera um SVG por etiqueta, indexado pelo id', async () => {
    const qrs = await gerarQrs([
      { id: 1, url: 'https://x/e/1' },
      { id: 2, url: 'https://x/e/2' },
    ])
    expect(Object.keys(qrs)).toEqual(['1', '2'])
    expect(qrs[1]).not.toBe(qrs[2])
  })
})
