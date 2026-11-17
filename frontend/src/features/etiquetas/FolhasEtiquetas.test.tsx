import { describe, expect, it } from 'vitest'
import { render, screen } from '@testing-library/react'
import { FolhasEtiquetas } from '@/features/etiquetas/FolhasEtiquetas'
import { LAYOUT_PADRAO } from '@/features/etiquetas/layout'
import type { Etiqueta } from '@/api/etiquetas'

function etiquetas(n: number): Etiqueta[] {
  return Array.from({ length: n }, (_, i) => ({
    id: i + 12,
    tipo: 'Notebook',
    brand: 'Dell',
    model: 'Latitude 5420',
    identificador: `SN-${i}`,
    revenda: 'Juazeiro',
    url: `https://x/e/${i + 12}`,
  }))
}

const qrs = (lista: Etiqueta[]) =>
  Object.fromEntries(lista.map((e) => [e.id, '<svg viewBox="0 0 10 10"><rect width="10" height="10"/></svg>']))

describe('FolhasEtiquetas', () => {
  it('26 etiquetas no layout padrão ocupam 2 folhas A4', () => {
    const lista = etiquetas(26)
    render(<FolhasEtiquetas etiquetas={lista} qrs={qrs(lista)} layout={LAYOUT_PADRAO} posicaoInicial={1} />)

    const folhas = screen.getAllByTestId('folha')
    expect(folhas).toHaveLength(2)
    expect(folhas[0].style.width).toBe('210mm')
    expect(folhas[0].style.height).toBe('297mm')
    expect(screen.getAllByTestId('etiqueta')).toHaveLength(26)
  })

  it('a posição inicial desloca a primeira etiqueta', () => {
    const lista = etiquetas(3)
    render(<FolhasEtiquetas etiquetas={lista} qrs={qrs(lista)} layout={LAYOUT_PADRAO} posicaoInicial={5} />)

    const primeira = screen.getAllByTestId('etiqueta')[0]
    expect(primeira.style.left).toBe('70mm')
    expect(primeira.style.top).toBe('37.5mm')
    expect(primeira.style.width).toBe('70mm')
    expect(primeira.style.height).toBe('37mm')
  })

  it('mostra patrimônio, tipo, marca e modelo, com o QR acessível', () => {
    const lista = etiquetas(1)
    render(<FolhasEtiquetas etiquetas={lista} qrs={qrs(lista)} layout={LAYOUT_PADRAO} posicaoInicial={1} />)

    expect(screen.getByText('#12')).toBeInTheDocument()
    expect(screen.getByText('Notebook')).toBeInTheDocument()
    expect(screen.getByText('Dell Latitude 5420')).toBeInTheDocument()
    expect(screen.getByLabelText('QR do patrimônio #12')).toBeInTheDocument()
  })
})
