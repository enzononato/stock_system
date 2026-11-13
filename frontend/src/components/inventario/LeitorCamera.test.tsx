import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, render, screen } from '@testing-library/react'

const mocks = vi.hoisted(() => ({ iniciarLeitor: vi.fn() }))

vi.mock('@/features/inventario/scanner', async (original) => ({
  ...(await original<typeof import('@/features/inventario/scanner')>()),
  iniciarLeitor: mocks.iniciarLeitor,
}))

import { LeitorCamera } from '@/components/inventario/LeitorCamera'
import { ErroCamera } from '@/features/inventario/scanner'

function esconderAba() {
  Object.defineProperty(document, 'visibilityState', { value: 'hidden', configurable: true })
  document.dispatchEvent(new Event('visibilitychange'))
}

beforeEach(() => {
  mocks.iniciarLeitor.mockReset()
  Object.defineProperty(document, 'visibilityState', { value: 'visible', configurable: true })
})

afterEach(() => vi.restoreAllMocks())

function props() {
  return { onCodigo: vi.fn(), onErro: vi.fn(), onPausar: vi.fn() }
}

describe('LeitorCamera', () => {
  it('renderiza o vídeo acessível e inicia a câmera ao montar', () => {
    mocks.iniciarLeitor.mockResolvedValue({ parar: vi.fn() })
    render(<LeitorCamera {...props()} />)

    expect(screen.getByLabelText('Imagem da câmera')).toBeInTheDocument()
    expect(mocks.iniciarLeitor).toHaveBeenCalledTimes(1)
  })

  it('desmontado antes de o início terminar, desliga assim que termina', async () => {
    const parar = vi.fn()
    let resolver: (c: { parar: () => void }) => void = () => {}
    mocks.iniciarLeitor.mockReturnValue(new Promise((r) => (resolver = r)))
    const { unmount } = render(<LeitorCamera {...props()} />)

    unmount()
    expect(parar).not.toHaveBeenCalled()
    await act(async () => resolver({ parar }))

    expect(parar).toHaveBeenCalledTimes(1)
  })

  it('desliga a câmera ao desmontar', async () => {
    const parar = vi.fn()
    mocks.iniciarLeitor.mockResolvedValue({ parar })
    const { unmount } = render(<LeitorCamera {...props()} />)
    await act(async () => {})

    unmount()

    expect(parar).toHaveBeenCalledTimes(1)
  })

  it('aba escondida desliga a câmera e avisa', async () => {
    const parar = vi.fn()
    const p = props()
    mocks.iniciarLeitor.mockResolvedValue({ parar })
    render(<LeitorCamera {...p} />)
    await act(async () => {})

    act(() => esconderAba())

    expect(parar).toHaveBeenCalledTimes(1)
    expect(p.onPausar).toHaveBeenCalledTimes(1)
  })

  it('aba escondida enquanto a câmera ainda abre: desliga quando ela abrir', async () => {
    const parar = vi.fn()
    let resolver: (c: { parar: () => void }) => void = () => {}
    mocks.iniciarLeitor.mockReturnValue(new Promise((r) => (resolver = r)))
    const p = props()
    render(<LeitorCamera {...p} />)

    act(() => esconderAba())
    await act(async () => resolver({ parar }))

    expect(p.onPausar).toHaveBeenCalledTimes(1)
    expect(parar).toHaveBeenCalledTimes(1)
  })

  it('o mesmo código em menos de 3 s conta uma vez; depois de 3 s, de novo', async () => {
    const p = props()
    let entregar: (t: string) => void = () => {}
    mocks.iniciarLeitor.mockImplementation(async (_video, cb) => {
      entregar = cb
      return { parar: vi.fn() }
    })
    const agora = vi.spyOn(Date, 'now').mockReturnValue(1_000_000)
    render(<LeitorCamera {...p} />)
    await act(async () => {})

    entregar('https://x/e/1')
    entregar('https://x/e/1')
    expect(p.onCodigo).toHaveBeenCalledTimes(1)

    agora.mockReturnValue(1_000_000 + 3001)
    entregar('https://x/e/1')
    expect(p.onCodigo).toHaveBeenCalledTimes(2)

    // Outro código passa na hora.
    entregar('https://x/e/2')
    expect(p.onCodigo).toHaveBeenCalledTimes(3)
    expect(p.onCodigo).toHaveBeenLastCalledWith('https://x/e/2')
  })

  it('erro de permissão chega ao pai com o tipo certo', async () => {
    const p = props()
    mocks.iniciarLeitor.mockRejectedValue(new ErroCamera('negada'))
    render(<LeitorCamera {...p} />)
    await act(async () => {})

    expect(p.onErro).toHaveBeenCalledTimes(1)
    expect(p.onErro.mock.calls[0][0]).toMatchObject({ tipo: 'negada' })
  })

  it('erro qualquer vira "falha"', async () => {
    const p = props()
    mocks.iniciarLeitor.mockRejectedValue(new Error('boom'))
    render(<LeitorCamera {...p} />)
    await act(async () => {})

    expect(p.onErro.mock.calls[0][0]).toMatchObject({ tipo: 'falha' })
  })
})
