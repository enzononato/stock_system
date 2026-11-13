import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

const mocks = vi.hoisted(() => ({
  construtor: vi.fn(),
  decodeFromConstraints: vi.fn(),
}))

vi.mock('@zxing/browser', () => ({
  BrowserQRCodeReader: class {
    constructor(...args: unknown[]) {
      mocks.construtor(...args)
    }
    decodeFromConstraints = mocks.decodeFromConstraints
  },
}))

import { ErroCamera, MENSAGEM_ERRO_CAMERA, iniciarLeitor } from '@/features/inventario/scanner'

function definir(alvo: object, chave: string, valor: unknown) {
  Object.defineProperty(alvo, chave, { value: valor, configurable: true, writable: true })
}

beforeEach(() => {
  mocks.construtor.mockReset()
  mocks.decodeFromConstraints.mockReset()
  definir(window, 'isSecureContext', true)
  definir(navigator, 'mediaDevices', { getUserMedia: vi.fn() })
})

afterEach(() => vi.restoreAllMocks())

describe('iniciarLeitor', () => {
  it('fora de HTTPS recusa sem nem carregar o leitor', async () => {
    definir(window, 'isSecureContext', false)

    await expect(iniciarLeitor(document.createElement('video'), vi.fn())).rejects.toMatchObject({
      tipo: 'inseguro',
      message: MENSAGEM_ERRO_CAMERA.inseguro,
    })
    expect(mocks.construtor).not.toHaveBeenCalled()
  })

  it('aparelho sem API de câmera vira "sem_camera"', async () => {
    definir(navigator, 'mediaDevices', undefined)
    await expect(iniciarLeitor(document.createElement('video'), vi.fn())).rejects.toMatchObject({
      tipo: 'sem_camera',
    })
  })

  it('traduz os erros do navegador pelo nome', async () => {
    const casos: [string, string][] = [
      ['NotAllowedError', 'negada'],
      ['SecurityError', 'negada'],
      ['NotFoundError', 'sem_camera'],
      ['OverconstrainedError', 'sem_camera'],
      ['AbortError', 'falha'],
    ]
    for (const [nome, tipo] of casos) {
      mocks.decodeFromConstraints.mockRejectedValueOnce(new DOMException('x', nome))
      await expect(iniciarLeitor(document.createElement('video'), vi.fn())).rejects.toBeInstanceOf(ErroCamera)
      mocks.decodeFromConstraints.mockRejectedValueOnce(new DOMException('x', nome))
      await expect(iniciarLeitor(document.createElement('video'), vi.fn())).rejects.toMatchObject({ tipo })
    }
  })

  it('pede a câmera traseira e entrega o texto lido', async () => {
    const onCodigo = vi.fn()
    let callback: (r?: { getText: () => string }) => void = () => {}
    mocks.decodeFromConstraints.mockImplementation(async (_c, _v, cb) => {
      callback = cb
      return { stop: vi.fn() }
    })
    const video = document.createElement('video')

    await iniciarLeitor(video, onCodigo)

    expect(mocks.construtor).toHaveBeenCalledWith(undefined, { delayBetweenScanAttempts: 150 })
    expect(mocks.decodeFromConstraints).toHaveBeenCalledWith(
      { audio: false, video: { facingMode: { ideal: 'environment' } } },
      video,
      expect.any(Function)
    )
    callback(undefined)
    expect(onCodigo).not.toHaveBeenCalled()
    callback({ getText: () => 'https://x/e/1' })
    expect(onCodigo).toHaveBeenCalledWith('https://x/e/1')
  })

  it('parar() desliga o leitor e cada trilha da câmera, e pode ser chamado duas vezes', async () => {
    const stop = vi.fn()
    mocks.decodeFromConstraints.mockResolvedValue({ stop })
    const trilha = { stop: vi.fn() }
    const video = document.createElement('video')
    definir(video, 'srcObject', { getTracks: () => [trilha] })

    const controles = await iniciarLeitor(video, vi.fn())
    controles.parar()
    controles.parar()

    expect(stop).toHaveBeenCalled()
    expect(trilha.stop).toHaveBeenCalled()
    expect(video.srcObject).toBeNull()
  })

  it('se a abertura falha depois de o fluxo ser ligado ao vídeo, solta a câmera', async () => {
    const trilha = { stop: vi.fn() }
    const video = document.createElement('video')
    definir(video, 'srcObject', { getTracks: () => [trilha] })
    mocks.decodeFromConstraints.mockRejectedValue(new DOMException('x', 'NotAllowedError'))

    await expect(iniciarLeitor(video, vi.fn())).rejects.toMatchObject({ tipo: 'negada' })

    expect(trilha.stop).toHaveBeenCalled()
    expect(video.srcObject).toBeNull()
  })
})
