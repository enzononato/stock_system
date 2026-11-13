export type TipoErroCamera = 'inseguro' | 'negada' | 'sem_camera' | 'falha'

export const MENSAGEM_ERRO_CAMERA: Record<TipoErroCamera, string> = {
  inseguro: 'A câmera só funciona em conexão segura (HTTPS). Use a digitação ou um leitor USB.',
  negada: 'Permissão da câmera negada. Use a digitação abaixo ou libere a câmera nas configurações do navegador.',
  sem_camera: 'Nenhuma câmera disponível neste aparelho. Use a digitação ou um leitor USB.',
  falha: 'Não foi possível abrir a câmera. Use a digitação ou um leitor USB.',
}

/** Erro ao abrir a câmera, já com a mensagem para o usuário em `message`. */
export class ErroCamera extends Error {
  constructor(readonly tipo: TipoErroCamera) {
    super(MENSAGEM_ERRO_CAMERA[tipo])
    this.name = 'ErroCamera'
  }
}

export interface ControlesLeitor {
  /** Desliga a câmera. Pode ser chamado mais de uma vez. */
  parar(): void
}

function tipoDoErro(e: unknown): TipoErroCamera {
  const nome = (e as { name?: string } | null)?.name
  if (nome === 'NotAllowedError' || nome === 'SecurityError') return 'negada'
  if (nome === 'NotFoundError' || nome === 'OverconstrainedError') return 'sem_camera'
  return 'falha'
}

function pararFluxo(video: HTMLVideoElement): void {
  const fluxo = video.srcObject as MediaStream | null
  fluxo?.getTracks().forEach((trilha) => trilha.stop())
  video.srcObject = null
}

/**
 * Liga a câmera traseira e lê QR codes no `video`, chamando `onCodigo` com o texto
 * de cada leitura. O `@zxing/browser` só é carregado aqui, sob demanda. A câmera só
 * funciona em contexto seguro (HTTPS ou localhost).
 */
export async function iniciarLeitor(
  video: HTMLVideoElement,
  onCodigo: (texto: string) => void
): Promise<ControlesLeitor> {
  if (!window.isSecureContext) throw new ErroCamera('inseguro')
  if (!navigator.mediaDevices?.getUserMedia) throw new ErroCamera('sem_camera')

  let controles: { stop: () => void }
  try {
    const { BrowserQRCodeReader } = await import('@zxing/browser')
    const leitor = new BrowserQRCodeReader(undefined, { delayBetweenScanAttempts: 150 })
    controles = await leitor.decodeFromConstraints(
      { audio: false, video: { facingMode: { ideal: 'environment' } } },
      video,
      (resultado) => {
        if (resultado) onCodigo(resultado.getText())
      }
    )
  } catch (e) {
    // O fluxo pode já ter sido ligado ao vídeo antes da falha: solta a câmera.
    pararFluxo(video)
    throw e instanceof ErroCamera ? e : new ErroCamera(tipoDoErro(e))
  }

  return {
    parar() {
      try {
        controles.stop()
      } catch {
        /* já parado */
      }
      // Garante que o indicador de câmera do navegador apague, mesmo se o leitor não soltou as trilhas.
      pararFluxo(video)
    },
  }
}
