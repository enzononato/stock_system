import { useEffect, useRef } from 'react'
import { ErroCamera, iniciarLeitor, type ControlesLeitor } from '@/features/inventario/scanner'

interface LeitorCameraProps {
  /** `false` = a tela não aceitou o código (ocupada): ele não entra na janela de repetição e será entregue de novo. */
  onCodigo: (texto: string) => boolean | void
  onErro: (erro: ErroCamera) => void
  /** A aba foi escondida e a câmera foi desligada: o pai avisa o usuário. */
  onPausar: () => void
}

/** Mesmo texto repetido dentro desta janela conta uma vez (a câmera relê o mesmo QR a cada quadro). */
const JANELA_REPETICAO_MS = 3000

/**
 * Vídeo da câmera com leitura de QR. Liga ao montar e desliga ao desmontar, ao
 * esconder a aba e quando o início termina depois de a tela já ter saído.
 */
export function LeitorCamera({ onCodigo, onErro, onPausar }: LeitorCameraProps) {
  const videoRef = useRef<HTMLVideoElement>(null)
  const funcoes = useRef({ onCodigo, onErro, onPausar })
  funcoes.current = { onCodigo, onErro, onPausar }

  useEffect(() => {
    const video = videoRef.current
    if (!video) return

    let encerrado = false
    let controles: ControlesLeitor | null = null
    let ultimo = { texto: '', em: 0 }

    function parar() {
      const atuais = controles
      controles = null
      atuais?.parar()
    }

    iniciarLeitor(video, (texto) => {
      const agora = Date.now()
      const repetido = texto === ultimo.texto && agora - ultimo.em < JANELA_REPETICAO_MS
      if (repetido) {
        // Continuar enxergando o mesmo QR renova a janela: segurar a câmera parada não repete a leitura.
        ultimo = { texto, em: agora }
        return
      }
      // Só o que a tela aceitou entra na janela: um código recusado (tela ocupada) segue
      // à vista e é entregue de novo no quadro seguinte.
      if (funcoes.current.onCodigo(texto) !== false) ultimo = { texto, em: agora }
    }).then(
      (c) => {
        // `encerrado` vale para desmontagem e para aba escondida: nos dois casos a câmera não deve ficar ligada.
        if (encerrado) c.parar()
        else controles = c
      },
      (e: unknown) => {
        if (encerrado) return
        funcoes.current.onErro(e instanceof ErroCamera ? e : new ErroCamera('falha'))
      }
    )

    function aoMudarVisibilidade() {
      if (document.visibilityState !== 'hidden') return
      encerrado = true
      parar()
      funcoes.current.onPausar()
    }
    document.addEventListener('visibilitychange', aoMudarVisibilidade)

    return () => {
      encerrado = true
      document.removeEventListener('visibilitychange', aoMudarVisibilidade)
      parar()
    }
  }, [])

  return (
    <video
      ref={videoRef}
      muted
      playsInline
      autoPlay
      aria-label="Imagem da câmera"
      className="aspect-square w-full max-w-sm rounded-md bg-surface-alt object-cover"
    />
  )
}
