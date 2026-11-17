import { useEffect, useState } from 'react'

interface QrCodeOtpauthProps {
  uri: string
  size?: number
}

/**
 * QR code do `otpauth://` para o aplicativo autenticador. A biblioteca `qrcode`
 * é carregada sob demanda (fica fora do bundle principal) e o SVG vira uma
 * imagem `data:` — sem `dangerouslySetInnerHTML`. Fundo branco fixo: o código
 * precisa de contraste nos dois temas.
 */
export function QrCodeOtpauth({ uri, size = 192 }: QrCodeOtpauthProps) {
  const [src, setSrc] = useState<string | null>(null)

  useEffect(() => {
    let cancelado = false
    setSrc(null)
    void (async () => {
      try {
        const mod = await import('qrcode')
        // CJS carregado por import(): conforme o bundler, as funções vêm em `default` ou no namespace.
        const QRCode = mod.default ?? mod
        const svg = await QRCode.toString(uri, { type: 'svg', margin: 1 })
        if (!cancelado) setSrc('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg))
      } catch {
        // sem QR o usuário ainda tem a chave em texto (CadastroAutenticador)
      }
    })()
    return () => {
      cancelado = true
    }
  }, [uri])

  return (
    <div
      className="inline-flex items-center justify-center rounded-md bg-white p-2"
      style={{ width: size + 16, height: size + 16 }}
    >
      {src ? (
        <img src={src} alt="QR code para o aplicativo autenticador" width={size} height={size} />
      ) : (
        <div
          aria-busy="true"
          className="animate-pulse rounded bg-neutral-200"
          style={{ width: size, height: size }}
        />
      )}
    </div>
  )
}
