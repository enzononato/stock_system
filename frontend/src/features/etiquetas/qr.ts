import type { Etiqueta } from '@/api/etiquetas'

/**
 * QR code em SVG (vetorial: fica nítido em qualquer resolução de impressora).
 * A biblioteca é carregada sob demanda, só na tela de etiquetas; nunca importe
 * `qrcode` de forma estática (o mesmo `import()` dinâmico vale para o 2FA).
 * As cores ficam no padrão da biblioteca: preto sobre branco, que é o que o
 * papel e os leitores esperam nos dois temas.
 */
export async function gerarQrSvg(conteudo: string): Promise<string> {
  const mod = await import('qrcode')
  const QRCode = (mod as unknown as { default?: typeof mod }).default ?? mod
  return QRCode.toString(conteudo, { type: 'svg', errorCorrectionLevel: 'M', margin: 1 })
}

/** SVG do QR de cada etiqueta, indexado pelo id do equipamento. */
export async function gerarQrs(etiquetas: Etiqueta[]): Promise<Record<number, string>> {
  const pares = await Promise.all(
    etiquetas.map(async (e) => [e.id, await gerarQrSvg(e.url)] as const)
  )
  return Object.fromEntries(pares)
}
