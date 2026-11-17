/** Converte um canal sRGB (0-255) em canal linear, conforme WCAG 2.x. */
function canalLinear(valor: number): number {
  const c = valor / 255
  return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4
}

/** Luminância relativa de uma cor hexadecimal `#RRGGBB` (WCAG 2.x). */
function luminanciaRelativa(hex: string): number {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim())
  if (!m) throw new Error(`Cor hexadecimal inválida: ${hex}`)
  const n = parseInt(m[1], 16)
  const r = canalLinear((n >> 16) & 0xff)
  const g = canalLinear((n >> 8) & 0xff)
  const b = canalLinear(n & 0xff)
  return 0.2126 * r + 0.7152 * g + 0.0722 * b
}

/** Razão de contraste WCAG 2.x entre duas cores `#RRGGBB` (de 1 a 21, independe da ordem). */
export function contrastRatio(hexA: string, hexB: string): number {
  const a = luminanciaRelativa(hexA)
  const b = luminanciaRelativa(hexB)
  const [claro, escuro] = a >= b ? [a, b] : [b, a]
  return (claro + 0.05) / (escuro + 0.05)
}
