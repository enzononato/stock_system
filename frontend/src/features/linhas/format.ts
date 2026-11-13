/**
 * Custo digitado no padrão brasileiro (`49,90`, `1.234,56`, `R$ 49,9`) ou com ponto
 * decimal (`49.9`) para o texto com 2 casas que o backend espera (`"49.90"`).
 * Devolve `null` quando não for um valor monetário válido (negativo, mais de 2
 * casas, `1.234` sem vírgula — ambíguo —, texto).
 */
export function parseCusto(texto: string): string | null {
  let valor = texto.replace(/\s/g, '').replace(/R\$/gi, '')
  if (valor.includes(',')) valor = valor.replace(/\./g, '').replace(',', '.')
  if (!/^\d+(\.\d{1,2})?$/.test(valor)) return null
  const [inteiro, decimais = ''] = valor.split('.')
  return `${inteiro}.${decimais.padEnd(2, '0')}`
}

/** `"49.90"` → `"49,90"` para preencher o campo de edição. */
export function custoParaInput(v: string | number | null | undefined): string {
  if (v === null || v === undefined || v === '') return ''
  return String(v).replace('.', ',')
}

/** `Celular Samsung A15 · CEL-A` — mesmo formato dos eventos do backend. */
export function descreverAparelho(
  a:
    | { tipo?: string | null; brand?: string | null; model?: string | null; identificador?: string | null }
    | null
    | undefined
): string {
  if (!a) return ''
  const base = [a.tipo, a.brand, a.model].filter(Boolean).join(' ')
  if (a.identificador) return base ? `${base} · ${a.identificador}` : a.identificador
  return base
}
