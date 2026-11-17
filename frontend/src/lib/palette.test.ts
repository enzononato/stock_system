import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { contrastRatio } from './contrast'

const css = readFileSync(resolve(__dirname, '../index.css'), 'utf8')

/** Extrai as declarações `--token: valor;` de um bloco `seletor { ... }` (sem chaves aninhadas). */
function lerBloco(seletor: RegExp): Record<string, string> {
  const m = css.match(seletor)
  if (!m) throw new Error(`Bloco não encontrado em index.css: ${seletor}`)
  const tokens: Record<string, string> = {}
  for (const d of m[1].matchAll(/(--[\w-]+)\s*:\s*([^;]+);/g)) {
    tokens[d[1]] = d[2].trim()
  }
  return tokens
}

const temas: Record<string, Record<string, string>> = {
  claro: lerBloco(/:root\s*\{([^}]*)\}/),
  escuro: lerBloco(/\.dark\s*\{([^}]*)\}/),
}

const tonsDeStatus = ['available', 'loaned', 'pending', 'return', 'maintenance', 'retired', 'danger']

const pares: [string, string][] = [
  ['--primary-foreground', '--primary'],
  ['--brand-text', '--surface'],
  ['--brand-text', '--canvas'],
  ['--brand-soft-foreground', '--brand-soft'],
  ['--lime-foreground', '--lime'],
  ...tonsDeStatus.map((t): [string, string] => [`--status-${t}`, `--status-${t}-bg`]),
]

describe('contraste da paleta (WCAG AA, 4.5:1)', () => {
  for (const [tema, tokens] of Object.entries(temas)) {
    describe(`tema ${tema}`, () => {
      for (const [fg, bg] of pares) {
        it(`${fg} sobre ${bg}`, () => {
          expect(tokens[fg], `${fg} ausente no tema ${tema}`).toMatch(/^#[0-9A-Fa-f]{6}$/)
          expect(tokens[bg], `${bg} ausente no tema ${tema}`).toMatch(/^#[0-9A-Fa-f]{6}$/)
          const razao = contrastRatio(tokens[fg], tokens[bg])
          expect(razao, `${fg} (${tokens[fg]}) × ${bg} (${tokens[bg]}) = ${razao.toFixed(2)}`).toBeGreaterThanOrEqual(4.5)
        })
      }
    })
  }

  it('azul da marca sobre branco fica entre 9 e 9,2', () => {
    const razao = contrastRatio('#0131CD', '#FFFFFF')
    expect(razao).toBeGreaterThanOrEqual(9)
    expect(razao).toBeLessThanOrEqual(9.2)
  })
})
