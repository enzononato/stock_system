import { describe, it, expect } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { resolve, join, relative } from 'node:path'

const css = readFileSync(resolve(__dirname, '../index.css'), 'utf8')
const config = readFileSync(resolve(__dirname, '../../tailwind.config.js'), 'utf8')

// --- Guarda contra opacidade em token customizado (ver descrição abaixo) ---

/**
 * Recorta um bloco `{ ... }` balanceado do texto, a partir do índice do `{`
 * de abertura. Usado para isolar `colors: { ... }` (e cada objeto aninhado
 * dentro dele) sem depender de um parser de JS de verdade — o arquivo não é
 * JSON válido (chaves sem aspas, vírgulas finais), então um parser de
 * chaves/vírgulas simples é o suficiente e não exige nenhuma dependência nova.
 */
function recortarBlocoBalanceado(texto: string, indiceAbre: number): string {
  let profundidade = 0
  for (let i = indiceAbre; i < texto.length; i++) {
    if (texto[i] === '{') profundidade++
    else if (texto[i] === '}') {
      profundidade--
      if (profundidade === 0) return texto.slice(indiceAbre, i + 1)
    }
  }
  throw new Error('Bloco de chaves não balanceado em tailwind.config.js')
}

/** Divide o conteúdo de um objeto `{ chave: valor, ... }` em entradas de topo (sem entrar em objetos aninhados). */
function dividirEntradasDeTopo(bloco: string): { chave: string; valorBruto: string }[] {
  const interior = bloco.slice(1, -1)
  const entradas: string[] = []
  let profundidade = 0
  let atual = ''
  for (const ch of interior) {
    if (ch === '{') profundidade++
    if (ch === '}') profundidade--
    if (ch === ',' && profundidade === 0) {
      entradas.push(atual)
      atual = ''
    } else {
      atual += ch
    }
  }
  if (atual.trim()) entradas.push(atual)

  return entradas
    .map((e) => e.trim())
    .filter(Boolean)
    .map((entrada) => {
      const idxDoisPontos = entrada.indexOf(':')
      const chave = entrada.slice(0, idxDoisPontos).trim().replace(/^['"]|['"]$/g, '')
      const valorBruto = entrada.slice(idxDoisPontos + 1).trim()
      return { chave, valorBruto }
    })
}

/**
 * Extrai todo token de cor definido em `theme.extend.colors` do
 * tailwind.config.js, incluindo os aninhados (ex.: `surface.alt` vira o
 * token `surface-alt`; `DEFAULT` aninhado vira o próprio nome do pai, ex.:
 * `primary.DEFAULT` vira `primary`). A lista é derivada do arquivo, não
 * copiada à mão, então não fica desatualizada quando alguém adiciona um token.
 */
function extrairTokensDeCor(configTexto: string): string[] {
  const idx = configTexto.indexOf('colors:')
  if (idx === -1) throw new Error('"colors:" não encontrado em tailwind.config.js')
  const indiceAbre = configTexto.indexOf('{', idx)
  const bloco = recortarBlocoBalanceado(configTexto, indiceAbre)

  const tokens = new Set<string>()
  for (const { chave, valorBruto } of dividirEntradasDeTopo(bloco)) {
    if (valorBruto.startsWith('{')) {
      for (const sub of dividirEntradasDeTopo(valorBruto)) {
        tokens.add(sub.chave === 'DEFAULT' ? chave : `${chave}-${sub.chave}`)
      }
    } else {
      tokens.add(chave)
    }
  }
  return Array.from(tokens)
}

/** Caminha src/**\/*.{ts,tsx} recursivamente com node:fs — sem dependência de glob. */
function listarArquivosFonte(dir: string): string[] {
  const resultado: string[] = []
  for (const entrada of readdirSync(dir, { withFileTypes: true })) {
    const caminho = join(dir, entrada.name)
    if (entrada.isDirectory()) {
      resultado.push(...listarArquivosFonte(caminho))
    } else if (/\.(ts|tsx)$/.test(entrada.name)) {
      resultado.push(caminho)
    }
  }
  return resultado
}

describe('tokens de design', () => {
  it('define a paleta monocromática no modo claro', () => {
    expect(css).toContain('--canvas: #F7F7F5')
    expect(css).toContain('--surface: #FFFFFF')
    expect(css).toContain('--surface-alt: #F0F0ED')
    expect(css).toContain('--border: #D9D9D4')
    expect(css).toContain('--border-strong: #BDBDB7')
    expect(css).toContain('--text-primary: #111111')
  })

  it('define a paleta do modo escuro', () => {
    expect(css).toContain('--canvas: #050505')
    expect(css).toContain('--surface: #0C0C0C')
    expect(css).toContain('--text-primary: #F2F2F0')
  })

  it('define os raios editoriais (2/4/6/8px)', () => {
    expect(css).toContain('--radius-sm: 2px')
    expect(css).toContain('--radius: 4px')
    expect(css).toContain('--radius-md: 6px')
    expect(css).toContain('--radius-lg: 8px')
  })

  it('mantém cor semântica apenas para status', () => {
    expect(css).toContain('--status-available')
    expect(css).toContain('--status-loaned')
    expect(css).toContain('--status-pending')
    expect(css).toContain('--status-return')
    expect(css).toContain('--status-danger')
  })

  it('expõe os utilitários de tipografia editorial', () => {
    for (const u of ['.text-caption', '.text-body-sm', '.text-body', '.text-heading', '.num']) {
      expect(css).toContain(u)
    }
  })

  it('expõe os utilitários de superfície', () => {
    for (const u of ['.surface-panel', '.surface-interactive', '.figure-ground-panel']) {
      expect(css).toContain(u)
    }
  })

  it('não usa sintaxe do Tailwind 4', () => {
    expect(css).not.toContain('@theme')
    expect(css).not.toContain('@utility')
    expect(css).not.toContain('@custom-variant')
    expect(css).not.toContain('source(none)')
  })

  it('registra os tokens novos no tailwind.config.js', () => {
    for (const t of ['canvas', 'surface', 'border-strong', 'sidebar', 'micro']) {
      expect(config).toContain(t)
    }
  })

  it('não deixa a paleta antiga no config', () => {
    expect(config).not.toContain('glow-indigo')
    expect(config).not.toContain('#6366f1')
  })
})

describe('guarda contra opacidade em token customizado', () => {
  it('nenhuma classe usa modificador de opacidade (ex.: /50) sobre um token customizado', () => {
    // Por quê isso quebra silenciosamente: os tokens deste projeto são
    // registrados no tailwind.config.js como string literal "var(--x)", não
    // no formato "rgb(var(--x) / <alpha-value>)" que o Tailwind exige para
    // suportar o modificador de opacidade (a barra seguida de um número
    // depois do nome do token, tipo "primary" + "/" + "50"). Uma classe
    // assim aplicada a um token deste projeto compila para NENHUM CSS — a
    // classe é silenciosamente descartada, sem erro de build, sem warning
    // do Tailwind e sem nada que o jsdom (que não calcula CSS) consiga
    // detectar. O elemento renderiza sem a cor de fundo/texto esperada.
    // Cores nativas do Tailwind (ex.: preto/branco, usadas em overlay de
    // modal) NÃO são afetadas — usam o formato com <alpha-value> e
    // continuam de fora da lista de tokens abaixo.
    const tokens = extrairTokensDeCor(config)
    const prefixos = [
      'bg', 'text', 'border', 'ring', 'divide', 'from', 'via', 'to', 'fill',
      'stroke', 'placeholder', 'outline', 'decoration', 'accent', 'caret', 'shadow',
    ]
    const padrao = new RegExp(`\\b(?:${prefixos.join('|')})-(?:${tokens.join('|')})/[0-9]+`, 'g')

    const srcDir = resolve(__dirname, '..')
    const ofensores: string[] = []
    for (const arquivo of listarArquivosFonte(srcDir)) {
      const conteudo = readFileSync(arquivo, 'utf8')
      conteudo.split('\n').forEach((linha, i) => {
        const achados = linha.match(padrao)
        if (achados) {
          for (const achado of achados) {
            ofensores.push(`${relative(srcDir, arquivo)}:${i + 1}:${achado}`)
          }
        }
      })
    }

    expect(ofensores, `Modificador de opacidade sobre token customizado (emite CSS vazio):\n${ofensores.join('\n')}`).toEqual([])
  })
})
