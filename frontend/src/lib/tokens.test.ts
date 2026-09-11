import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const css = readFileSync(resolve(__dirname, '../index.css'), 'utf8')
const config = readFileSync(resolve(__dirname, '../../tailwind.config.js'), 'utf8')

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
