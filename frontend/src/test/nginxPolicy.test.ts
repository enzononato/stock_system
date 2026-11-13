import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const conf = readFileSync(resolve(dirname(fileURLToPath(import.meta.url)), '../../nginx.conf.template'), 'utf-8')

describe('nginx.conf.template', () => {
  it('libera a câmera só para o próprio site (leitura de QR do inventário)', () => {
    expect(conf).toContain('Permissions-Policy')
    expect(conf).toContain('camera=(self)')
    expect(conf).not.toMatch(/camera=\(\)/)
  })
})
