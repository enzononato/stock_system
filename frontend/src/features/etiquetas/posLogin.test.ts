import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  CHAVE_DESTINO,
  consumirDestinoPosLogin,
  guardarDestinoPosLogin,
} from '@/features/etiquetas/posLogin'

beforeEach(() => sessionStorage.clear())
afterEach(() => vi.restoreAllMocks())

describe('posLogin', () => {
  it('guarda o destino e o consome uma única vez', () => {
    guardarDestinoPosLogin('/items/42')

    expect(consumirDestinoPosLogin()).toBe('/items/42')
    expect(consumirDestinoPosLogin()).toBeNull()
    expect(sessionStorage.getItem(CHAVE_DESTINO)).toBeNull()
  })

  it('passados 10 minutos o destino não vale mais (e é descartado)', () => {
    guardarDestinoPosLogin('/items/42')

    expect(consumirDestinoPosLogin(Date.now() + 10 * 60 * 1000 + 1)).toBeNull()
    expect(consumirDestinoPosLogin()).toBeNull()
  })

  it('dentro dos 10 minutos ainda vale', () => {
    guardarDestinoPosLogin('/items/42')
    expect(consumirDestinoPosLogin(Date.now() + 9 * 60 * 1000)).toBe('/items/42')
  })

  it('só aceita /items/<número>: nada de redirecionamento aberto', () => {
    guardarDestinoPosLogin('//evil.com')
    guardarDestinoPosLogin('/users')
    guardarDestinoPosLogin('https://evil.com/items/1')
    guardarDestinoPosLogin('/items/1/../../users')

    expect(sessionStorage.getItem(CHAVE_DESTINO)).toBeNull()
    expect(consumirDestinoPosLogin()).toBeNull()
  })

  it('valor adulterado no storage é ignorado', () => {
    sessionStorage.setItem(CHAVE_DESTINO, JSON.stringify({ destino: '//evil.com', em: Date.now() }))
    expect(consumirDestinoPosLogin()).toBeNull()
    sessionStorage.setItem(CHAVE_DESTINO, '{x')
    expect(consumirDestinoPosLogin()).toBeNull()
  })

  it('sessionStorage lançando não quebra', () => {
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new Error('bloqueado')
    })
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new Error('bloqueado')
    })
    expect(() => guardarDestinoPosLogin('/items/1')).not.toThrow()
    expect(consumirDestinoPosLogin()).toBeNull()
  })
})
