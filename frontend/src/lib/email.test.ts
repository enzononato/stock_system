import { describe, expect, it } from 'vitest'
import { isValidEmail } from './email'

describe('isValidEmail', () => {
  it('aceita endereços comuns, inclusive pessoais', () => {
    expect(isValidEmail('maria.souza@gmail.com')).toBe(true)
    expect(isValidEmail(' joao@empresa.com.br ')).toBe(true)
  })

  it('recusa endereços incompletos ou com espaço', () => {
    expect(isValidEmail('')).toBe(false)
    expect(isValidEmail('joao@')).toBe(false)
    expect(isValidEmail('joao@empresa')).toBe(false)
    expect(isValidEmail('joao silva@empresa.com')).toBe(false)
  })
})
