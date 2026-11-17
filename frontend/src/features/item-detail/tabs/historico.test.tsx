import { describe, expect, it } from 'vitest'
import tab from './historico'

describe('aba Histórico', () => {
  it('está registrada como aba 90 só para Gestor e Técnico (a API é restrita a eles)', () => {
    expect(tab.id).toBe('historico')
    expect(tab.label).toBe('Histórico')
    expect(tab.order).toBe(90)
    expect(tab.roles).toEqual(['Gestor', 'Técnico'])
  })
})
