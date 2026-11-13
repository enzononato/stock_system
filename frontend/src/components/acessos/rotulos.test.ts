import { describe, expect, it } from 'vitest'
import {
  descreverEquipamento,
  ROTULO_TIPO_EVENTO,
  STATUS_EVENTO,
  STATUS_TAREFA,
} from './rotulos'

describe('rótulos do módulo de acessos', () => {
  it('descreve o equipamento como o backend', () => {
    expect(
      descreverEquipamento({ id: 4, tipo: 'Notebook', brand: 'Dell', model: 'X', identificador: 'NB-1' })
    ).toBe('Notebook Dell X · NB-1')
    expect(descreverEquipamento({ id: 4 })).toBe('Equipamento #4')
  })

  it('mapeia tons e rótulos', () => {
    expect(STATUS_EVENTO.concluido.tone).toBe('available')
    expect(STATUS_EVENTO.aberto.label).toBe('Aberto')
    expect(STATUS_TAREFA.nao_se_aplica.label).toBe('Não se aplica')
    expect(ROTULO_TIPO_EVENTO.mudanca_funcao).toBe('Mudança de função')
  })
})
