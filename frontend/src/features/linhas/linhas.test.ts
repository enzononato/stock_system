import { describe, expect, it } from 'vitest'
import type { LinhaEvento, TipoEventoLinha } from '@/api/linhas'
import { formatMoeda } from '@/lib/format'
import { custoParaInput, descreverAparelho, parseCusto } from './format'
import { descreverEvento, ROTULO_CAMPO, ROTULO_EVENTO } from './eventos'
import { acoesDisponiveis, STATUS_DA_ACAO, STATUS_LINHA_OPCOES, statusLinhaInfo } from './status'

function evento(parcial: Partial<LinhaEvento> & { tipo: TipoEventoLinha }): LinhaEvento {
  return {
    id: 1,
    linha_id: 5,
    linha_numero: '74999990001',
    campo: null,
    de_valor: null,
    para_valor: null,
    item_id: null,
    item_id_anterior: null,
    ocorrido_em: '2026-11-12T10:00:00',
    por_usuario: 'gestor',
    observacao: null,
    ...parcial,
  }
}

describe('parseCusto', () => {
  it.each([
    ['49,90', '49.90'],
    ['1.234,56', '1234.56'],
    ['R$ 49,9', '49.90'],
    ['49.9', '49.90'],
    ['49', '49.00'],
    ['0', '0.00'],
    ['  R$ 0,05 ', '0.05'],
  ])('converte %s em %s', (entrada, esperado) => {
    expect(parseCusto(entrada)).toBe(esperado)
  })

  it.each(['49,999', '-1', 'abc', '', '1.234', '12,345', '1,2,3'])('recusa %s', (entrada) => {
    expect(parseCusto(entrada)).toBeNull()
  })
})

describe('custoParaInput', () => {
  it('troca o ponto pela vírgula e aceita vazio', () => {
    expect(custoParaInput('49.90')).toBe('49,90')
    expect(custoParaInput(59.9)).toBe('59,9')
    expect(custoParaInput(null)).toBe('')
    expect(custoParaInput(undefined)).toBe('')
    expect(custoParaInput('')).toBe('')
  })
})

describe('descreverAparelho', () => {
  it('monta o mesmo texto do backend', () => {
    expect(
      descreverAparelho({ tipo: 'Celular', brand: 'Samsung', model: 'A15', identificador: 'CEL-A' })
    ).toBe('Celular Samsung A15 · CEL-A')
    expect(descreverAparelho({ tipo: 'Tablet', brand: 'Apple', model: 'iPad' })).toBe('Tablet Apple iPad')
    expect(descreverAparelho(null)).toBe('')
    expect(descreverAparelho(undefined)).toBe('')
    expect(descreverAparelho({})).toBe('')
  })
})

describe('status da linha', () => {
  it('lista os três status com tom', () => {
    expect(STATUS_LINHA_OPCOES).toEqual([
      { value: 'ativa', label: 'Ativa', tone: 'available' },
      { value: 'suspensa', label: 'Suspensa', tone: 'pending' },
      { value: 'cancelada', label: 'Cancelada', tone: 'retired' },
    ])
    expect(statusLinhaInfo('suspensa')).toEqual({ label: 'Suspensa', tone: 'pending' })
  })

  it('define as ações por status', () => {
    expect(acoesDisponiveis('ativa')).toEqual(['suspender', 'cancelar'])
    expect(acoesDisponiveis('suspensa')).toEqual(['reativar', 'cancelar'])
    expect(acoesDisponiveis('cancelada')).toEqual(['reativar'])
    expect(STATUS_DA_ACAO).toEqual({ suspender: 'suspensa', reativar: 'ativa', cancelar: 'cancelada' })
  })
})

describe('descreverEvento', () => {
  it('rotula tipos e campos', () => {
    expect(ROTULO_EVENTO.troca_aparelho).toBe('Troca de aparelho')
    expect(ROTULO_EVENTO.ativacao).toBe('Ativação')
    expect(ROTULO_CAMPO.custo_mensal).toBe('Custo mensal')
  })

  it('troca de aparelho sem aparelho anterior', () => {
    const e = evento({ tipo: 'troca_aparelho', de_valor: null, para_valor: 'Celular Motorola G54 · CEL-B' })
    expect(descreverEvento(e)).toBe('sem aparelho → Celular Motorola G54 · CEL-B')
  })

  it('troca de usuário e portabilidade', () => {
    expect(descreverEvento(evento({ tipo: 'troca_usuario', de_valor: 'Maria', para_valor: null }))).toBe(
      'Maria → sem usuário'
    )
    expect(descreverEvento(evento({ tipo: 'portabilidade', de_valor: 'Vivo', para_valor: 'Claro' }))).toBe(
      'Vivo → Claro'
    )
  })

  it('eventos de status', () => {
    expect(descreverEvento(evento({ tipo: 'ativacao' }))).toBe('Linha ativada')
    expect(descreverEvento(evento({ tipo: 'suspensao' }))).toBe('Linha suspensa')
    expect(descreverEvento(evento({ tipo: 'reativacao' }))).toBe('Linha reativada')
    expect(descreverEvento(evento({ tipo: 'cancelamento' }))).toBe('Linha cancelada')
  })

  it('edição formata custo, número e usa traço para vazio', () => {
    expect(
      descreverEvento(evento({ tipo: 'edicao', campo: 'custo_mensal', de_valor: '49.90', para_valor: '59.90' }))
    ).toBe(`Custo mensal: ${formatMoeda('49.90')} → ${formatMoeda('59.90')}`)
    expect(
      descreverEvento(
        evento({ tipo: 'edicao', campo: 'numero', de_valor: '74999990001', para_valor: '74999990002' })
      )
    ).toBe('Número: (74) 99999-0001 → (74) 99999-0002')
    expect(
      descreverEvento(evento({ tipo: 'edicao', campo: 'plano', de_valor: null, para_valor: 'Controle 30GB' }))
    ).toBe('Plano: — → Controle 30GB')
    expect(
      descreverEvento(evento({ tipo: 'edicao', campo: 'unidade', de_valor: 'Juazeiro', para_valor: 'Petrolina' }))
    ).toBe('Unidade: JUAZEIRO → PETROLINA')
  })
})
