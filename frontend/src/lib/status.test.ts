import { describe, it, expect } from 'vitest'
import { statusTone } from './status'

describe('statusTone', () => {
  it.each([
    ['Disponível', 'available'], ['Ativo', 'available'], ['Ativa', 'available'], ['Assinado', 'available'], ['OK', 'available'],
    ['Indisponível', 'loaned'], ['Em uso', 'loaned'], ['Em Uso', 'loaned'], ['Enviado', 'loaned'],
    ['Pendente', 'pending'], ['Agendado', 'pending'], ['Aberta', 'pending'], ['Aberto', 'pending'],
    ['Pendente Devolução', 'return'],
    ['Em Manutenção', 'maintenance'], ['Enviada', 'maintenance'], ['Em reparo', 'maintenance'],
    ['Baixado', 'retired'], ['Descartado', 'retired'], ['Substituido', 'retired'], ['Desativado', 'retired'],
    ['Cancelada', 'retired'], ['Cancelado', 'retired'], ['Desligado', 'retired'], ['Encerrado', 'retired'],
    ['Erro', 'danger'], ['Recusado', 'danger'], ['Expirado', 'danger'], ['Sem conserto', 'danger'],
    ['Qualquer Coisa', 'neutral'], ['', 'neutral'],
  ])('%s -> %s', (status, tone) => {
    expect(statusTone(status)).toBe(tone)
  })

  it('ignora acento e caixa', () => {
    expect(statusTone('DISPONIVEL')).toBe('available')
    expect(statusTone('em manutencao')).toBe('maintenance')
    expect(statusTone('pendente devolucao')).toBe('return')
    expect(statusTone('  Ativo ')).toBe('available')
  })

  it('null e undefined são neutros', () => {
    expect(statusTone(null)).toBe('neutral')
    expect(statusTone(undefined)).toBe('neutral')
  })
})
