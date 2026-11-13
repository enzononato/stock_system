import { describe, expect, it } from 'vitest'
import {
  ORDEM_CATEGORIAS,
  ROTULO_CATEGORIA,
  ROTULO_RESULTADO,
  ROTULO_STATUS_INVENTARIO,
  TOM_RESULTADO,
  rotuloCategoria,
} from '@/features/inventario/resultados'
import { statusTone } from '@/lib/status'

describe('resultados do inventário', () => {
  it('"não encontrado" vira "Ainda não lidos" só com o inventário aberto', () => {
    expect(rotuloCategoria('nao_encontrado', 'aberto')).toBe('Ainda não lidos')
    expect(rotuloCategoria('nao_encontrado', 'encerrado')).toBe('Não encontrados')
    expect(rotuloCategoria('ok', 'aberto')).toBe('Conferidos')
  })

  it('tons e ordem seguem o backend', () => {
    expect(TOM_RESULTADO.ok).toBe('available')
    expect(TOM_RESULTADO.nao_cadastrado).toBe('danger')
    expect(ORDEM_CATEGORIAS[0]).toBe('nao_encontrado')
    expect(ORDEM_CATEGORIAS[ORDEM_CATEGORIAS.length - 1]).toBe('ok')
    expect(ORDEM_CATEGORIAS).toHaveLength(6)
  })

  it('rótulos no singular e no plural', () => {
    expect(ROTULO_RESULTADO.ok).toBe('Conferido')
    expect(ROTULO_RESULTADO.nao_encontrado).toBe('Não encontrado')
    expect(ROTULO_CATEGORIA.outra_unidade).toBe('De outra unidade')
  })

  it('os rótulos de status do inventário já têm tom no StatusPill', () => {
    expect(statusTone(ROTULO_STATUS_INVENTARIO.aberto)).toBe('pending')
    expect(statusTone(ROTULO_STATUS_INVENTARIO.encerrado)).toBe('retired')
    expect(statusTone(ROTULO_STATUS_INVENTARIO.cancelado)).toBe('retired')
  })
})
