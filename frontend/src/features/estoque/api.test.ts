import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import { listEstoque, proximaOrdenacao } from './api'

describe('listEstoque', () => {
  it('repassa sort, dir, colaborador_id, limit e offset como query params', async () => {
    let params!: URLSearchParams
    server.use(
      http.get('/api/items', ({ request }) => {
        params = new URL(request.url).searchParams
        return HttpResponse.json({ items: [], total: 0 })
      })
    )
    await listEstoque({ colaborador_id: 7, sort: 'setor', dir: 'asc', limit: 25, offset: 50 })
    expect(params.get('sort')).toBe('setor')
    expect(params.get('dir')).toBe('asc')
    expect(params.get('colaborador_id')).toBe('7')
    expect(params.get('limit')).toBe('25')
    expect(params.get('offset')).toBe('50')
  })
})

describe('proximaOrdenacao', () => {
  it('coluna nova começa em asc', () => {
    expect(proximaOrdenacao({ sort: 'id', dir: 'desc' }, 'tipo')).toEqual({ sort: 'tipo', dir: 'asc' })
  })
  it('mesma coluna inverte', () => {
    expect(proximaOrdenacao({ sort: 'tipo', dir: 'asc' }, 'tipo')).toEqual({ sort: 'tipo', dir: 'desc' })
    expect(proximaOrdenacao({ sort: 'tipo', dir: 'desc' }, 'tipo')).toEqual({ sort: 'tipo', dir: 'asc' })
  })
})
