import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import { listItemsPaginated } from '@/api/items'

describe('listItemsPaginated', () => {
  it('repassa limit/offset/search como query params e devolve { items, total }', async () => {
    let queryRecebida: Record<string, string> = {}
    server.use(
      http.get('/api/items', ({ request }) => {
        const url = new URL(request.url)
        queryRecebida = Object.fromEntries(url.searchParams.entries())
        return HttpResponse.json({
          items: [{ id: 1, tipo: 'Notebook', identificador: 'SN-1' }],
          total: 42,
        })
      })
    )

    const resultado = await listItemsPaginated({ limit: 10, offset: 20, search: 'dell' })

    expect(queryRecebida).toEqual({ limit: '10', offset: '20', search: 'dell' })
    expect(resultado).toEqual({
      items: [{ id: 1, tipo: 'Notebook', identificador: 'SN-1' }],
      total: 42,
    })
  })

  it('repassa também tipo/status/revenda quando fornecidos', async () => {
    let queryRecebida: Record<string, string> = {}
    server.use(
      http.get('/api/items', ({ request }) => {
        const url = new URL(request.url)
        queryRecebida = Object.fromEntries(url.searchParams.entries())
        return HttpResponse.json({ items: [], total: 0 })
      })
    )

    await listItemsPaginated({ tipo: 'Notebook', status: 'Disponível', revenda: 'Revalle Juazeiro' })

    expect(queryRecebida).toEqual({
      tipo: 'Notebook',
      status: 'Disponível',
      revenda: 'Revalle Juazeiro',
    })
  })

  it('sem parâmetros, não envia nenhum query param', async () => {
    let url: URL | null = null
    server.use(
      http.get('/api/items', ({ request }) => {
        url = new URL(request.url)
        return HttpResponse.json({ items: [], total: 0 })
      })
    )

    const resultado = await listItemsPaginated()

    expect(url!.search).toBe('')
    expect(resultado).toEqual({ items: [], total: 0 })
  })
})
