import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import { listHistoryPaginated, reverseEntryWithPassword } from '@/api/history'

describe('listHistoryPaginated', () => {
  it('repassa limit/offset/search como query params e devolve { items, total }', async () => {
    let queryRecebida: Record<string, string> = {}
    server.use(
      http.get('/api/history', ({ request }) => {
        const url = new URL(request.url)
        queryRecebida = Object.fromEntries(url.searchParams.entries())
        return HttpResponse.json({
          items: [{ id: 5, operation: 'Cadastro', operador: 'fulano' }],
          total: 7,
        })
      })
    )

    const resultado = await listHistoryPaginated({ limit: 20, offset: 40, search: 'fulano' })

    expect(queryRecebida).toEqual({ limit: '20', offset: '40', search: 'fulano' })
    expect(resultado).toEqual({
      items: [{ id: 5, operation: 'Cadastro', operador: 'fulano' }],
      total: 7,
    })
  })

  it('sem search, não envia o parâmetro (undefined é omitido pelo axios)', async () => {
    let url: URL | null = null
    server.use(
      http.get('/api/history', ({ request }) => {
        url = new URL(request.url)
        return HttpResponse.json({ items: [], total: 0 })
      })
    )

    await listHistoryPaginated({ limit: 20, offset: 0 })

    expect(url!.searchParams.has('search')).toBe(false)
    expect(url!.searchParams.get('limit')).toBe('20')
    expect(url!.searchParams.get('offset')).toBe('0')
  })
})

describe('reverseEntryWithPassword', () => {
  it('envia a senha no corpo e devolve os dados de sucesso', async () => {
    let corpoRecebido: unknown = null
    server.use(
      http.post('/api/history/123/reverse', async ({ request }) => {
        corpoRecebido = await request.json()
        return HttpResponse.json({ detail: 'Operação estornada com sucesso.' })
      })
    )

    const resultado = await reverseEntryWithPassword(123, 'minha-senha')

    expect(corpoRecebido).toEqual({ password: 'minha-senha' })
    expect(resultado).toEqual({ detail: 'Operação estornada com sucesso.' })
  })

  it('propaga o erro 403 (senha incorreta) para quem chamou', async () => {
    server.use(
      http.post('/api/history/123/reverse', () =>
        HttpResponse.json({ detail: 'Senha incorreta.' }, { status: 403 })
      )
    )

    await expect(reverseEntryWithPassword(123, 'senha-errada')).rejects.toMatchObject({
      response: { status: 403, data: { detail: 'Senha incorreta.' } },
    })
  })
})
