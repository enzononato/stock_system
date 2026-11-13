import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import {
  createAcesso,
  desativarAcesso,
  listEventos,
  resolverTarefa,
  savePacote,
} from '@/api/acessos'

describe('api/acessos', () => {
  it('createAcesso envia exatamente o corpo informado em POST /acessos/ficha', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/ficha', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ id: 1 }, { status: 201 })
      })
    )
    const dados = { colaborador_id: 7, sistema_id: 1, login: 'maria.silva', item_id: null }
    await createAcesso(dados)
    expect(corpo).toEqual(dados)
  })

  it('listEventos manda os filtros na query', async () => {
    let url = ''
    server.use(
      http.get('/api/acessos/eventos', ({ request }) => {
        url = request.url
        return HttpResponse.json({ items: [], total: 0 })
      })
    )
    await listEventos({ status: 'aberto', limit: 25, offset: 0 })
    expect(new URL(url).pathname + new URL(url).search).toBe(
      '/api/acessos/eventos?status=aberto&limit=25&offset=0'
    )
  })

  it('resolverTarefa faz POST /acessos/tarefas/:id/resolver', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/tarefas/5/resolver', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ id: 1 })
      })
    )
    await resolverTarefa(5, { status: 'feito', login: 'a', item_id: 11 })
    expect(corpo).toEqual({ status: 'feito', login: 'a', item_id: 11 })
  })

  it('savePacote usa PUT /acessos/pacotes', async () => {
    let metodo = ''
    server.use(
      http.put('/api/acessos/pacotes', ({ request }) => {
        metodo = request.method
        return HttpResponse.json({ cargo: 'Vendedor', sistemas: [] })
      })
    )
    await savePacote({ cargo: 'Vendedor', sistema_ids: [] })
    expect(metodo).toBe('PUT')
  })

  it('desativarAcesso faz POST /acessos/ficha/:id/desativar', async () => {
    let chamado = false
    server.use(
      http.post('/api/acessos/ficha/3/desativar', () => {
        chamado = true
        return HttpResponse.json({ id: 3, status: 'desativado' })
      })
    )
    const r = await desativarAcesso(3)
    expect(chamado).toBe(true)
    expect(r.status).toBe('desativado')
  })
})
