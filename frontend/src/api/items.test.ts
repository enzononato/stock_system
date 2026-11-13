import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import {
  getLinhaDoTempo,
  listItemsPaginated,
  removeItem,
  transferirItem,
  uploadNotaFiscal,
} from '@/api/items'

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

// O File do jsdom não passa pelo parser de FormData do MSW (`request.formData()` estoura com
// arquivo): para multipart com arquivo, os nomes dos campos saem do corpo em texto.
async function nomesDeCampos(request: Request): Promise<string[]> {
  const corpo = await request.text()
  return [...corpo.matchAll(/; name="([^"]+)"/g)].map((m) => m[1])
}

describe('ciclo de vida do equipamento', () => {
  it('removeItem envia o apagamento de dados, sem comprovante quando não há', async () => {
    let campos: Record<string, string> = {}
    server.use(
      http.delete('/api/items/77', async ({ request }) => {
        const form = await request.formData()
        campos = Object.fromEntries([...form.entries()].map(([k, v]) => [k, String(v)]))
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    await removeItem(77, 'Doação', undefined, {
      dados_apagados: 'sim',
      metodo_apagamento: 'formatacao_segura',
      apagado_por: 'Ana',
      apagado_em: '10/11/2026',
    })
    expect(campos).toEqual({
      reason: 'Doação',
      dados_apagados: 'sim',
      metodo_apagamento: 'formatacao_segura',
      apagado_por: 'Ana',
      apagado_em: '10/11/2026',
    })
    expect(campos.comprovante_apagamento).toBeUndefined()
  })

  it('removeItem manda o comprovante como comprovante_apagamento', async () => {
    let nomes: string[] = []
    server.use(
      http.delete('/api/items/78', async ({ request }) => {
        nomes = await nomesDeCampos(request)
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    await removeItem(78, 'Doação', undefined, {
      dados_apagados: 'sim',
      metodo_apagamento: 'reset_fabrica',
      apagado_por: 'Ana',
      apagado_em: '10/11/2026',
      comprovante: new File(['%PDF'], 'c.pdf', { type: 'application/pdf' }),
    })
    expect(nomes).toContain('comprovante_apagamento')
  })

  it('removeItem sem apagamento segue como antes', async () => {
    let nomes: string[] = []
    server.use(
      http.delete('/api/items/79', async ({ request }) => {
        nomes = [...(await request.formData()).keys()]
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    await removeItem(79, 'Perda')
    expect(nomes).toEqual(['reason'])
  })

  it('uploadNotaFiscal envia o campo arquivo', async () => {
    let nomes: string[] = []
    server.use(
      http.post('/api/items/5/nota-fiscal', async ({ request }) => {
        nomes = await nomesDeCampos(request)
        return HttpResponse.json({ detail: 'Nota fiscal anexada.', nf_anexo_key: 'notas_fiscais/x.pdf' })
      })
    )
    const r = await uploadNotaFiscal(5, new File(['%PDF'], 'nf.pdf', { type: 'application/pdf' }))
    expect(nomes).toEqual(['arquivo'])
    expect(r.nf_anexo_key).toBe('notas_fiscais/x.pdf')
  })

  it('transferirItem envia o JSON', async () => {
    let corpo: unknown
    server.use(
      http.post('/api/items/5/transferir', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: 'Equipamento transferido.', id: 1 })
      })
    )
    await transferirItem(5, { para_unidade_id: 2, para_local_id: null, motivo: 'Mudança de setor' })
    expect(corpo).toEqual({ para_unidade_id: 2, para_local_id: null, motivo: 'Mudança de setor' })
  })

  it('getLinhaDoTempo lê a linha do tempo do equipamento', async () => {
    server.use(
      http.get('/api/items/5/linha-do-tempo', () =>
        HttpResponse.json({ item_id: 5, eventos: [], fontes_com_erro: [] })
      )
    )
    expect(await getLinhaDoTempo(5)).toEqual({ item_id: 5, eventos: [], fontes_com_erro: [] })
  })
})
