import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import {
  alterarStatusLinha,
  exportarRelatorioCusto,
  getRelatorioCusto,
  listAparelhos,
  listEventosDoAparelho,
  listEventosLinha,
  listLinhas,
  updateLinha,
} from '@/api/linhas'

describe('api/linhas', () => {
  it('listLinhas envia os filtros como parâmetros', async () => {
    let params: URLSearchParams | null = null
    server.use(
      http.get('/api/linhas', ({ request }) => {
        params = new URL(request.url).searchParams
        return HttpResponse.json({ items: [], total: 0 })
      })
    )

    await listLinhas({ status: 'ativa', unidade_id: 2, limit: 25, offset: 0 })

    expect(Object.fromEntries(params!.entries())).toEqual({
      status: 'ativa',
      unidade_id: '2',
      limit: '25',
      offset: '0',
    })
  })

  it('alterarStatusLinha faz POST com status e observação', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/linhas/5/status', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ id: 5, status: 'cancelada' })
      })
    )

    await alterarStatusLinha(5, 'cancelada', 'Chip perdido')

    expect(corpo).toEqual({ status: 'cancelada', observacao: 'Chip perdido' })
  })

  it('updateLinha faz PUT e os eventos usam as rotas certas', async () => {
    let put: unknown = null
    let queryEventos: string | null = null
    let queryAparelhos: URLSearchParams | null = null
    server.use(
      http.put('/api/linhas/5', async ({ request }) => {
        put = await request.json()
        return HttpResponse.json({ id: 5 })
      }),
      http.get('/api/linhas/5/eventos', () => HttpResponse.json([])),
      http.get('/api/linhas/eventos', ({ request }) => {
        queryEventos = new URL(request.url).searchParams.get('item_id')
        return HttpResponse.json([])
      }),
      http.get('/api/linhas/aparelhos', ({ request }) => {
        queryAparelhos = new URL(request.url).searchParams
        return HttpResponse.json([])
      }),
      http.get('/api/linhas/relatorio/custo', () =>
        HttpResponse.json({ unidades: [], total_linhas: 0, custo_total: '0.00', gerado_em: '2026-11-12T10:00:00' })
      )
    )

    await updateLinha(5, { plano: 'X', motivo: 'Upgrade' })
    expect(put).toEqual({ plano: 'X', motivo: 'Upgrade' })
    expect(await listEventosLinha(5)).toEqual([])
    expect(await listEventosDoAparelho(9)).toEqual([])
    expect(queryEventos).toBe('9')
    await listAparelhos({ busca: 'A15', linha_id: 9, limit: 20 })
    expect(Object.fromEntries(queryAparelhos!.entries())).toEqual({ busca: 'A15', linha_id: '9', limit: '20' })
    expect((await getRelatorioCusto()).custo_total).toBe('0.00')
  })

  describe('exportarRelatorioCusto', () => {
    beforeEach(() => {
      vi.stubGlobal(
        'URL',
        Object.assign(URL, {
          createObjectURL: vi.fn(() => 'blob:mock-url'),
          revokeObjectURL: vi.fn(),
        })
      )
    })
    afterEach(() => {
      vi.unstubAllGlobals()
    })

    it('baixa o PDF autenticado com nome custo_linhas_AAAAMMDD.pdf', async () => {
      let url = ''
      server.use(
        http.get('/api/linhas/relatorio/custo/export', ({ request }) => {
          url = new URL(request.url).search
          return HttpResponse.text('%PDF', { headers: { 'Content-Type': 'application/pdf' } })
        })
      )
      let download = ''
      const clickSpy = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (
        this: HTMLAnchorElement
      ) {
        download = this.download
      })

      await exportarRelatorioCusto('pdf')

      expect(url).toBe('?formato=pdf')
      expect(download).toMatch(/^custo_linhas_\d{8}\.pdf$/)
      clickSpy.mockRestore()
    })
  })
})
