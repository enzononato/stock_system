import { afterEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import {
  cancelarInventario,
  createInventario,
  encerrarInventario,
  excluirLeitura,
  exportarRelatorio,
  getRelatorio,
  listInventarios,
  listLeituras,
  registrarLeitura,
  atualizarLeitura,
} from '@/api/inventarios'
import { gerarEtiquetas, listIdsParaEtiquetas } from '@/api/etiquetas'

afterEach(() => vi.restoreAllMocks())

describe('api/inventarios', () => {
  it('registrarLeitura faz POST com o corpo informado', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/inventarios/12/leituras', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ id: 1, resultado: 'ok', duplicada: false }, { status: 201 })
      })
    )

    const leitura = await registrarLeitura(12, { codigo: '#5', local_id_encontrado: 3 })

    expect(corpo).toEqual({ codigo: '#5', local_id_encontrado: 3 })
    expect(leitura.resultado).toBe('ok')
  })

  it('exportarRelatorio baixa o arquivo autenticado', async () => {
    let url = ''
    server.use(
      http.get('/api/inventarios/12/relatorio/exportar', ({ request }) => {
        url = new URL(request.url).search
        return new HttpResponse(new Blob(['%PDF']), { headers: { 'Content-Type': 'application/pdf' } })
      })
    )
    const criar = vi.fn(() => 'blob:x')
    const revogar = vi.fn()
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
    Object.defineProperty(URL, 'createObjectURL', { value: criar, configurable: true })
    Object.defineProperty(URL, 'revokeObjectURL', { value: revogar, configurable: true })

    await exportarRelatorio(12, 'pdf')

    expect(url).toBe('?formato=pdf')
    expect(criar).toHaveBeenCalledTimes(1)
  })

  it('listInventarios, createInventario, getRelatorio, cancelar e encerrar usam as rotas certas', async () => {
    const chamadas: string[] = []
    server.use(
      http.get('/api/inventarios', ({ request }) => {
        chamadas.push(`GET /inventarios${new URL(request.url).search}`)
        return HttpResponse.json({ items: [], total: 0 })
      }),
      http.post('/api/inventarios', async ({ request }) => {
        chamadas.push(`POST /inventarios ${JSON.stringify(await request.json())}`)
        return HttpResponse.json({ id: 1 }, { status: 201 })
      }),
      http.get('/api/inventarios/1/relatorio', () => {
        chamadas.push('GET relatorio')
        return HttpResponse.json({ inventario: { id: 1 }, linhas: [] })
      }),
      http.post('/api/inventarios/1/encerrar', () => {
        chamadas.push('POST encerrar')
        return HttpResponse.json({ id: 1 })
      }),
      http.post('/api/inventarios/1/cancelar', async ({ request }) => {
        chamadas.push(`POST cancelar ${JSON.stringify(await request.json())}`)
        return HttpResponse.json({ id: 1 })
      }),
      http.get('/api/inventarios/1/leituras', ({ request }) => {
        chamadas.push(`GET leituras${new URL(request.url).search}`)
        return HttpResponse.json({ items: [], total: 0 })
      }),
      http.patch('/api/inventarios/1/leituras/9', async ({ request }) => {
        chamadas.push(`PATCH ${JSON.stringify(await request.json())}`)
        return HttpResponse.json({ id: 9 })
      }),
      http.delete('/api/inventarios/1/leituras/9', () => {
        chamadas.push('DELETE leitura')
        return HttpResponse.json({ detail: 'Leitura excluída.' })
      })
    )

    await listInventarios({ unidade_id: 2, status: 'aberto', limit: 25, offset: 0 })
    await createInventario({ unidade_id: 2, observacao: 'x' })
    await getRelatorio(1)
    await encerrarInventario(1)
    await cancelarInventario(1, 'Contagem refeita')
    await listLeituras(1, { limit: 10 })
    await atualizarLeitura(1, 9, { usuario_conferido: true, colaborador_id_encontrado: null })
    await excluirLeitura(1, 9)

    expect(chamadas).toEqual([
      'GET /inventarios?unidade_id=2&status=aberto&limit=25&offset=0',
      'POST /inventarios {"unidade_id":2,"observacao":"x"}',
      'GET relatorio',
      'POST encerrar',
      'POST cancelar {"motivo":"Contagem refeita"}',
      'GET leituras?limit=10',
      'PATCH {"usuario_conferido":true,"colaborador_id_encontrado":null}',
      'DELETE leitura',
    ])
  })
})

describe('api/etiquetas', () => {
  it('listIdsParaEtiquetas e gerarEtiquetas', async () => {
    let query = ''
    let corpo: unknown = null
    server.use(
      http.get('/api/etiquetas/ids', ({ request }) => {
        query = new URL(request.url).search
        return HttpResponse.json({ ids: [1, 2], total: 2 })
      }),
      http.post('/api/etiquetas', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ base_url: 'https://x', etiquetas: [], ignorados: [] })
      })
    )

    expect(await listIdsParaEtiquetas({ revenda: 'Juazeiro', tipo: 'Notebook' })).toEqual({ ids: [1, 2], total: 2 })
    expect(query).toBe('?revenda=Juazeiro&tipo=Notebook')
    await gerarEtiquetas([3, 5])
    expect(corpo).toEqual({ item_ids: [3, 5] })
  })
})
