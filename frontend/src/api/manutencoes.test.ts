import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import { listManutencoes, registrarRetorno, abrirManutencao, baixarAnexoManutencao } from '@/api/manutencoes'

describe('manutencoes API', () => {
  it('listManutencoes repassa os parâmetros', async () => {
    let query: Record<string, string> = {}
    server.use(
      http.get('/api/manutencoes', ({ request }) => {
        query = Object.fromEntries(new URL(request.url).searchParams.entries())
        return HttpResponse.json({ items: [], total: 0, custo_total: '0.00' })
      })
    )
    const r = await listManutencoes({ status: 'em_andamento', limit: 25, offset: 0 })
    expect(query).toEqual({ status: 'em_andamento', limit: '25', offset: '0' })
    expect(r).toEqual({ items: [], total: 0, custo_total: '0.00' })
  })

  it('registrarRetorno envia multipart com custo de duas casas', async () => {
    let campos: Record<string, string> = {}
    server.use(
      http.post('/api/manutencoes/5/retorno', async ({ request }) => {
        const form = await request.formData()
        campos = Object.fromEntries([...form.entries()].map(([k, v]) => [k, String(v)]))
        return HttpResponse.json({ detail: 'ok', sugerir_baixa: false })
      })
    )
    await registrarRetorno(5, { resultado: 'retornada', custo: 150 })
    expect(campos.resultado).toBe('retornada')
    expect(campos.custo).toBe('150.00')
    expect(campos.anexo).toBeUndefined()
  })

  it('registrarRetorno sem custo não envia o campo', async () => {
    let campos: string[] = []
    server.use(
      http.post('/api/manutencoes/6/retorno', async ({ request }) => {
        campos = [...(await request.formData()).keys()]
        return HttpResponse.json({ detail: 'ok', sugerir_baixa: true })
      })
    )
    const r = await registrarRetorno(6, { resultado: 'sem_conserto', custo: null, observacoes: 'Placa queimada' })
    expect(campos.sort()).toEqual(['observacoes', 'resultado'])
    expect(r.sugerir_baixa).toBe(true)
  })

  it('abrirManutencao envia o JSON', async () => {
    let corpo: unknown
    server.use(
      http.post('/api/manutencoes', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: 'Manutenção aberta.', id: 9 })
      })
    )
    await abrirManutencao({ item_id: 7, defeito: 'Tela quebrada' })
    expect(corpo).toEqual({ item_id: 7, defeito: 'Tela quebrada' })
  })

  it('baixarAnexoManutencao busca o anexo autenticado', async () => {
    let chamado = false
    server.use(
      http.get('/api/manutencoes/3/anexo', () => {
        chamado = true
        return new HttpResponse(new Blob(['%PDF']), { headers: { 'Content-Type': 'application/pdf' } })
      })
    )
    URL.createObjectURL = () => 'blob:teste'
    URL.revokeObjectURL = () => {}
    await baixarAnexoManutencao({ id: 3, anexo_key: 'manutencoes/manutencao_3_x.pdf' })
    expect(chamado).toBe(true)
  })
})
