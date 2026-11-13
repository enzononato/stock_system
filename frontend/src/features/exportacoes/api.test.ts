import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import { getErrorMessage } from '@/lib/api-error'
import {
  exportar,
  exportarEstoque,
  exportarHistorico,
  exportarLinhas,
  exportarManutencoes,
  exportarRelatorioMensal,
  exportarTermosPendentes,
  montarQuery,
} from './api'

const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

beforeEach(() => {
  URL.createObjectURL = vi.fn(() => 'blob:teste')
  URL.revokeObjectURL = vi.fn()
  // O jsdom não implementa navegação: o clique no link de download não deve tentar abrir o blob.
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
})

afterEach(() => {
  vi.restoreAllMocks()
})

function capturar(caminho: string) {
  const urls: URL[] = []
  server.use(
    http.get(`/api/exportacoes/${caminho}`, ({ request }) => {
      urls.push(new URL(request.url))
      return new HttpResponse(new Blob(['x']), { headers: { 'Content-Type': XLSX } })
    })
  )
  return urls
}

describe('montarQuery', () => {
  it('pula undefined, null e vazio, mas mantém zero', () => {
    expect(montarQuery({ a: 'x', b: undefined, c: '', d: 0, e: null })).toBe('a=x&d=0')
  })
})

describe('exportar', () => {
  it('exportarEstoque envia os filtros da tela e o formato', async () => {
    const urls = capturar('estoque')
    await exportarEstoque({ status: 'Disponível', sort: 'colaborador', dir: 'asc' }, 'xlsx')
    expect(urls).toHaveLength(1)
    const p = urls[0].searchParams
    expect(p.get('status')).toBe('Disponível')
    expect(p.get('sort')).toBe('colaborador')
    expect(p.get('dir')).toBe('asc')
    expect(p.get('formato')).toBe('xlsx')
    expect(p.has('tipo')).toBe(false)
    expect(URL.createObjectURL).toHaveBeenCalledTimes(1)
  })

  it('cada wrapper usa a rota e os parâmetros certos', async () => {
    const hist = capturar('historico')
    await exportarHistorico({ search: 'SN-1' }, 'pdf')
    expect(hist[0].search).toBe('?search=SN-1&formato=pdf')

    const man = capturar('manutencoes')
    await exportarManutencoes({ status: 'em_andamento', item_id: 7 }, 'xlsx')
    expect(man[0].searchParams.get('status')).toBe('em_andamento')
    expect(man[0].searchParams.get('item_id')).toBe('7')
    expect(man[0].searchParams.has('limit')).toBe(false)

    const lin = capturar('linhas')
    await exportarLinhas({ status: 'ativa', unidade_id: 2 }, 'xlsx')
    expect(lin[0].searchParams.get('status')).toBe('ativa')
    expect(lin[0].searchParams.get('unidade_id')).toBe('2')

    const ter = capturar('termos-pendentes')
    await exportarTermosPendentes({}, 'xlsx')
    expect(ter[0].search).toBe('?formato=xlsx')

    const rel = capturar('relatorio-mensal')
    await exportarRelatorioMensal(2026, 3, 'pdf')
    expect(rel[0].searchParams.get('year')).toBe('2026')
    expect(rel[0].searchParams.get('month')).toBe('3')
    expect(rel[0].searchParams.get('formato')).toBe('pdf')
  })

  it('o nome do arquivo baixado leva o nome base e o formato', async () => {
    capturar('relatorio-mensal')
    const links: HTMLAnchorElement[] = []
    const original = document.body.appendChild.bind(document.body)
    vi.spyOn(document.body, 'appendChild').mockImplementation(((el: Node) => {
      if (el instanceof HTMLAnchorElement) links.push(el)
      return original(el)
    }) as typeof document.body.appendChild)
    await exportarRelatorioMensal(2026, 3, 'xlsx')
    expect(links[0].download).toBe('relatorio_mensal_2026_03.xlsx')
  })

  it('erro 400 com corpo em Blob mostra o detail do servidor', async () => {
    const detail =
      'A exportação em PDF está limitada a 5.000 linhas e esta consulta tem 7.000. Refine os filtros ou exporte em Excel.'
    server.use(http.get('/api/exportacoes/estoque', () => HttpResponse.json({ detail }, { status: 400 })))
    const erro = await exportar('estoque', { sort: 'id' }, 'estoque', 'pdf').then(
      () => null,
      (e: unknown) => e
    )
    expect(erro).not.toBeNull()
    expect(getErrorMessage(erro)).toBe(detail)
  })

  it('erro sem JSON legível não vaza HTML', async () => {
    server.use(
      http.get('/api/exportacoes/estoque', () => new HttpResponse('<html>erro</html>', { status: 500 }))
    )
    const erro = await exportar('estoque', {}, 'estoque', 'pdf').then(
      () => null,
      (e: unknown) => e
    )
    expect(erro).not.toBeNull()
    expect(getErrorMessage(erro, 'Falhou')).not.toContain('<html>')
  })
})
