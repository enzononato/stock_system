import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { AparelhoSelect } from '@/components/linhas/AparelhoSelect'
import type { AparelhoOpcao } from '@/api/linhas'

const celularA: AparelhoOpcao = {
  id: 7,
  tipo: 'Celular',
  brand: 'Samsung',
  model: 'A15',
  identificador: 'CEL-A',
  codigo_patrimonial: 'PAT-1',
  revenda: 'Juazeiro',
  status: 'Disponível',
  responsavel_nome: 'Maria Souza',
}

describe('AparelhoSelect', () => {
  it('busca no servidor com debounce, limit 20 e linha_id, e devolve o aparelho escolhido', async () => {
    const chamadas: URLSearchParams[] = []
    server.use(
      http.get('/api/linhas/aparelhos', ({ request }) => {
        const params = new URL(request.url).searchParams
        chamadas.push(params)
        return HttpResponse.json(params.get('busca') === 'A15' ? [celularA] : [])
      })
    )
    const onChange = vi.fn()
    renderWithClient(<AparelhoSelect value={null} onChange={onChange} linhaId={9} />)
    const user = userEvent.setup()

    await user.click(screen.getByRole('button', { name: /selecione/i }))
    await user.type(screen.getByPlaceholderText(/buscar/i), 'A15')

    const opcao = await screen.findByText('Celular Samsung A15 · CEL-A')
    expect(screen.getByText('JUAZEIRO · MARIA SOUZA · PAT-1')).toBeInTheDocument()
    const comBusca = chamadas.filter((p) => p.get('busca') === 'A15')
    expect(comBusca.length).toBeGreaterThan(0)
    expect(comBusca[0].get('limit')).toBe('20')
    expect(comBusca[0].get('linha_id')).toBe('9')
    expect(chamadas.some((p) => p.get('busca') === 'A' || p.get('busca') === 'A1')).toBe(false)

    await user.click(opcao)
    expect(onChange).toHaveBeenCalledWith(celularA)
  })

  it('sem linhaId não envia linha_id e mostra o texto de lista vazia', async () => {
    const chamadas: URLSearchParams[] = []
    server.use(
      http.get('/api/linhas/aparelhos', ({ request }) => {
        chamadas.push(new URL(request.url).searchParams)
        return HttpResponse.json([])
      })
    )
    renderWithClient(<AparelhoSelect value={null} onChange={vi.fn()} />)
    const user = userEvent.setup()

    await user.click(screen.getByRole('button', { name: /selecione/i }))

    expect(await screen.findByText('Nenhum celular ou tablet livre encontrado.')).toBeInTheDocument()
    expect(chamadas[0].has('linha_id')).toBe(false)
  })

  it('busca o item quando o valor não está nos resultados e mostra o rótulo dele', async () => {
    server.use(
      http.get('/api/linhas/aparelhos', () => HttpResponse.json([])),
      http.get('/api/items/7', () =>
        HttpResponse.json({
          id: 7,
          tipo: 'Celular',
          brand: 'Samsung',
          model: 'A15',
          identificador: 'CEL-A',
          status: 'Em uso',
        })
      )
    )
    renderWithClient(<AparelhoSelect value={7} onChange={vi.fn()} />)

    expect(await screen.findByText('Celular Samsung A15 · CEL-A')).toBeInTheDocument()
  })

  it('o botão Limpar aparelho devolve null', async () => {
    server.use(
      http.get('/api/linhas/aparelhos', () => HttpResponse.json([celularA]))
    )
    const onChange = vi.fn()
    renderWithClient(<AparelhoSelect value={7} onChange={onChange} />)
    const user = userEvent.setup()

    await user.click(await screen.findByRole('button', { name: 'Limpar aparelho' }))

    expect(onChange).toHaveBeenCalledWith(null)
  })
})
