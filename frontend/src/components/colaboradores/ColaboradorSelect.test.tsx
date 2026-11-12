import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ColaboradorSelect } from '@/components/colaboradores/ColaboradorSelect'
import type { Colaborador } from '@/api/colaboradores'

const maria: Colaborador = {
  id: 7,
  cpf: '12345678909',
  nome: 'MARIA SILVA',
  cargo: 'Analista',
  unidade_nome: 'Juazeiro',
  situacao: 'ativo',
  email_manual: false,
  telefone_manual: false,
  origem: 'rh',
}

describe('ColaboradorSelect', () => {
  it('busca no servidor com o termo digitado (somente ativos) e devolve o colaborador escolhido', async () => {
    const chamadas: URLSearchParams[] = []
    server.use(
      http.get('/api/colaboradores', ({ request }) => {
        const params = new URL(request.url).searchParams
        chamadas.push(params)
        const busca = params.get('busca')
        return HttpResponse.json({ items: busca === 'mar' ? [maria] : [], total: busca === 'mar' ? 1 : 0 })
      })
    )
    const onChange = vi.fn()
    renderWithClient(<ColaboradorSelect value={null} onChange={onChange} />)
    const user = userEvent.setup()

    await user.click(screen.getByRole('button', { name: /selecione/i }))
    await user.type(screen.getByPlaceholderText(/buscar/i), 'mar')

    const opcao = await screen.findByText('MARIA SILVA')
    expect(screen.getByText('123.456.789-09 · Analista · JUAZEIRO')).toBeInTheDocument()

    const comBusca = chamadas.filter((p) => p.get('busca') === 'mar')
    expect(comBusca.length).toBeGreaterThan(0)
    expect(comBusca[0].get('situacao')).toBe('ativo')
    expect(comBusca[0].get('limit')).toBe('20')
    // debounce: nenhuma chamada intermediária com "m" ou "ma"
    expect(chamadas.some((p) => p.get('busca') === 'm' || p.get('busca') === 'ma')).toBe(false)

    await user.click(opcao)
    expect(onChange).toHaveBeenCalledWith(maria)
  })

  it('com somenteAtivos=false não filtra por situação', async () => {
    const chamadas: URLSearchParams[] = []
    server.use(
      http.get('/api/colaboradores', ({ request }) => {
        chamadas.push(new URL(request.url).searchParams)
        return HttpResponse.json({ items: [], total: 0 })
      })
    )
    renderWithClient(<ColaboradorSelect value={null} onChange={vi.fn()} somenteAtivos={false} />)
    const user = userEvent.setup()
    await user.click(screen.getByRole('button', { name: /selecione/i }))
    await waitFor(() => expect(chamadas.length).toBeGreaterThan(0))
    expect(chamadas[0].has('situacao')).toBe(false)
  })

  it('mostra o colaborador do value mesmo fora da lista carregada e permite limpar', async () => {
    server.use(
      http.get('/api/colaboradores', () => HttpResponse.json({ items: [], total: 0 })),
      http.get('/api/colaboradores/7', () => HttpResponse.json(maria))
    )
    const onChange = vi.fn()
    renderWithClient(<ColaboradorSelect value={7} onChange={onChange} />)
    expect(await screen.findByText('MARIA SILVA')).toBeInTheDocument()

    await userEvent.setup().click(screen.getByRole('button', { name: /limpar/i }))
    expect(onChange).toHaveBeenCalledWith(null)
  })
})
