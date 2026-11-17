import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import LinhasRelatorioPage from '@/pages/LinhasRelatorioPage'
import { formatMoeda } from '@/lib/format'
import type { RelatorioCusto } from '@/api/linhas'

const relatorio: RelatorioCusto = {
  unidades: [
    { unidade_id: 1, unidade_nome: 'Unidade A', quantidade: 3, custo_total: '80.05' },
    { unidade_id: 2, unidade_nome: 'Unidade B', quantidade: 1, custo_total: '15.55' },
  ],
  total_linhas: 4,
  custo_total: '95.60',
  gerado_em: '2026-11-12T10:00:00',
}

// O Intl separa "R$" do valor com espaço não separável; o getByText normaliza o texto do nó para espaço comum.
const moeda = (v: string) => formatMoeda(v).replace(/\s/g, ' ')

function mockRelatorio(corpo: RelatorioCusto = relatorio) {
  server.use(http.get('/api/linhas/relatorio/custo', () => HttpResponse.json(corpo)))
}

function renderPage() {
  return renderWithClient(
    <MemoryRouter>
      <LinhasRelatorioPage />
      <ToastContainer />
    </MemoryRouter>
  )
}

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
  vi.restoreAllMocks()
})

describe('LinhasRelatorioPage', () => {
  it('mostra o custo por unidade e o total', async () => {
    mockRelatorio()
    renderPage()

    const tabela = within(await screen.findByRole('table'))
    expect(await tabela.findByText('UNIDADE A')).toBeInTheDocument()
    expect(tabela.getByText(moeda('80.05'))).toBeInTheDocument()
    expect(tabela.getByText('UNIDADE B')).toBeInTheDocument()
    expect(tabela.getByText(moeda('15.55'))).toBeInTheDocument()
    expect(tabela.getByText('Total')).toBeInTheDocument()
    expect(tabela.getByText(moeda('95.60'))).toBeInTheDocument()
    expect(tabela.getByText('4')).toBeInTheDocument()
    expect(screen.getByText(/^Gerado em /)).toBeInTheDocument()
    expect(screen.getByRole('link', { name: /Voltar para linhas/ })).toHaveAttribute('href', '/linhas')
    expect(screen.getByText('Somente linhas ativas, somadas por unidade.')).toBeInTheDocument()
  })

  it('Exportar PDF chama a rota com formato=pdf e o botão fica ocupado até a resposta', async () => {
    mockRelatorio()
    let liberar: () => void = () => {}
    const portao = new Promise<void>((resolve) => {
      liberar = resolve
    })
    let url = ''
    server.use(
      http.get('/api/linhas/relatorio/custo/export', async ({ request }) => {
        url = new URL(request.url).search
        await portao
        return HttpResponse.text('%PDF', { headers: { 'Content-Type': 'application/pdf' } })
      })
    )
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
    renderPage()
    await screen.findByText('UNIDADE A')
    const user = userEvent.setup()

    await user.click(screen.getByRole('button', { name: 'Exportar PDF' }))

    await waitFor(() => expect(screen.getByRole('button', { name: 'Exportar PDF' })).toHaveAttribute('aria-busy', 'true'))
    expect(screen.getByRole('button', { name: 'Exportar Excel' })).not.toHaveAttribute('aria-busy')
    expect(url).toBe('?formato=pdf')
    liberar()
    await waitFor(() =>
      expect(screen.getByRole('button', { name: 'Exportar PDF' })).not.toHaveAttribute('aria-busy')
    )
  })

  it('Exportar Excel usa formato=xlsx', async () => {
    mockRelatorio()
    let url = ''
    server.use(
      http.get('/api/linhas/relatorio/custo/export', ({ request }) => {
        url = new URL(request.url).search
        return HttpResponse.text('xlsx')
      })
    )
    const clique = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})
    renderPage()
    await screen.findByText('UNIDADE A')

    await userEvent.setup().click(screen.getByRole('button', { name: 'Exportar Excel' }))

    await waitFor(() => expect(clique).toHaveBeenCalledTimes(1))
    expect(url).toBe('?formato=xlsx')
  })

  it('falha na exportação mostra o toast de erro', async () => {
    mockRelatorio()
    server.use(
      http.get('/api/linhas/relatorio/custo/export', () => new HttpResponse(null, { status: 500 }))
    )
    renderPage()
    await screen.findByText('UNIDADE A')

    await userEvent.setup().click(screen.getByRole('button', { name: 'Exportar PDF' }))

    expect(await screen.findByText('Não foi possível exportar o relatório.')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Exportar PDF' })).not.toHaveAttribute('aria-busy')
  })

  it('sem linhas ativas mostra o estado vazio', async () => {
    mockRelatorio({ unidades: [], total_linhas: 0, custo_total: '0.00', gerado_em: '2026-11-12T10:00:00' })
    renderPage()

    expect(await screen.findByText('Nenhuma linha ativa')).toBeInTheDocument()
    expect(screen.queryByRole('table')).not.toBeInTheDocument()
  })
})
