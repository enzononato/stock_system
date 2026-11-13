import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import App from '@/App'
import InventariosPage from '@/pages/InventariosPage'
import { ThemeProvider } from '@/lib/theme'

let role = 'Técnico'
vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'teste', role },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes(role),
  }),
}))

const unidades = [
  { id: 1, nome: 'Juazeiro', razao_social: 'X', cnpj: '1', is_active: true },
  { id: 2, nome: 'Petrolina', razao_social: 'Y', cnpj: '2', is_active: true },
]

const inventario = {
  id: 7,
  unidade_id: 1,
  unidade_nome: 'Juazeiro',
  status: 'aberto',
  iniciado_por: 'tecnico.teste',
  iniciado_em: '2026-11-10T10:00:00',
  encerrado_por: null,
  encerrado_em: null,
  observacao: null,
  motivo_cancelamento: null,
  total_leituras: 3,
}

let consultas: string[] = []

beforeEach(() => {
  role = 'Técnico'
  consultas = []
  localStorage.clear()
  server.use(
    http.get('/api/unidades', () => HttpResponse.json(unidades)),
    http.get('/api/inventarios', ({ request }) => {
      consultas.push(new URL(request.url).search)
      return HttpResponse.json({
        items: [inventario, { ...inventario, id: 6, status: 'encerrado', encerrado_em: '2026-11-09T18:00:00' }],
        total: 2,
      })
    })
  )
})

function renderPagina() {
  return renderWithClient(
    <MemoryRouter initialEntries={['/inventarios']}>
      <Routes>
        <Route path="/inventarios" element={<InventariosPage />} />
        <Route path="/inventarios/:id/leitura" element={<p>tela de leitura</p>} />
        <Route path="/inventarios/:id" element={<p>relatório</p>} />
      </Routes>
    </MemoryRouter>
  )
}

describe('InventariosPage', () => {
  it('lista com unidade em caixa alta e status', async () => {
    renderPagina()

    const tabela = await screen.findByRole('table')
    expect(within(tabela).getAllByText('JUAZEIRO')).toHaveLength(2)
    expect(within(tabela).getByText('Aberto')).toBeInTheDocument()
    expect(within(tabela).getByText('Encerrado')).toBeInTheDocument()
    expect(consultas[0]).toBe('?limit=25&offset=0')
    // "Ler" só para o inventário aberto; "Relatório" para todos.
    expect(within(tabela).getAllByRole('link', { name: 'Ler' })).toHaveLength(1)
    expect(within(tabela).getAllByRole('link', { name: 'Relatório' })).toHaveLength(2)
    expect(within(tabela).getByRole('link', { name: 'Ler' })).toHaveAttribute('href', '/inventarios/7/leitura')
  })

  it('filtro de status volta à página 1', async () => {
    const user = userEvent.setup()
    renderPagina()
    await screen.findByRole('table')

    await user.selectOptions(screen.getByLabelText('Status'), 'Aberto')

    await waitFor(() => expect(consultas.some((c) => c.includes('status=aberto') && c.includes('offset=0'))).toBe(true))
  })

  it('abrir inventário vai para a leitura', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/inventarios', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...inventario, id: 9, resumo: {} }, { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderPagina()

    await user.click(await screen.findByRole('button', { name: 'Novo inventário' }))
    const seletor = await screen.findByLabelText('Unidade do inventário')
    expect(screen.getByRole('button', { name: 'Abrir inventário' })).toBeDisabled()
    await user.selectOptions(seletor, 'PETROLINA')
    await user.click(screen.getByRole('button', { name: 'Abrir inventário' }))

    expect(await screen.findByText('tela de leitura')).toBeInTheDocument()
    expect(corpo).toEqual({ unidade_id: 2 })
  })

  it('manda a observação quando preenchida', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/inventarios', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...inventario, id: 9, resumo: {} }, { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderPagina()

    await user.click(await screen.findByRole('button', { name: 'Novo inventário' }))
    await user.selectOptions(await screen.findByLabelText('Unidade do inventário'), 'JUAZEIRO')
    await user.type(screen.getByLabelText('Observação'), 'Contagem anual')
    await user.click(screen.getByRole('button', { name: 'Abrir inventário' }))

    await screen.findByText('tela de leitura')
    expect(corpo).toEqual({ unidade_id: 1, observacao: 'Contagem anual' })
  })

  it('inventário já aberto mostra a mensagem da API', async () => {
    server.use(
      http.post('/api/inventarios', () =>
        HttpResponse.json({ detail: 'Já existe um inventário aberto para esta unidade (#7).' }, { status: 409 })
      )
    )
    const user = userEvent.setup()
    // O toast vive no App; aqui basta o ToastContainer.
    const { ToastContainer } = await import('@/components/ui/toast')
    renderWithClient(
      <MemoryRouter initialEntries={['/inventarios']}>
        <InventariosPage />
        <ToastContainer />
      </MemoryRouter>
    )

    await user.click(await screen.findByRole('button', { name: 'Novo inventário' }))
    await user.selectOptions(await screen.findByLabelText('Unidade do inventário'), 'JUAZEIRO')
    await user.click(screen.getByRole('button', { name: 'Abrir inventário' }))

    expect(await screen.findByText('Já existe um inventário aberto para esta unidade (#7).')).toBeInTheDocument()
  })

  it('Jovem Aprendiz não acessa', async () => {
    role = 'Jovem Aprendiz'
    renderWithClient(
      <ThemeProvider>
        <MemoryRouter initialEntries={['/inventarios']}>
          <App />
        </MemoryRouter>
      </ThemeProvider>
    )
    expect(await screen.findByText('Acesso restrito')).toBeInTheDocument()
  })
})
