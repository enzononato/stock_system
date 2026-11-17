import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import App from '@/App'
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

function renderRota() {
  return renderWithClient(
    <ThemeProvider>
      <MemoryRouter initialEntries={['/locais']}>
        <App />
      </MemoryRouter>
    </ThemeProvider>
  )
}

beforeEach(() => {
  role = 'Técnico'
  server.use(http.get('/api/unidades', () => HttpResponse.json(unidades)))
})

describe('LocaisPage (/locais)', () => {
  it('Técnico abre a tela, escolhe uma unidade e cria um local', async () => {
    let corpo: unknown = null
    server.use(
      http.get('/api/locais', ({ request }) =>
        HttpResponse.json(
          new URL(request.url).searchParams.get('unidade_id') === '2'
            ? [{ id: 20, unidade_id: 2, nome: 'Depósito', descricao: null, ativo: true }]
            : []
        )
      ),
      http.post('/api/locais', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ id: 21 }, { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderRota()

    expect(await screen.findByRole('heading', { name: 'Locais' })).toBeInTheDocument()
    expect(screen.getByText('Selecione uma unidade')).toBeInTheDocument()

    const seletor = screen.getByLabelText('Unidade')
    expect(await screen.findByRole('option', { name: 'PETROLINA' })).toBeInTheDocument()
    await user.selectOptions(seletor, 'PETROLINA')

    expect(await screen.findByText('Depósito')).toBeInTheDocument()
    await user.type(screen.getByLabelText('Nome do novo local'), 'Recepção')
    await user.click(screen.getByRole('button', { name: 'Adicionar local' }))

    await waitFor(() => expect(corpo).toEqual({ unidade_id: 2, nome: 'Recepção' }))
    expect(await screen.findByText('Local criado com sucesso!')).toBeInTheDocument()
  })

  it('Jovem Aprendiz vê acesso restrito', async () => {
    role = 'Jovem Aprendiz'
    renderRota()
    expect(await screen.findByText('Acesso restrito')).toBeInTheDocument()
    expect(screen.queryByLabelText('Unidade')).not.toBeInTheDocument()
  })
})
