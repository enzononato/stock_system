import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import { getVisibleNavItems } from '@/components/layout/navigation'
import SistemasPage from './SistemasPage'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'gestor', role: 'Gestor' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Gestor'),
  }),
}))

const sistemas = [
  { id: 1, nome: 'AD', descricao: 'Active Directory', ativo: true },
  { id: 2, nome: 'Promax', descricao: null, ativo: false },
]

function renderPagina() {
  return renderWithClient(
    <MemoryRouter>
      <SistemasPage />
      <ToastContainer />
    </MemoryRouter>
  )
}

beforeEach(() => {
  server.use(http.get('/api/acessos/sistemas', () => HttpResponse.json(sistemas)))
})

describe('SistemasPage', () => {
  it('lista os sistemas com a situação de cada um (inclusive desativados)', async () => {
    let query = ''
    server.use(
      http.get('/api/acessos/sistemas', ({ request }) => {
        query = new URL(request.url).search
        return HttpResponse.json(sistemas)
      })
    )
    renderPagina()
    const linhaAd = (await screen.findByText('AD')).closest('tr') as HTMLElement
    expect(within(linhaAd).getByText('Ativo')).toBeInTheDocument()
    const linhaPromax = screen.getByText('Promax').closest('tr') as HTMLElement
    expect(within(linhaPromax).getByText('Desativado')).toBeInTheDocument()
    expect(query).toBe('?include_inactive=true')
  })

  it('cadastra um sistema novo', async () => {
    let corpo: unknown = null
    server.use(
      http.post('/api/acessos/sistemas', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ id: 3, nome: 'GLPI', descricao: null, ativo: true }, { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('AD')
    await user.click(screen.getByRole('button', { name: 'Novo sistema' }))
    await user.type(screen.getByLabelText(/^Nome/), 'GLPI')
    await user.click(screen.getByRole('button', { name: 'Salvar' }))
    await waitFor(() => expect(corpo).toEqual({ nome: 'GLPI', descricao: null }))
    expect(await screen.findByText('Sistema cadastrado.')).toBeInTheDocument()
  })

  it('mostra o erro da API quando o nome já existe', async () => {
    server.use(
      http.post('/api/acessos/sistemas', () =>
        HttpResponse.json({ detail: 'Já existe um sistema com esse nome.' }, { status: 409 })
      )
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('AD')
    await user.click(screen.getByRole('button', { name: 'Novo sistema' }))
    await user.type(screen.getByLabelText(/^Nome/), 'ad')
    await user.click(screen.getByRole('button', { name: 'Salvar' }))
    expect(await screen.findByText('Já existe um sistema com esse nome.')).toBeInTheDocument()
  })

  it('desativa um sistema ativo', async () => {
    let corpo: unknown = null
    server.use(
      http.put('/api/acessos/sistemas/1', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...sistemas[0], ativo: false })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('AD')
    await user.click(screen.getByRole('button', { name: 'Desativar AD' }))
    await waitFor(() => expect(corpo).toEqual({ ativo: false }))
    expect(await screen.findByText('Sistema desativado.')).toBeInTheDocument()
  })

  it('reativa um sistema desativado', async () => {
    let corpo: unknown = null
    server.use(
      http.put('/api/acessos/sistemas/2', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ ...sistemas[1], ativo: true })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('Promax')
    await user.click(screen.getByRole('button', { name: 'Reativar Promax' }))
    await waitFor(() => expect(corpo).toEqual({ ativo: true }))
    expect(await screen.findByText('Sistema reativado.')).toBeInTheDocument()
  })

  it('edita nome e descrição', async () => {
    let corpo: unknown = null
    server.use(
      http.put('/api/acessos/sistemas/1', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(sistemas[0])
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await screen.findByText('AD')
    await user.click(screen.getByRole('button', { name: 'Editar AD' }))
    const nome = screen.getByLabelText(/^Nome/)
    await user.clear(nome)
    await user.type(nome, 'AD Corp')
    await user.click(screen.getByRole('button', { name: 'Salvar' }))
    await waitFor(() => expect(corpo).toEqual({ nome: 'AD Corp', descricao: 'Active Directory' }))
    expect(await screen.findByText('Sistema atualizado.')).toBeInTheDocument()
  })

  it('o menu mostra Sistemas só para o Gestor', () => {
    expect(getVisibleNavItems('Técnico').map((i) => i.to)).not.toContain('/acessos/sistemas')
    expect(getVisibleNavItems('Gestor').map((i) => i.to)).toContain('/acessos/sistemas')
  })
})
