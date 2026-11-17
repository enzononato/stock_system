import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import App from '@/App'
import { ThemeProvider } from '@/lib/theme'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 2, username: 'tecnico', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

function renderRota(caminho: string) {
  return renderWithClient(
    <ThemeProvider>
      <MemoryRouter initialEntries={[caminho]}>
        <App />
      </MemoryRouter>
    </ThemeProvider>
  )
}

describe('rotas do módulo de Acessos', () => {
  it('/acessos redireciona para a lista de eventos em /acessos/eventos', async () => {
    server.use(http.get('/api/acessos/eventos', () => HttpResponse.json({ items: [], total: 0 })))
    renderRota('/acessos')
    expect(await screen.findByRole('heading', { level: 1, name: 'Eventos de acesso' })).toBeInTheDocument()
  })

  it('/acessos/eventos abre a lista de eventos', async () => {
    server.use(http.get('/api/acessos/eventos', () => HttpResponse.json({ items: [], total: 0 })))
    renderRota('/acessos/eventos')
    expect(await screen.findByRole('heading', { level: 1, name: 'Eventos de acesso' })).toBeInTheDocument()
  })

  it('Técnico não entra nas telas só do Gestor', async () => {
    renderRota('/acessos/sistemas')
    expect(screen.queryByRole('heading', { level: 1, name: 'Sistemas' })).not.toBeInTheDocument()
  })
})
