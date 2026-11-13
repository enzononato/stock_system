import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import App from '@/App'
import { ThemeProvider } from '@/lib/theme'
import EtiquetaRedirectPage from '@/pages/EtiquetaRedirectPage'
import { PosLoginRedirect } from '@/features/etiquetas/PosLoginRedirect'
import { consumirDestinoPosLogin, guardarDestinoPosLogin } from '@/features/etiquetas/posLogin'

let user: { id: number; username: string; role: string } | null = null
let isLoading = false
vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user,
    isLoading,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => !!user && roles.includes(user.role),
  }),
}))

beforeEach(() => {
  user = null
  isLoading = false
  sessionStorage.clear()
})

function rotas(entrada: string) {
  return renderWithClient(
    <MemoryRouter initialEntries={[entrada]}>
      <PosLoginRedirect />
      <Routes>
        <Route path="/e/:id" element={<EtiquetaRedirectPage />} />
        <Route path="/login" element={<p>tela de login</p>} />
        <Route path="/items/:id" element={<p>detalhe</p>} />
        <Route path="/stock" element={<p>estoque</p>} />
        <Route path="/" element={<p>início</p>} />
      </Routes>
    </MemoryRouter>
  )
}

describe('EtiquetaRedirectPage (/e/:id)', () => {
  it('logado vai direto ao detalhe do equipamento', async () => {
    user = { id: 1, username: 'teste', role: 'Técnico' }
    rotas('/e/42')
    expect(await screen.findByText('detalhe')).toBeInTheDocument()
  })

  it('deslogado vai ao login e guarda o destino para depois', async () => {
    rotas('/e/42')

    expect(await screen.findByText('tela de login')).toBeInTheDocument()
    expect(consumirDestinoPosLogin()).toBe('/items/42')
  })

  it('id que não é número volta ao início', async () => {
    user = { id: 1, username: 'teste', role: 'Técnico' }
    rotas('/e/abc')
    expect(await screen.findByText('início')).toBeInTheDocument()
  })

  it('enquanto a sessão carrega, mostra o carregamento e não redireciona', () => {
    isLoading = true
    rotas('/e/42')
    expect(screen.getByText('Abrindo equipamento...')).toBeInTheDocument()
    expect(consumirDestinoPosLogin()).toBeNull()
  })
})

describe('PosLoginRedirect', () => {
  it('logado em "/" com destino guardado leva ao equipamento', async () => {
    user = { id: 1, username: 'teste', role: 'Técnico' }
    guardarDestinoPosLogin('/items/42')
    rotas('/')
    expect(await screen.findByText('detalhe')).toBeInTheDocument()
  })

  it('logado em outra rota não navega nem consome o destino', async () => {
    user = { id: 1, username: 'teste', role: 'Técnico' }
    guardarDestinoPosLogin('/items/42')
    rotas('/stock')
    expect(await screen.findByText('estoque')).toBeInTheDocument()
    expect(consumirDestinoPosLogin()).toBe('/items/42')
  })

  it('deslogado não consome o destino', async () => {
    guardarDestinoPosLogin('/items/42')
    rotas('/')
    expect(await screen.findByText('início')).toBeInTheDocument()
    expect(consumirDestinoPosLogin()).toBe('/items/42')
  })
})

describe('App com /e/:id', () => {
  it('logado, abre o equipamento pelo atalho do QR', async () => {
    user = { id: 1, username: 'teste', role: 'Técnico' }
    let chamado = false
    server.use(
      http.get('/api/items/42', () => {
        chamado = true
        return HttpResponse.json({ id: 42, tipo: 'Notebook', brand: 'Dell', model: 'Latitude', status: 'Disponível' })
      })
    )

    renderWithClient(
      <ThemeProvider>
        <MemoryRouter initialEntries={['/e/42']}>
          <App />
        </MemoryRouter>
      </ThemeProvider>
    )

    await waitFor(() => expect(chamado).toBe(true))
  })
})
