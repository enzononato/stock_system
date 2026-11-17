import { beforeAll, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ThemeProvider } from '@/lib/theme'
import ChartsPage from '@/pages/ChartsPage'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'gestor.teste', role: 'Gestor' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Gestor'),
  }),
}))

beforeAll(() => {
  Element.prototype.hasPointerCapture = Element.prototype.hasPointerCapture ?? (() => false)
  Element.prototype.setPointerCapture = Element.prototype.setPointerCapture ?? (() => {})
  Element.prototype.releasePointerCapture = Element.prototype.releasePointerCapture ?? (() => {})
  Element.prototype.scrollIntoView = Element.prototype.scrollIntoView ?? (() => {})
})

const dias = Array.from({ length: 30 }, (_, i) => i + 1)
const zeros = dias.map(() => 0)

function renderPage() {
  server.use(
    http.get('/api/reports/charts/loans', () => HttpResponse.json({ days: dias, values: zeros, values2: zeros })),
    http.get('/api/reports/charts/registrations', () => HttpResponse.json({ days: dias, values: zeros })),
    http.get('/api/unidades', () =>
      HttpResponse.json([{ id: 1, nome: 'Revalle Juazeiro', razao_social: 'X', cnpj: '1', is_active: true }])
    )
  )
  return renderWithClient(
    <ThemeProvider>
      <ChartsPage />
    </ThemeProvider>
  )
}

describe('ChartsPage', () => {
  it('unidade em CAIXA ALTA no seletor', async () => {
    renderPage()
    const user = userEvent.setup()
    await user.click(screen.getByRole('combobox', { name: 'Unidade' }))
    expect(await screen.findByRole('option', { name: 'REVALLE JUAZEIRO' })).toBeInTheDocument()
  })

  it('indicadores em caixas azuis', async () => {
    renderPage()
    const rotulo = await screen.findByText('Empréstimos no mês')
    expect(rotulo.closest('div')).toHaveClass('bg-brand-soft')
  })
})
