import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import ReportPage from '@/pages/ReportPage'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'gestor.teste', role: 'Gestor' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Gestor'),
  }),
}))

// jsdom não implementa a Pointer Events API que o Select (Radix) usa.
beforeAll(() => {
  Element.prototype.hasPointerCapture = Element.prototype.hasPointerCapture ?? (() => false)
  Element.prototype.setPointerCapture = Element.prototype.setPointerCapture ?? (() => {})
  Element.prototype.releasePointerCapture = Element.prototype.releasePointerCapture ?? (() => {})
  Element.prototype.scrollIntoView = Element.prototype.scrollIntoView ?? (() => {})
})

const linhas = Array.from({ length: 30 }, (_, i) => ({
  history_id: i + 1,
  item_id: i + 1,
  operador: 'gestor.teste',
  operation_type: 'Cadastro',
  tipo: 'Notebook',
  revenda: i < 20 ? 'Revalle Juazeiro' : 'Revalle Petrolina',
  data_emprestimo: '2026-01-10T09:00:00',
}))

function renderPage() {
  server.use(
    http.get('/api/unidades', () => HttpResponse.json([])),
    http.get('/api/reports/monthly', () => HttpResponse.json(linhas))
  )
  return renderWithClient(<ReportPage />)
}

describe('ReportPage', () => {
  beforeEach(() => localStorage.clear())

  it('pagina no cliente com 25 por padrão', async () => {
    renderPage()
    expect(await screen.findByText('Mostrando 1–25 de 30')).toBeInTheDocument()
  })

  it('revenda em CAIXA ALTA e filtro pelo nome original', async () => {
    renderPage()
    const user = userEvent.setup()
    expect((await screen.findAllByText('REVALLE JUAZEIRO')).length).toBeGreaterThan(0)
    await user.click(screen.getByRole('combobox', { name: 'Revenda' }))
    await user.click(await screen.findByRole('option', { name: 'REVALLE PETROLINA' }))
    expect(await screen.findByText('Mostrando 1–10 de 10')).toBeInTheDocument()
  })

  it('indicadores em caixas azuis', async () => {
    renderPage()
    const rotulo = await screen.findByText('Total de Registros')
    expect(rotulo.closest('div')).toHaveClass('bg-brand-soft')
  })
})
