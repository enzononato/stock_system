import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import PeripheralsPage from '@/pages/PeripheralsPage'

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

const perifericos = Array.from({ length: 30 }, (_, i) => ({
  id: i + 1,
  tipo: 'Mouse',
  brand: 'Logitech',
  model: `M${i + 1}`,
  identificador: `SN-${i + 1}`,
  status: i === 0 ? 'Em Uso' : 'Disponível',
  link_id: i === 0 ? 33 : undefined,
  date_registered: '2026-01-10',
}))

function renderPage() {
  server.use(http.get('/api/peripherals', () => HttpResponse.json(perifericos)))
  return renderWithClient(
    <MemoryRouter>
      <PeripheralsPage />
      <ToastContainer />
    </MemoryRouter>
  )
}

describe('PeripheralsPage', () => {
  beforeEach(() => localStorage.clear())

  it('pagina no cliente com 25 por padrão', async () => {
    renderPage()
    const user = userEvent.setup()
    expect(await screen.findByText('Mostrando 1–25 de 30')).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'Próxima' }))
    expect(await screen.findByText('Mostrando 26–30 de 30')).toBeInTheDocument()
  })

  it('status em pílula', async () => {
    renderPage()
    const pilula = (await screen.findAllByText('Em Uso'))[0]
    expect(pilula).toHaveClass('font-bold')
    expect(pilula.style.color).toBe('var(--status-loaned)')
  })

  it('desvincular com duplo clique envia uma requisição só', async () => {
    let liberar!: () => void
    const pendente = new Promise<void>((r) => {
      liberar = r
    })
    let chamadas = 0
    server.use(
      http.delete('/api/peripherals/links/33', async () => {
        chamadas += 1
        await pendente
        return HttpResponse.json({ detail: 'ok' })
      })
    )
    renderPage()
    const user = userEvent.setup()
    await user.click(await screen.findByText('#1'))
    const botao = await screen.findByRole('button', { name: 'Desvincular do Ativo' })
    await user.click(botao)
    await user.click(botao)
    await waitFor(() => expect(chamadas).toBe(1))
    expect(botao).toHaveAttribute('aria-busy', 'true')
    liberar()
  })
})
