import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import LinkPeripheralPage from '@/pages/LinkPeripheralPage'

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

const equipamento = {
  id: 5,
  tipo: 'Desktop',
  brand: 'Dell',
  model: 'OptiPlex',
  status: 'Indisponível',
  revenda: 'Revalle Juazeiro',
}

let liberar!: () => void
let chamadasDelete = 0

function mockServidor() {
  chamadasDelete = 0
  const pendente = new Promise<void>((r) => {
    liberar = r
  })
  server.use(
    http.get('/api/items', ({ request }) => {
      const tipo = new URL(request.url).searchParams.get('tipo')
      const items = tipo === 'Desktop' ? [equipamento] : []
      return HttpResponse.json({ items, total: items.length })
    }),
    http.get('/api/items/5/peripherals', () =>
      HttpResponse.json([{ id: 9, tipo: 'Mouse', status: 'Em Uso', link_id: 33 }])
    ),
    http.get('/api/peripherals', () => HttpResponse.json([])),
    http.delete('/api/peripherals/links/33', async () => {
      chamadasDelete += 1
      await pendente
      return HttpResponse.json({ detail: 'ok' })
    })
  )
}

describe('LinkPeripheralPage', () => {
  beforeEach(() => localStorage.clear())

  async function escolherEquipamento() {
    const user = userEvent.setup()
    renderWithClient(
      <MemoryRouter>
        <LinkPeripheralPage />
        <ToastContainer />
      </MemoryRouter>
    )
    await user.click(screen.getByRole('button', { name: /Selecione ou busque um equipamento/ }))
    await user.click(await screen.findByText('#5 — Desktop Dell OptiPlex'))
    return user
  }

  it('mostra a unidade em CAIXA ALTA e o status em pílula', async () => {
    mockServidor()
    await escolherEquipamento()
    expect(await screen.findByText('Em Uso')).toHaveClass('font-bold')
    expect(screen.getByText('REVALLE JUAZEIRO')).toBeInTheDocument()
    expect(screen.getByText('Indisponível')).toHaveClass('font-bold')
    liberar()
  })

  it('desvincular com duplo clique envia uma requisição só', async () => {
    mockServidor()
    const user = await escolherEquipamento()
    const botao = await screen.findByRole('button', { name: /Desvincular/ })
    await user.click(botao)
    await user.click(botao)
    await waitFor(() => expect(chamadasDelete).toBe(1))
    expect(botao).toHaveAttribute('aria-busy', 'true')
    liberar()
  })
})
