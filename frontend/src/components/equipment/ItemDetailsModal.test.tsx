import { describe, expect, it, vi } from 'vitest'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { http, HttpResponse } from 'msw'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ItemDetailsModal } from '@/components/equipment/ItemDetailsModal'
import { ToastContainer } from '@/components/ui/toast'
import type { Item } from '@/api/items'

// A ficha técnica não depende do papel do usuário logado — mocka-se apenas
// o suficiente para satisfazer o hook (evita renderizar sem AuthProvider).
vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'tecnico.teste', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

const desktopItem: Item = {
  id: 42,
  tipo: 'Desktop',
  brand: 'Dell',
  model: 'OptiPlex 7090',
  status: 'Indisponível',
  cpu: 'Intel Core i5-11500',
  ram: '16GB',
  host: 'TI-DESKTOP-042',
}

function renderModal(item: Item | null) {
  return renderWithClient(
    <MemoryRouter>
      <ItemDetailsModal item={item} onClose={vi.fn()} />
      <ToastContainer />
    </MemoryRouter>
  )
}

describe('ItemDetailsModal — seção Hardware & Rede', () => {
  // Regressão: o enum real de tipos do backend usa "Desktop", nunca o nome
  // popular do equipamento (ver backend/app/core/config.py EQUIPMENT_TYPES).
  // Uma condição que checasse pelo nome errado escondia a seção inteira para
  // todo desktop da frota — e, com a edição inline, também impediria editar
  // CPU, RAM, storage, host, IP e licença desses itens.
  it('exibe a seção Hardware & Rede para um item do tipo Desktop', () => {
    renderModal(desktopItem)

    expect(screen.getByText('Especificações de Hardware & Sistema')).toBeInTheDocument()
    expect(screen.getByText('Processador (CPU)')).toBeInTheDocument()
    expect(screen.getByText('Intel Core i5-11500')).toBeInTheDocument()
  })

  it('exibe a seção Hardware & Rede também para um item do tipo Notebook', () => {
    renderModal({ ...desktopItem, id: 43, tipo: 'Notebook' })

    expect(screen.getByText('Especificações de Hardware & Sistema')).toBeInTheDocument()
  })
})

describe('ItemDetailsModal — edição inline', () => {
  it('entra em modo de edição, altera a marca e salva as alterações', async () => {
    const capturado: { corpo: Record<string, unknown> | null } = { corpo: null }
    server.use(
      http.put('/api/items/:id', async ({ request }) => {
        capturado.corpo = (await request.json()) as Record<string, unknown>
        return HttpResponse.json({ detail: 'Item atualizado com sucesso.' })
      })
    )

    const user = userEvent.setup()
    renderModal(desktopItem)

    // Fora do modo de edição, os dados aparecem como texto — não como input.
    expect(screen.queryByLabelText('Marca')).not.toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: 'Editar' }))

    const campoMarca = screen.getByLabelText('Marca')
    await user.clear(campoMarca)
    await user.type(campoMarca, 'HP')

    const campoCpu = screen.getByLabelText('Processador (CPU)')
    await user.clear(campoCpu)
    await user.type(campoCpu, 'Intel Core i7-12700')

    await user.click(screen.getByRole('button', { name: /Salvar Alterações/ }))

    await waitFor(() => {
      expect(capturado.corpo).not.toBeNull()
    })
    expect(capturado.corpo?.brand).toBe('HP')
    expect(capturado.corpo?.cpu).toBe('Intel Core i7-12700')

    // Após salvar com sucesso, o modal sai do modo de edição.
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Editar' })).toBeInTheDocument()
    })
  })

  it('rejeita um endereço MAC inválido e não chama o backend', async () => {
    let chamouBackend = false
    server.use(
      http.put('/api/items/:id', () => {
        chamouBackend = true
        return HttpResponse.json({ detail: 'ok' })
      })
    )

    const user = userEvent.setup()
    renderModal(desktopItem)

    await user.click(screen.getByRole('button', { name: 'Editar' }))

    // "AB" passa pela máscara (são dígitos hex válidos), mas não forma um MAC
    // completo — isValidMac exige os 6 pares (AA:BB:CC:DD:EE:FF).
    const campoMac = screen.getByLabelText('Endereço MAC / Físico')
    await user.type(campoMac, 'AB')
    await user.click(screen.getByRole('button', { name: /Salvar Alterações/ }))

    expect(screen.getByText(/Endereço MAC inválido/)).toBeInTheDocument()
    expect(chamouBackend).toBe(false)
  })
})
