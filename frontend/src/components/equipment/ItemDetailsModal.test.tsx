import { describe, expect, it, vi } from 'vitest'
import { screen } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import { renderWithClient } from '@/test/render'
import { ItemDetailsModal } from '@/components/equipment/ItemDetailsModal'
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
    </MemoryRouter>
  )
}

describe('ItemDetailsModal — seção Hardware & Rede', () => {
  // Regressão: o enum real de tipos do backend usa "Desktop", nunca
  // "Computador" (ver backend/app/core/config.py EQUIPMENT_TYPES). A
  // condição antiga (`['Computador', 'Notebook'].includes(item.tipo)`)
  // escondia a seção inteira para todo desktop da frota.
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
