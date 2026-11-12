import { describe, expect, it, vi } from 'vitest'
import { screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { http, HttpResponse } from 'msw'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import ItemDetailPage from '@/pages/ItemDetailPage'
import { buildItemDetailTabs } from '@/features/item-detail/registry'
import type { ItemDetailTab } from '@/features/item-detail/types'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'tecnico.teste', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

const item = {
  id: 7,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude 5420',
  identificador: 'NB-0007',
  status: 'Em uso',
  revenda: 'Revalle Juazeiro',
  cpu: 'Intel i5',
  colaborador_id: 3,
  colaborador_nome: 'Maria Souza',
  colaborador_setor: 'Financeiro',
  colaborador_cargo: 'Analista',
  superior_nome: 'João Pereira',
  local_id: 2,
  local_nome: 'Sala 12',
}

function renderPage(path: string) {
  return renderWithClient(
    <MemoryRouter initialEntries={[path]}>
      <Routes>
        <Route path="/items/:id" element={<ItemDetailPage />} />
      </Routes>
    </MemoryRouter>
  )
}

describe('ItemDetailPage', () => {
  it('renderiza cabeçalho e as abas Dados e Histórico', async () => {
    server.use(http.get('/api/items/7', () => HttpResponse.json(item)))
    renderPage('/items/7')

    expect(await screen.findByRole('heading', { name: /Dell Latitude 5420/ })).toBeInTheDocument()
    expect(screen.getAllByText('REVALLE JUAZEIRO').length).toBeGreaterThan(0)
    expect(screen.getAllByText('NB-0007').length).toBeGreaterThan(0)

    const tablist = screen.getByRole('tablist')
    const tabs = within(tablist).getAllByRole('tab')
    expect(tabs.map((t) => t.textContent)).toEqual(['Dados', 'Histórico'])
    expect(screen.getByRole('tab', { name: 'Dados' })).toHaveAttribute('aria-selected', 'true')
    expect(screen.getByRole('tabpanel')).toBeInTheDocument()
  })

  it('mostra colaborador, superior imediato e local na aba Dados', async () => {
    server.use(http.get('/api/items/7', () => HttpResponse.json(item)))
    renderPage('/items/7')

    expect(await screen.findByText('Superior imediato')).toBeInTheDocument()
    expect(screen.getByText('João Pereira')).toBeInTheDocument()
    expect(screen.getByText('Maria Souza')).toBeInTheDocument()
    expect(screen.getByText('Financeiro')).toBeInTheDocument()
    expect(screen.getByText('Analista')).toBeInTheDocument()
    expect(screen.getByText('Sala 12')).toBeInTheDocument()
  })

  it('?aba=historico abre o histórico e consulta a API com item_id', async () => {
    server.use(http.get('/api/items/7', () => HttpResponse.json(item)))
    const itemIds: Array<string | null> = []
    server.use(
      http.get('/api/history', ({ request }) => {
        itemIds.push(new URL(request.url).searchParams.get('item_id'))
        return HttpResponse.json({
          items: [{ id: 1, item_id: 7, operation: 'Cadastro', operador: 'gestor', data_operacao: '2026-01-10T09:00:00' }],
          total: 1,
        })
      })
    )
    renderPage('/items/7?aba=historico')

    expect(await screen.findByText('Cadastro')).toBeInTheDocument()
    expect(screen.getByRole('tab', { name: 'Histórico' })).toHaveAttribute('aria-selected', 'true')
    expect(itemIds[0]).toBe('7')
  })

  it('trocar de aba atualiza a seleção', async () => {
    server.use(http.get('/api/items/7', () => HttpResponse.json(item)))
    renderPage('/items/7')
    await screen.findByRole('tablist')
    await userEvent.click(screen.getByRole('tab', { name: 'Histórico' }))
    await waitFor(() =>
      expect(screen.getByRole('tab', { name: 'Histórico' })).toHaveAttribute('aria-selected', 'true')
    )
  })

  it('item inexistente mostra "Equipamento não encontrado"', async () => {
    server.use(http.get('/api/items/999', () => HttpResponse.json({ detail: 'Item não encontrado' }, { status: 404 })))
    renderPage('/items/999')
    expect(await screen.findByText('Equipamento não encontrado')).toBeInTheDocument()
  })
})

describe('registry', () => {
  const Vazio = () => null
  const tab = (over: Partial<ItemDetailTab>): ItemDetailTab => ({
    id: 'x',
    label: 'X',
    order: 10,
    Component: Vazio,
    ...over,
  })

  it('filtra aba com roles: ["Gestor"] para Técnico e ordena por order', () => {
    const mods = {
      './tabs/a.tsx': { default: tab({ id: 'a', order: 30 }) },
      './tabs/b.tsx': { default: tab({ id: 'b', order: 20, roles: ['Gestor'] }) },
      './tabs/c.tsx': { default: tab({ id: 'c', order: 10 }) },
    }
    expect(buildItemDetailTabs(mods, 'Técnico').map((t) => t.id)).toEqual(['c', 'a'])
    expect(buildItemDetailTabs(mods, 'Gestor').map((t) => t.id)).toEqual(['c', 'b', 'a'])
    expect(buildItemDetailTabs(mods, undefined).map((t) => t.id)).toEqual(['c', 'a'])
  })
})
