import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, within } from '@testing-library/react'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import ItemDetailPage from '@/pages/ItemDetailPage'
import { getItemDetailTabs } from '@/features/item-detail/registry'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 2, username: 'tecnico.teste', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

const item = {
  id: 5,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude 5420',
  identificador: 'NB-0005',
  status: 'Indisponível',
  revenda: 'Revalle Juazeiro',
}

const acesso = {
  id: 1,
  colaborador_id: 7,
  colaborador_nome: 'Maria Silva',
  sistema_id: 1,
  sistema_nome: 'AD',
  login: 'maria.silva',
  status: 'ativo',
  item_id: 5,
  item_descricao: 'Notebook Dell Latitude 5420 · NB-0005',
  criado_por: 'tecnico.teste',
  criado_em: '2026-11-10T09:30:00',
  desativado_por: null,
  desativado_em: null,
}

function renderPage() {
  return renderWithClient(
    <MemoryRouter initialEntries={['/items/5?aba=acessos']}>
      <Routes>
        <Route path="/items/:id" element={<ItemDetailPage />} />
      </Routes>
    </MemoryRouter>
  )
}

describe('aba Acessos do equipamento', () => {
  it('lista os acessos instalados no equipamento, consultando por item_id', async () => {
    let query = ''
    server.use(
      http.get('/api/items/5', () => HttpResponse.json(item)),
      http.get('/api/acessos/ficha', ({ request }) => {
        query = new URL(request.url).search
        return HttpResponse.json([acesso])
      })
    )
    renderPage()
    const painel = within(await screen.findByRole('tabpanel'))
    expect(await painel.findByText('maria.silva')).toBeInTheDocument()
    expect(painel.getByText('AD')).toBeInTheDocument()
    expect(painel.getByRole('link', { name: 'MARIA SILVA' })).toHaveAttribute('href', '/colaboradores/7')
    expect(painel.getByText('Ativo')).toBeInTheDocument()
    expect(query).toBe('?item_id=5')

    const abas = within(screen.getByRole('tablist')).getAllByRole('tab').map((t) => t.textContent)
    expect(abas.indexOf('Acessos')).toBeGreaterThan(abas.indexOf('Dados'))
    expect(abas.indexOf('Acessos')).toBeLessThan(abas.indexOf('Histórico'))
  })

  it('sem acessos mostra a frase de vazio', async () => {
    server.use(
      http.get('/api/items/5', () => HttpResponse.json(item)),
      http.get('/api/acessos/ficha', () => HttpResponse.json([]))
    )
    renderPage()
    expect(await screen.findByText('Nenhum acesso vinculado a este equipamento.')).toBeInTheDocument()
  })

  it('Jovem Aprendiz não vê a aba', () => {
    expect(getItemDetailTabs('Jovem Aprendiz').map((t) => t.id)).not.toContain('acessos')
    expect(getItemDetailTabs('Gestor').map((t) => t.id)).toContain('acessos')
  })
})
