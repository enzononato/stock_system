import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, within } from '@testing-library/react'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import ItemDetailPage from '@/pages/ItemDetailPage'
import tab from '@/features/item-detail/tabs/linha'
import type { Item } from '@/api/items'
import type { Linha, LinhaEvento } from '@/api/linhas'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'tecnico.teste', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

const celular: Item = {
  id: 7,
  tipo: 'Celular',
  brand: 'Samsung',
  model: 'A15',
  identificador: 'CEL-A',
  status: 'Em uso',
  revenda: 'Juazeiro',
}

const linha: Linha = {
  id: 5,
  numero: '74999990001',
  operadora: 'Vivo',
  plano: 'Controle 20GB',
  custo_mensal: '49.90',
  status: 'ativa',
  unidade_id: 1,
  unidade_nome: 'Unidade A',
  item_id: 7,
  item_tipo: 'Celular',
  item_brand: 'Samsung',
  item_model: 'A15',
  item_identificador: 'CEL-A',
  item_ativo: true,
  colaborador_id: 3,
  colaborador_nome: 'Maria Souza',
  colaborador_situacao: 'ativo',
  iccid: null,
  observacoes: null,
  criada_em: '2026-11-12T10:00:00',
  atualizada_em: '2026-11-12T10:00:00',
}

const evento: LinhaEvento = {
  id: 1,
  linha_id: 5,
  linha_numero: '74999990001',
  tipo: 'troca_aparelho',
  campo: null,
  de_valor: null,
  para_valor: 'Celular Samsung A15 · CEL-A',
  item_id: 7,
  item_id_anterior: null,
  ocorrido_em: '2026-11-12T10:00:00',
  por_usuario: 'gestor.teste',
  observacao: null,
}

function Aba({ item }: { item: Item }) {
  const Component = tab.Component
  return (
    <MemoryRouter>
      <Component item={item} />
    </MemoryRouter>
  )
}

describe('aba Linha do equipamento', () => {
  it('é a aba "linha" de ordem 50, sem restrição de papel', () => {
    expect(tab.id).toBe('linha')
    expect(tab.label).toBe('Linha')
    expect(tab.order).toBe(50)
    expect(tab.roles).toBeUndefined()
  })

  it('equipamento que não é celular nem tablet mostra o aviso e não faz requisição', async () => {
    // Sem handlers: qualquer requisição faria o MSW (onUnhandledRequest: 'error') falhar o teste.
    renderWithClient(<Aba item={{ ...celular, tipo: 'Notebook' }} />)

    expect(screen.getByText('Linhas telefônicas só se aplicam a celulares e tablets.')).toBeInTheDocument()
    await new Promise((r) => setTimeout(r, 50))
  })

  it('celular com linha mostra o cartão da linha e o histórico de linhas do aparelho', async () => {
    let params: URLSearchParams | null = null
    server.use(
      http.get('/api/linhas', ({ request }) => {
        params = new URL(request.url).searchParams
        return HttpResponse.json({ items: [linha], total: 1 })
      }),
      http.get('/api/linhas/eventos', () => HttpResponse.json([evento]))
    )
    renderWithClient(<Aba item={celular} />)

    expect(await screen.findByRole('heading', { name: '(74) 99999-0001' })).toBeInTheDocument()
    expect(screen.getByText('Ativa')).toBeInTheDocument()
    expect(screen.getByText('UNIDADE A')).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Abrir em Linhas' })).toHaveAttribute('href', '/linhas?busca=74999990001')
    expect(screen.getByText('Histórico de linhas neste aparelho')).toBeInTheDocument()
    expect(await screen.findByText('Troca de aparelho')).toBeInTheDocument()
    expect(params!.get('item_id')).toBe('7')
    expect(params!.get('limit')).toBe('25')
  })

  it('celular sem linha mostra a mensagem de vazio', async () => {
    server.use(
      http.get('/api/linhas', () => HttpResponse.json({ items: [], total: 0 })),
      http.get('/api/linhas/eventos', () => HttpResponse.json([]))
    )
    renderWithClient(<Aba item={{ ...celular, tipo: 'Tablet' }} />)

    expect(await screen.findByText('Nenhuma linha vinculada a este aparelho.')).toBeInTheDocument()
    expect(await screen.findByText('Nenhum evento registrado.')).toBeInTheDocument()
  })

  it('no detalhe do equipamento a aba Linha fica depois de Dados e antes de Histórico', async () => {
    server.use(http.get('/api/items/7', () => HttpResponse.json(celular)))
    renderWithClient(
      <MemoryRouter initialEntries={['/items/7']}>
        <Routes>
          <Route path="/items/:id" element={<ItemDetailPage />} />
        </Routes>
      </MemoryRouter>
    )

    const abas = within(await screen.findByRole('tablist')).getAllByRole('tab').map((t) => t.textContent)
    expect(abas).toContain('Linha')
    expect(abas.indexOf('Dados')).toBeLessThan(abas.indexOf('Linha'))
    expect(abas.indexOf('Linha')).toBeLessThan(abas.indexOf('Histórico'))
  })
})
